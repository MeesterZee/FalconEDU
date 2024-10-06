<script type="text/javascript">
  // Global variables  
  let STUDENT_DATA;
  let SETTINGS;
  let DATA_SET = "active";
  let previousStudent;
  
  // Global flags
  let saveFlag = true; // True if all changes saved, false if unsaved changes
  let busyFlag = false; // True if operation in progress, false if no operation in progress

  // Initialize application
  // Conversion to async allows for parallel data retrieval from Apps Script
  window.onload = async function() {
    console.log("Initializing dashboard...");

    // Get main elements
    const toolbar = document.getElementById('toolbar');
    const page = document.getElementById('page');
    const loadingIndicator = document.getElementById('loading-indicator');

    try {
      // Show loading indicator and hide page elements
      loadingIndicator.style.display = 'block';
      toolbar.style.display = 'none';
      page.style.display = 'none';

      // Fetch data in parallel
      const [studentData, settings] = await Promise.all([
        new Promise((resolve, reject) => {
          if (DATA_SET === "archive") {
            google.script.run.withSuccessHandler(resolve).withFailureHandler(reject).getArchiveData();
          } else {
            google.script.run.withSuccessHandler(resolve).withFailureHandler(reject).getActiveData();
          }
        }),
        new Promise((resolve, reject) => {
          google.script.run.withSuccessHandler(resolve).withFailureHandler(reject).getSettings();
        })
      ]);

      // Assign data to global variables
      SETTINGS = settings;
      STUDENT_DATA = studentData;

      // Initialize the dashboard
      setEventListeners();
      populateDashboard();

      console.log("Initialization complete!");

    } catch (error) {
      console.error("Error during initialization:", error);
    
    } finally {
      // Hide loading indicator and show page elements
      loadingIndicator.style.display = 'none';
      toolbar.style.display = 'block';
      page.style.display = 'flex';
    }
  };
    
  function setEventListeners() {
    console.log("Setting event listeners...")
    
    // Check for unsaved changes or busy state before closing the window
    window.addEventListener('beforeunload', function (e) {
      if (!saveFlag || busyFlag) {
        e.preventDefault();
        e.returnValue = '';
      }
    });

    // Add event listeners for tool bar buttons
    document.getElementById('saveChangesButton').addEventListener('click', saveProfile);
    document.getElementById('addStudentButton').addEventListener('click', addStudent);
    document.getElementById('removeStudentButton').addEventListener('click', removeStudent);
    document.getElementById('renameStudentButton').addEventListener('click', renameStudent);
    document.getElementById('restoreStudentButton').addEventListener('click', restoreStudent);
    document.getElementById('deleteStudentButton').addEventListener('click', deleteStudent);
    document.getElementById('emailButton').addEventListener('click', composeEmail);
    document.getElementById('exportFormsButton').addEventListener('click', exportForms);
    document.getElementById('exportDataButton').addEventListener('click', exportData);
    document.getElementById('archiveButton').addEventListener('click', toggleDataView);
    document.getElementById('backButton').addEventListener('click', toggleDataView);

    // Dropdown event listeners
    document.querySelectorAll('.dropdown').forEach(dropdown => {
      const dropdownContent = dropdown.querySelector('.dropdown-content');

      dropdown.addEventListener('mouseenter', () => {
        dropdown.classList.add('active');
      });

      dropdown.addEventListener('mouseleave', () => {
        setTimeout(() => {
          if (!dropdown.matches(':hover')) {
            dropdown.classList.remove('active');
          }
        }); // Small delay to prevent flickering
      });
    });

    // Add event listener for studentNameSelectBox
    const studentNameSelectBox = document.getElementById('studentName');

    studentNameSelectBox.addEventListener('change', function() {
      const currentStudent = studentNameSelectBox.value;
      if (!saveFlag) {
        showError("unsavedChanges");
        studentNameSelectBox.value = previousStudent;
      }
      else {
        updateStudentData(currentStudent);
      }
    });

    // Add event listeners for select boxes
    const selectIds = document.querySelectorAll('#gender, #incomingGradeLevel, #enrollmentManager, #gradeLevelStatus, #enrolledInEEC, #evaluationEmail, #studentEvaluationForm, #contactedToSchedule, #screeningEmail, #reportCard, #iepDocumentation, #screeningFee, #adminAcceptance, #acceptanceEmail, #familyAcceptance, #blackbaudAccount, #birthCertificatePassport, #immunizationRecords, #admissionContractForm, #tuitionPaymentForm, #medicalConsentForm, #emergencyContactsForm, #registrationFee');

    selectIds.forEach(selectBox => {
      selectBox.addEventListener('change', () => {
        selectBox.style.backgroundColor = getColor(selectBox.value);
        saveAlert();
      });
    });

    // Add event listeners for input boxes
    const inputIds = document.querySelectorAll('#dateOfBirth, #parentGuardianName, #currentSchoolName, #currentTeacherName, #evaluationDueDate, #screeningDate, #screeningTime, #adminSubmissionDate, #acceptanceDueDate, #parentGuardianEmail, #currentTeacherEmail, #parentGuardianPhone, #enrollmentNotes');
    const phonePattern = /^\(\d{3}\) \d{3}-\d{4}$/;
    const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

    inputIds.forEach(input => {
      input.addEventListener('input', () => {
        saveAlert();
        if (input.id === 'parentGuardianPhone') {
          saveAlert();
          formatPhoneNumber(input);
          input.style.backgroundColor = phonePattern.test(input.value) ? getColor() : '';
        } else if (input.id === 'parentGuardianEmail' || input.id === 'currentTeacherEmail') {
          saveAlert();
          input.style.backgroundColor = emailPattern.test(input.value) ? getColor() : '';
        } else if (input.id === 'enrollmentNotes') {
          saveAlert();
        } else {
          saveAlert();
          input.style.backgroundColor = input.value.trim() !== '' ? getColor() : '';
        }
      });
    });

    // Allow deletion of select box entry, except for Student Name and modal selects
    document.querySelectorAll("select:not(#studentName, #templateSelect, #formSelect, #dataTypeSelect, #fileTypeSelect)").forEach(function(select) {
      select.addEventListener("keydown", function(event) {
        if (event.key === "Backspace" || event.key === "Delete") {
          if (!select.closest("#addStudentModal")) {
            saveAlert();
          }
          select.value = '';
          select.style.backgroundColor = '';
        }
      });
    });

    // Add event listener for phone number input on Add Student modal
    const addParentGuardianPhoneInput = document.getElementById('addParentGuardianPhone');
    addParentGuardianPhoneInput.addEventListener('input', function(event) {
      formatPhoneNumber(this);
    });

    // Add event listener for email template selection on Email modal
    const templateSelect = document.getElementById('templateSelect');
    templateSelect.addEventListener('change', function() {
      getEmailTemplate();
    });

    // Profile search event listener
    const profileSearchInput = document.getElementById('profileSearch');

    profileSearchInput.addEventListener('keyup', () => {
      const filter = profileSearchInput.value.toLowerCase();

      // Clear the current options in the studentNameSelectBox
      while (studentNameSelectBox.firstChild) {
        studentNameSelectBox.removeChild(studentNameSelectBox.firstChild);
      }

      // Filter the STUDENT_DATA based on the search input
      const filteredStudents = STUDENT_DATA.filter(student => {
        return Object.keys(student).some(key => {
          if ([
            'Student Name',
            'Gender', 
            'Date Of Birth', 
            'Incoming Grade Level', 
            'Grade Level Status', 
            'Enrollment Manager', 
            'Parent/Guardian Name', 
            'Parent/Guardian Phone', 
            'Current School Name', 
            'Current Teacher Name', 
            'Current Teacher Email', 
            'Enrolled In EEC'
          ].includes(key)) {
            const value = student[key] ? student[key].toString().toLowerCase() : '';
            const isMatch = value.includes(filter);
        
            // Log the key being checked and whether it matches
            return isMatch;
          }
          return false;
        });
      });

    // Populate the studentNameSelectBox with the filtered results
    filteredStudents.forEach(student => {
      const option = document.createElement('option');
      option.value = student['Student Name'];
      option.textContent = student['Student Name'];
      studentNameSelectBox.appendChild(option);
    });

    if (filteredStudents.length === 0) {
      const option = document.createElement('option');
      option.value = '';
      option.textContent = '';
      studentNameSelectBox.appendChild(option);
      updateStudentData("");
    }
    else {
      updateStudentData(filteredStudents[0]['Student Name']);
    }
  });
    
    console.log("Complete!");
  }

  // Get select box/input box color based on value
  function getColor(value) {
    switch (value) {
      case "Male":
        return 'var(--blue)';
      case "Female":
        return 'var(--pink)';
      case "Requested":
      case "In Progress":
      case "Pending":
      case "In Review":
        return 'var(--orange)';
      case "Waitlist":
      case "No":
      case "Rejected":
        return 'var(--red)';
      case "N/A":
      case "Non-binary":
        return 'var(--gray)';
      case "":
        return '';
      default:
        return 'var(--green)';
    }
  }

  function populateDashboard() {
    console.log("Populating dashboard...");

    // Add enrollment manager options to dashboard and 'Add Student' modal
    const enrollmentManagerSelect = document.getElementById('enrollmentManager');
    const addEnrollmentManagerSelect = document.getElementById('addEnrollmentManager');
    for (let i = 1; i <= 5; i++) {
      let key = `Enrollment Manager ${i}`;
      if (SETTINGS[key].trim() !== '') {
        let option = document.createElement('option');
        option.text = SETTINGS[key];
        enrollmentManagerSelect.add(option);
        addEnrollmentManagerSelect.add(option.cloneNode(true));
      }
    }

    // Set initial values for Add modal select boxes
    addStudentModal.querySelectorAll('select').forEach(function(select) {
      select.value = '';
    });
    
    updateStudentNames();
    
    console.log("Complete!");
  }
  
  /////////////////////
  // MODAL FUNCTIONS //
  /////////////////////

  function resetModal() {
    const modalInputs = document.querySelectorAll('#addStudentModal input, #addStudentModal select, #renameStudentModal input, #emailModal input, #emailModal select, #emailBody, #exportFormsModal input, #exportFormsModal select, #exportDataModal input, #exportDataModal select');
    
    modalInputs.forEach(function(input) {
      if (input.id === 'emailBody') {
        input.innerHTML = '';
      } else if (input.id === 'templateSelect' || input.id === 'formSelect' || input.id === 'dataTypeSelect' || input.id === 'fileTypeSelect') {
        input.selectedIndex = 0; // Reset to the first option
      } else {
        input.value = '';
      }
    });

    // Reset the scroll position of all modal bodies
    const modalBodies = document.querySelectorAll('.modal-htmlbody');
    modalBodies.forEach(modalBody => {
      modalBody.scrollTop = 0;
    });
  }

  //////////////////
  // SAVE PROFILE //
  //////////////////

  function saveProfile() {
    const studentNameSelectBox = document.getElementById('studentName');
    
    if (busyFlag) {
      showError("operationInProgress");
      return;
    }

    if (studentNameSelectBox.options.length === 0) {
      showError("missingData");
      return true;
    }
    
    const selectedStudent = studentNameSelectBox.value;
    let toastMessage;
    busyFlag = true;

    let student = STUDENT_DATA.find(function(item) {
      return item['Student Name'] === selectedStudent;
    });

    const elementKeyMappings = {
      'gender': 'Gender', 'dateOfBirth': 'Date Of Birth', 'incomingGradeLevel': 'Incoming Grade Level', 'gradeLevelStatus': 'Grade Level Status', 'enrollmentManager': 'Enrollment Manager',
      'parentGuardianName': 'Parent/Guardian Name', 'parentGuardianPhone': 'Parent/Guardian Phone', 'parentGuardianEmail': 'Parent/Guardian Email',
      'currentSchoolName': 'Current School Name', 'currentTeacherName': 'Current Teacher Name', 'currentTeacherEmail': 'Current Teacher Email',
      'enrolledInEEC': 'Enrolled In EEC',
      'evaluationDueDate': 'Evaluation Due Date', 'evaluationEmail': 'Evaluation Email Sent',
      'studentEvaluationForm': 'Evaluation Form',
      'contactedToSchedule': 'Contacted To Schedule', 'screeningDate': 'Screening Date', 'screeningTime': 'Screening Time', 'screeningEmail': 'Screening Email Sent',
      'reportCard': 'Report Card', 'iepDocumentation': 'IEP Documentation', 'screeningFee': 'Screening Fee',
      'adminSubmissionDate': 'Admin Submission Date', 'adminAcceptance': 'Admin Acceptance', 'acceptanceDueDate': 'Acceptance Due Date', 'acceptanceEmail': 'Acceptance Email Sent', 'familyAcceptance': 'Family Acceptance',
      'blackbaudAccount': 'Blackbaud Account', 'birthCertificatePassport': 'Birth Certificate/Passport',   'immunizationRecords': 'Immunization Records', 'admissionContractForm': 'Admission Contract Form',   'tuitionPaymentForm': 'Tuition Payment Form', 'medicalConsentForm': 'Medical Consent Form',  'emergencyContactsForm': 'Emergency Contacts Form', 'registrationFee': 'Registration Fee',
      'enrollmentNotes': 'Enrollment Notes'
    };

    // Update STUDENT_DATA object with dashboard data
    STUDENT_DATA.forEach((item) => {
      if (item['Student Name'] === selectedStudent) {
        Object.keys(elementKeyMappings).forEach((id) => {
          const element = document.getElementById(id);
          if (element) {
            item[elementKeyMappings[id]] = element.id === 'enrollmentNotes' ? element.innerHTML : element.value;
          }
        });
      }
    });
    
    const dashboardDataArray = [[
      selectedStudent,
      ...Object.keys(elementKeyMappings).map(key => student[elementKeyMappings[key]]),
    ]];

    // Update the state of 'Save Changes' button
    const saveChangesButton = document.getElementById('saveChangesButton');
    saveChangesButton.classList.remove('tool-bar-button-unsaved');
    saveFlag = true;
    
    // Show save confirmation toast
    google.script.run.withSuccessHandler(function(response) {
      if (response === "duplicateDatabaseEntry") {
        showError("duplicateDatabaseEntry");
      } 
      else if (response === "missingDatabaseEntry") {
        showError("missingDatabaseEntry");
      } 
      else {
        toastMessage = "'" + selectedStudent + "' saved successfully!";
        playNotificationSound("success");
        showToast("", toastMessage, 5000);
      }
      busyFlag = false;
    }).saveStudentData(DATA_SET, dashboardDataArray);
    
    toastMessage = "Saving changes...";
    showToast("", toastMessage, 5000);
  }

  /////////////////
  // ADD STUDENT //
  /////////////////

  function addStudent() {
    // Show error and prevent Add Student if there are unsaved changes
    if (!saveFlag) {
      showError("unsavedChanges");
      return;
    }

    if (busyFlag) {
      showError("operationInProgress");
      return;
    }
    
    showHtmlModal("addStudentModal")
    const addStudentModalButton = document.getElementById("addStudentModalButton");
    
    addStudentModalButton.onclick = function() {
      busyFlag = true;
      
      if (addStudentErrorCheck()) {
        busyFlag = false;
        return;
      }
      else {
        // If no errors, get the data from the 'Add Student' modal
        const firstName = document.getElementById('addFirstName').value;
        const lastName = document.getElementById('addLastName').value;
        const studentName = lastName + ", " + firstName;
        const newStudent = {
          'Student Name': studentName,
          'Gender': document.getElementById('addGender').value,
          'Date Of Birth': document.getElementById('addDateOfBirth').value,
          'Incoming Grade Level': document.getElementById('addIncomingGradeLevel').value,
          'Grade Level Status': document.getElementById('addGradeLevelStatus').value,
          'Enrollment Manager': document.getElementById('addEnrollmentManager').value,
          'Parent/Guardian Name': document.getElementById('addParentGuardianName').value,
          'Parent/Guardian Phone': document.getElementById('addParentGuardianPhone').value,
          'Parent/Guardian Email': document.getElementById('addParentGuardianEmail').value,
          'Current School Name': document.getElementById('addCurrentSchoolName').value,
          'Current Teacher Name': document.getElementById('addCurrentTeacherName').value,
          'Current Teacher Email': document.getElementById('addCurrentTeacherEmail').value,
          'Enrolled In EEC': document.getElementById('addEnrolledInEEC').value
        };

        // Check for duplicate student
        let duplicateStudent = false;

        for (let i = 0; i < STUDENT_DATA.length; i++) {
          if (STUDENT_DATA[i]['Student Name'] === newStudent['Student Name']) {
            duplicateStudent = true;
            break;
          }
        }

        // Exit the function
        if (duplicateStudent) {
          showError("duplicateDatabaseEntry");
          busyFlag = false;
          return;
        }

        // Add the new student to STUDENT_DATA and sort by name
        STUDENT_DATA.push(newStudent);
        sortStudentData();

        // Rebuild the select box with the new student names
        const studentNameSelectBox = document.getElementById('studentName');
        studentNameSelectBox.innerHTML = ''; // Clear the select box
        STUDENT_DATA.forEach(function(item) {
          let option = document.createElement('option');
          option.text = item['Student Name'];
          option.value = item['Student Name'];
          studentNameSelectBox.add(option);
        });

        // Switch the select box name to the new student
        studentNameSelectBox.value = newStudent['Student Name'];

        // Build the array of data for the sheet
        const newStudentArray = [
          newStudent['Student Name'],
          newStudent['Gender'],
          newStudent['Date Of Birth'],
          newStudent['Incoming Grade Level'],
          newStudent['Grade Level Status'],
          newStudent['Enrollment Manager'],
          newStudent['Parent/Guardian Name'],
          newStudent['Parent/Guardian Phone'],
          newStudent['Parent/Guardian Email'],
          newStudent['Current School Name'],
          newStudent['Current Teacher Name'],
          newStudent['Current Teacher Email'],
          newStudent['Enrolled In EEC']
        ];
    
        // Update the dashboard and close the 'Add Student' modal
        updateStudentData(newStudent['Student Name']);
        closeHtmlModal("addStudentModal");
        resetModal();

        google.script.run.withSuccessHandler(function(response) {
          if (!response) {
            showError("duplicateDatabaseEntry");
            for (let i = 0; i < STUDENT_DATA.length; i++) {
              if (STUDENT_DATA[i]['Student Name'] === newStudent['Student Name']) {
                STUDENT_DATA.splice(i, 1);
                break;
              }
            }
            sortStudentData();
            updateStudentNames();
            busyFlag = false;
            return;
          } else {
            const toastMessage = newStudent['Student Name'] + " added successfully!";
            playNotificationSound("success");
            showToast("", toastMessage, 5000);
          }
          busyFlag = false;
        }).addStudentData(newStudentArray);
      
        toastMessage = "Adding student...";
        showToast("", toastMessage, 5000);
      }
    };
  }

  function addStudentErrorCheck() {
    const firstName = document.getElementById('addFirstName').value;
    const lastName = document.getElementById('addLastName').value;
    const gender = document.getElementById('addGender').value;
    const dateOfBirth = document.getElementById('addDateOfBirth').value;
    const incomingGradeLevel = document.getElementById('addIncomingGradeLevel').value;
    const gradeLevelStatus = document.getElementById('addGradeLevelStatus').value;
    const enrollmentManager = document.getElementById('addEnrollmentManager').value;
    const parentGuardianName = document.getElementById('addParentGuardianName').value;
    const parentGuardianPhone = document.getElementById('addParentGuardianPhone').value;
    const parentGuardianEmail = document.getElementById('addParentGuardianEmail').value;
    const currentSchoolName = document.getElementById('addCurrentSchoolName').value;
    const currentTeacherName = document.getElementById('addCurrentTeacherName').value;
    const currentTeacherEmail = document.getElementById('addCurrentTeacherEmail').value;
    const enrolledInEEC = document.getElementById('addEnrolledInEEC').value;

    // Define regular expression patterns for error handling
    const phonePattern = /^\(\d{3}\) \d{3}-\d{4}$/;
    const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

    if (!firstName) {
      showError("missingFirstName");
      return true;
    }
    if (!lastName) {
      showError("missingLastName");
      return true;
    }
    if (!gender) {
      showError("missingGender");
      return true;
    }
    if (!dateOfBirth) {
      showError("missingDateOfBirth");
      return true;
    }
    if (!incomingGradeLevel) {
      showError("missingIncomingGradeLevel");
      return true;
    }
    if (!gradeLevelStatus) {
      showError("missingGradeLevelStatus");
      return true;
    }
    if (!enrollmentManager) {
      showError("missingEnrollmentManager");
      return true;
    }
    if (!parentGuardianName) {
      showError("missingParentGuardianName");
      return true;
    }
    if (!parentGuardianPhone) {
      showError("missingParentGuardianPhone");
      return true;
    }
    if (!phonePattern.test(parentGuardianPhone)) {
      showError("invalidParentGuardianPhone");
      return true;
    }
    if (!parentGuardianEmail) {
      showError("missingParentGuardianEmail");
      return true;
    }
    if (!emailPattern.test(parentGuardianEmail)) {
      showError("invalidParentGuardianEmail");
      return true;
    }
    if (!currentSchoolName) {
      showError("missingCurrentSchoolName");
      return true;
    }
    if (!currentTeacherName) {
      showError("missingCurrentTeacherName");
      return true;
    }
    if (!currentTeacherEmail) {
      showError("missingCurrentTeacherEmail");
      return true;
    }
    if (!emailPattern.test(currentTeacherEmail)) {
      showError("invalidCurrentTeacherEmail");
      return true;
    }
    if (!enrolledInEEC) {
      showError("missingEnrolledInEEC");
      return true;
    }

    return false;
  }

  ////////////////////
  // REMOVE STUDENT //
  ////////////////////

  function removeStudent() {
    if (busyFlag) {
      showError("operationInProgress");
      return;
    }
    
    if (removeStudentErrorCheck()) {
      busyFlag = false;
      return;
    } 

    const studentNameSelectBox = document.getElementById('studentName');
    const message = "Are you sure you want to remove and archive the data for '" + studentNameSelectBox.value + "'?";
    const title = `<i class="bi bi-exclamation-triangle-fill" style="color: var(--warning-color); margin-right: 10px;"></i>Remove Student`
    showModal(title, message, "Cancel", "Remove")
      .then((buttonText) => {
        if (buttonText === "Cancel") {
          return;
        } 
        else {
          busyFlag = true;
          const selectedIndex = studentNameSelectBox.selectedIndex; // Get the index of the selected option

          if (selectedIndex >= 0) {
            let selectedStudent = studentNameSelectBox.options[selectedIndex].value; 
            studentNameSelectBox.remove(selectedIndex); // Remove the selected option from the select box
            studentNameSelectBox.selectedIndex = -1; // Deselect any option after removal
        
            // Temporarily store and then remove the selected student object from the STUDENT_DATA array by name
            let removedStudent = STUDENT_DATA.find(student => student['Student Name'] === selectedStudent);
            STUDENT_DATA = STUDENT_DATA.filter(student => student['Student Name'] !== selectedStudent);

            updateStudentNames();

            let toastMessage = "Removing and achiving student...";

            google.script.run.withSuccessHandler(function(response) {
              if (response === "duplicateDatabaseEntry" || response === "missingDatabaseEntry") {
                STUDENT_DATA.push(removedStudent);
                sortStudentData();
                updateStudentNames();
                showError(response);
                busyFlag = false;
                return;
              }
              else {
                toastMessage = "'" + selectedStudent + "' removed and archived successfully!";
                playNotificationSound("remove");
                showToast("", toastMessage, 5000);
                busyFlag = false;
              }
            }).removeStudentData(selectedStudent);
          
            showToast("", toastMessage, 5000);
          }
        }
      }); 
  }

  function removeStudentErrorCheck() {
    const studentNameSelectBox = document.getElementById('studentName');
    
    if (studentNameSelectBox.options.length === 0) {
      showError("missingData");
      return true;
    }

    return false;
  }

  ////////////////////
  // RENAME STUDENT //
  ////////////////////

  function renameStudent() {
    const studentNameSelectBox = document.getElementById('studentName');
    
    if (!saveFlag) {
      showError("unsavedChanges");
      return;
    }

    if (busyFlag) {
      showError("operationInProgress");
      return;
    }

    if (studentNameSelectBox.options.length === 0) {
      showError("missingData");
      return true;
    }

    // Add the student's name to the modal    
    const selectedStudent = studentNameSelectBox.value;
    document.getElementById('currentStudentName').innerHTML = "";
    document.getElementById('currentStudentName').innerHTML = selectedStudent;
    
    showHtmlModal("renameStudentModal");
    const renameStudentModalButton = document.getElementById("renameStudentModalButton");
    
    renameStudentModalButton.onclick = function() {
      busyFlag = true;
      
      if (renameStudentErrorCheck()) {
        busyFlag = false;
        return;
      }
      else {
        const firstName = document.getElementById('renameFirst').value;
        const lastName = document.getElementById('renameLast').value;
        const newStudentName = lastName + ", " + firstName;

        // Check for duplicate student
        let duplicateStudent = false;

        for (let i = 0; i < STUDENT_DATA.length; i++) {
          if (STUDENT_DATA[i]['Student Name'] === newStudentName) {
            duplicateStudent = true;
            break;
          }
        }

        // Exit the function
        if (duplicateStudent) {
          showError("duplicateDatabaseEntry");
          busyFlag = false;
          return;
        }

        // Find and update the student name in STUDENT_DATA
        STUDENT_DATA.forEach(student => {
          if (student['Student Name'] === selectedStudent) {
            student['Student Name'] = newStudentName;
          }
        });
        sortStudentData();

        // Rebuild the select box with the new student names
        const studentNameSelectBox = document.getElementById('studentName');
        studentNameSelectBox.innerHTML = ''; // Clear the select box
        STUDENT_DATA.forEach(function(item) {
          let option = document.createElement('option');
          option.text = item['Student Name'];
          option.value = item['Student Name'];
          studentNameSelectBox.add(option);
        });
        
        // Switch the select box name to the renamed student
        studentNameSelectBox.value = newStudentName;

        // Update the dashboard and close the 'Add Student' modal
        updateStudentData(newStudentName);
        closeHtmlModal("renameStudentModal");
        resetModal();

        google.script.run.withSuccessHandler(function(response) {
          if (response === "duplicateDatabaseEntry" || response === "missingDatabaseEntry") {
            STUDENT_DATA.forEach(student => {
              if (student['Student Name'] === newStudentName) {
                student['Student Name'] = selectedStudent;
              }
            });
            sortStudentData();
            updateStudentNames();
            showError(response);
            busyFlag = false;
            return;
          }
          else {
            toastMessage = selectedStudent + " renamed to " + newStudentName + " successfully!";
            playNotificationSound("success");
            showToast("", toastMessage, 5000);
          }
          busyFlag = false;
        }).renameStudent(DATA_SET, selectedStudent, newStudentName);
      
        let toastMessage = "Renaming student...";
        showToast("", toastMessage, 5000);
      }
    };
  }

  function renameStudentErrorCheck() {
    const firstNameInput = document.getElementById('renameFirst').value;
    const lastNameInput = document.getElementById('renameLast').value;
    
    if (!firstNameInput) {
      showError("missingFirstName");
      return true;
    }

    if (!lastNameInput) {
      showError("missingLastName");
      return true;
    }
    
    return false;
  }

  ///////////////////
  // COMPOSE EMAIL //
  ///////////////////

  async function composeEmail() {
    if (busyFlag) {
      showError("operationInProgress");
      return;
    }
    
    document.getElementById('templateWarning').style.display = 'none';
    showHtmlModal("emailModal");

    const sendEmailModalButton = document.getElementById('sendEmailModalButton');
    sendEmailModalButton.onclick = async function() {
      busyFlag = true;
      
      if (sendEmailErrorCheck()) {
        busyFlag = false;
        return;
      }

      const template = document.getElementById('templateSelect').value;
      const recipient = document.getElementById('emailRecipient').value;
      const subject = document.getElementById('emailSubject').value;
      const body = document.getElementById('emailBody').innerHTML;
      let toastMessage = "";

      closeHtmlModal("emailModal");
      resetModal();

      if (template === "acceptance" || template === "acceptanceConditional") {
        toastMessage = "Attaching forms and sending email...";
        showToast("", toastMessage, 10000);
        setTimeout(async function () {  // Wrap in an async function inside setTimeout
          try {
              // Generate PDF and send the email
              await generatePDFPacket(recipient, subject, body);
          } catch (error) {
              console.error("Error sending email: ", error);
              showError("emailFailure");  // Optionally show an error if email fails
          }
        }, 100);  // Short delay to allow UI update to process before PDF generation
      }
      else {
        toastMessage = "Sending email...";
        showToast("", toastMessage, 5000);
        await sendEmail(recipient, subject, body);
      }
    }
  }

  async function generatePDFPacket(recipient, subject, body) {
    // Get the document definitions for each PDF
    const page1 = createAdmissionContract();
    const page2 = createTuitionPaymentOptions();
    const page3 = createMedicalConsentToTreat();
    const page4 = createStudentEmergencyContacts();
    const page5 = createBlackbaudTuitionInformation();
          
    // Concatenate the document definitions into one, adding page breaks
    const packetContent = [].concat(page1.content, { text: '', pageBreak: 'after' }, page2.content, { text: '', pageBreak: 'after' }, page3.content, { text: '', pageBreak: 'after' }, page4.content, { text: '', pageBreak: 'after' }, page5.content);

    // Use 'page1' to define other sections/styles since they are the same across all document pages
    const docDefinition = {
      pageSize: page1.pageSize,
      pageOrientation: page1.pageOrientation,
      pageMargins: page1.pageMargins,
      header: page1.header,
      footer: page1.footer,
      content: packetContent,
      defaultStyle: page1.defaultStyle,
      styles: page1.styles,
      images: page1.images
    };

    // Use async/await for PDF generation and sending the email
    return new Promise((resolve, reject) => {
      pdfMake.createPdf(docDefinition).getBlob((blob) => {
        blob.arrayBuffer().then((arrayBuffer) => {
          const uint8Array = new Uint8Array(arrayBuffer);
          const byteArray = Array.from(uint8Array);
          sendEmail(recipient, subject, body, byteArray)
          .then(resolve)  // Resolve the promise when done
          .catch(reject);  // Reject if there's an error
        });
      });
    });
  }

  function sendEmail(recipient, subject, body, attachments) {
    return new Promise((resolve, reject) => {
      google.script.run.withSuccessHandler(function(response) {
        if (response === "emailQuotaLimit") {
          showError("emailQuotaLimit");
          reject("emailQuotaLimit");
        } 
        else if (response === "emailFailure") {
          showError("emailFailure");
          reject("emailFailure");
        } 
        else {
          playNotificationSound("success");
          showToast("", "Email successfully sent to: " + recipient, 10000);
          resolve(response);
        }
        busyFlag = false;
      }).createEmail(recipient, subject, body, attachments);
    });
  }

  function getEmailTemplate() {
    // Get references to the selectbox and text areas
    const mergeData = getMergeData();
    const templateType = document.getElementById('templateSelect').value;
    const templateWarning = document.getElementById('templateWarning');
    const parentGuardianEmail = document.getElementById('parentGuardianEmail').value;
    const currentTeacherEmail = document.getElementById('currentTeacherEmail').value;
    const recipient = document.getElementById('emailRecipient');
    const subjectTemplate = document.getElementById('emailSubject');
    const bodyTemplate = document.getElementById('emailBody');
    
    templateWarning.style.display = 'none';

    // Update the template content based on the selected option
    switch (templateType) {
      case 'waitlist':
        recipient.value = parentGuardianEmail;
        subjectTemplate.value = SETTINGS['Waitlist Subject'];
        bodyTemplate.innerHTML = getEmailBody(SETTINGS['Waitlist Body'], mergeData);
        break;

      case 'evaluation':
        recipient.value = currentTeacherEmail;
        subjectTemplate.value = SETTINGS['Evaluation Subject'];
        bodyTemplate.innerHTML = getEmailBody(SETTINGS['Evaluation Body'], mergeData);
        break;

      case 'screeningEEC':      
        recipient.value = parentGuardianEmail;
        subjectTemplate.value = SETTINGS['Screening (EEC) Subject'];
        bodyTemplate.innerHTML = getEmailBody(SETTINGS['Screening (EEC) Body'], mergeData);
        break;

      case 'screeningSchool':      
        recipient.value = parentGuardianEmail;
        subjectTemplate.value = SETTINGS['Screening (School) Subject'];
        bodyTemplate.innerHTML = getEmailBody(SETTINGS['Screening (School) Body'], mergeData);
        break;

      case 'acceptance':
        recipient.value = parentGuardianEmail;
        subjectTemplate.value = SETTINGS['Acceptance Subject'];
        bodyTemplate.innerHTML = getEmailBody(SETTINGS['Acceptance Body'], mergeData);
        break;
      
      case 'acceptanceConditional':
        recipient.value = parentGuardianEmail;
        subjectTemplate.value = SETTINGS['Acceptance (Conditional) Subject'];
        bodyTemplate.innerHTML = getEmailBody(SETTINGS['Acceptance (Conditional) Body'], mergeData);
        break;
      
      case 'rejection':
        recipient.value = parentGuardianEmail;
        subjectTemplate.value = SETTINGS['Rejection Subject'];
        bodyTemplate.innerHTML = getEmailBody(SETTINGS['Rejection Body'], mergeData);
        break;
      
      case 'completion':
        recipient.value = parentGuardianEmail;
        subjectTemplate.value = SETTINGS['Completion Subject'];
        bodyTemplate.innerHTML = getEmailBody(SETTINGS['Completion Body'], mergeData);
        break;

      default:
        recipient.value = "";
        subjectTemplate.value = "";
        bodyTemplate.innerHTML = "";
        break;
    }
  }

  function getEmailBody(message, mergeData) {
    // Regular expression to match text within curly braces
    const regex = /{{([^}]+)}}/g;
    const warningIcon = '<i class="bi-exclamation-triangle-fill" style="color: var(--warning-color)"></i>';

    // Use replace() to find and replace text within curly braces
    const bodyTemplate = message.replace(regex, (match, variableName) => {
      
      // Check if the variable exists in the provided mapping
      if (mergeData.hasOwnProperty(variableName)) {
        // Replace the variable with its corresponding value
        return mergeData[variableName];
      }
      else {
        // If the variable is not found, leave it unchanged
        return match;
      }
    });

    if (bodyTemplate.includes(warningIcon)) {
      document.getElementById('templateWarning').style.display = "block";
    }

    return bodyTemplate;
  }

  function getMergeData() {
    // Split the student name into first and last
    const studentName = document.getElementById('studentName').value;
    let nameParts = studentName.match(/(\w+),\s*(\w+)/);
    let studentLastName = nameParts[1];
    let studentFirstName = nameParts[2];

    // Set the screening type and fee based on grade level
    const incomingGradeLevel = document.getElementById('incomingGradeLevel').value;
    const enrolledInEEC = document.getElementById('enrolledInEEC').value;
    const developmentalScreeningFeeEEC = SETTINGS['Developmental Screening Fee (EEC)'];
    const developmentalScreeningFeeTKK = SETTINGS['Developmental Screening Fee (TK/K)'];
    const academicScreeningFee = SETTINGS['Academic Screening Fee'];
    let screeningType;
    let screeningFee;

    if (!incomingGradeLevel) {
      screeningType = "";
      screeningFee = "";
    }
    else if (incomingGradeLevel === "Transitional Kindergarten" || incomingGradeLevel === "Kindergarten") {
      screeningType = "Developmental Screening";
      if (!enrolledInEEC) {
        screeningFee = "";
      }
      else if (enrolledInEEC === "Yes") {
        screeningFee = developmentalScreeningFeeEEC;
      }
      else if (enrolledInEEC === "No") {
        screeningFee = developmentalScreeningFeeTKK;
      }
    }
    else {
      screeningType = "Academic Screening";
      screeningFee = academicScreeningFee;
    }

    // Format evaluationDueDate
    const evaluationDueDate = document.getElementById('evaluationDueDate').value;
    const formattedEvaluationDueDate = formatDate(evaluationDueDate);

    // Format screeningDate
    const screeningDate = document.getElementById('screeningDate').value;
    const formattedScreeningDate = formatDate(screeningDate);

    // Format screeningTime
    const screeningTime = document.getElementById('screeningTime').value;
    const formattedScreeningTime = formatTime(screeningTime);

    // Format acceptanceDate
    const acceptanceDueDate = document.getElementById('acceptanceDueDate').value;
    const formattedAcceptanceDueDate = formatDate(acceptanceDueDate);

    // Create the mergeData object
    mergeData = {
      schoolYear: SETTINGS['School Year'],
      lastName: studentLastName,
      firstName: studentFirstName,
      gradeLevel: incomingGradeLevel,
      teacherName: document.getElementById('currentTeacherName').value,
      evaluationDueDate: formattedEvaluationDueDate,
      screeningType: screeningType, 
      screeningFee: screeningFee,
      screeningDate: formattedScreeningDate,
      screeningTime: formattedScreeningTime,
      acceptanceDueDate: formattedAcceptanceDueDate
    };

    const warningIcon = '<i class="bi-exclamation-triangle-fill" style="color: var(--warning-color)"></i>';

    // Add error icon to missing mergeData data
    Object.keys(mergeData).forEach(key => {
      if (mergeData[key] === "") {
        mergeData[key] = warningIcon;
      }
    });

    return mergeData;
  }

  function sendEmailErrorCheck() {
    const recipient = document.getElementById('emailRecipient').value;
    const body = document.getElementById('emailBody').innerHTML;
    const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    const warningIcon = '<i class="bi-exclamation-triangle-fill" style="color: var(--warning-color)"></i>';
    
    if (!recipient) {
      showError("missingEmailRecipient");
      return true;
    }

    const recipients = recipient.split(',');
    for (let i = 0; i < recipients.length; i++) {
      if (!emailPattern.test(recipients[i].trim())) {
        showError("invalidEmail");
        return true;
      }
    }

    if (body.includes(warningIcon)) {
      showError("missingEmailTemplateData")
      return true;
    }

    return false;
  }

  //////////////////
  // EXPORT FORMS //
  //////////////////

  function exportForms() {
    if (busyFlag) {
      showError("operationInProgress");
      return;
    }
    
    showHtmlModal("exportFormsModal");
    const exportFormsModalButton = document.getElementById('exportFormsModalButton');
    
    exportFormsModalButton.onclick = function() {
      busyFlag = true;
      const formType = document.getElementById('formSelect').value;
      
      closeHtmlModal("exportFormsModal");
      resetModal();

      setTimeout(function() {
        switch (formType) {
          case 'Admission Contract':
            pdfMake.createPdf(createAdmissionContract()).download('Admission Contract.pdf');
            break;
          case 'Tuition Payment Options':
            pdfMake.createPdf(createTuitionPaymentOptions()).download('Tuition Payment Options.pdf');
            break;
          case 'Medical Consent To Treat':
            pdfMake.createPdf(createMedicalConsentToTreat()).download('Medical Consent To Treat.pdf');
            break;
          case 'Student Emergency Contacts':
            pdfMake.createPdf(createStudentEmergencyContacts()).download('Student Emergency Contacts.pdf');
            break;
          case 'Blackbaud Tuition Information':
            pdfMake.createPdf(createBlackbaudTuitionInformation()).download('Blackbaud Tuition Information.pdf');
            break;
        }

        busyFlag = false;
      }, 100); // Short delay to allow UI update to process before PDF generation
    };
  }

  /////////////////
  // EXPORT DATA //
  /////////////////

  function exportData() {
    if (busyFlag) {
      showError("operationInProgress");
      return;
    }
    
    showHtmlModal("exportDataModal");
    const exportDataModalButton = document.getElementById('exportDataModalButton');
    
    exportDataModalButton.onclick = function() {
      busyFlag = true;
    
      const dataType = document.getElementById('dataTypeSelect').value;
      const fileType = document.getElementById('fileTypeSelect').value;
      let fileName;

      if (dataType === 'activeData') {
        fileName = 'Active Enrollment Data - ' + SETTINGS['School Year'];
      } else {
        fileName = 'Archive Enrollment Data - ' + SETTINGS['School Year'];
      }
      
      switch (fileType) {
        case 'csv':
          google.script.run.withSuccessHandler(function(data) {
            let a = document.createElement('a');
            
            a.href = 'data:text/csv;charset=utf-8,' + encodeURIComponent(data);
            a.download = fileName + '.csv';
            a.click();
            busyFlag = false;
          }).getCsv(dataType);
          break;
        case 'xlsx':
          google.script.run.withSuccessHandler(function(data) {
            // Convert the raw data into a Uint8Array
            const uint8Array = new Uint8Array(data);
                    
            // Create a Blob from the binary data
            const blob = new Blob([uint8Array], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
            
            const url = URL.createObjectURL(blob);
            let a = document.createElement('a');
            a.href = url;
            a.download = fileName + '.xlsx';
            a.click();
            URL.revokeObjectURL(url);
            busyFlag = false;
          }).getXlsx(dataType);
          break;
      }
      
      closeHtmlModal("exportDataModal");
      resetModal();
    };
  }

  //////////////////////////////
  // ARCHIVE/ACTIVE DATA VIEW //
  //////////////////////////////

  function toggleDataView() {
    // Show error and prevent Add Student if there are unsaved changes
    if (!saveFlag) {
      showError("unsavedChanges");
      return;
    }

    if (busyFlag) {
      showError("operationInProgress");
      return;
    }
    
    const toolbar = document.getElementById('toolbar');
    const page = document.getElementById('page');
    const loadingIndicator = document.getElementById('loading-indicator');

    // Show the loading indicator and hide the page
    toolbar.style.display = 'none';
    page.style.display = 'none';
    loadingIndicator.style.display = 'block';
    
    DATA_SET = (DATA_SET === "active") ? "archive" : "active";

    let promise;
    const header = document.getElementById('header-text');

    if (DATA_SET === "archive") {
      document.getElementById('header-text').innerText += " - Archive";
      promise = new Promise((resolve, reject) => {
        google.script.run.withSuccessHandler(resolve).withFailureHandler(reject).getArchiveData();
      });
    } else {
      header.innerText = header.innerText.replace(" - Archive", "");
      promise = new Promise((resolve, reject) => {
        google.script.run.withSuccessHandler(resolve).withFailureHandler(reject).getActiveData();
      });
    }

    promise.then(studentData => {
        STUDENT_DATA = studentData;
        
        if (DATA_SET === "active") {
          document.getElementById('addStudentButton').style.display = "block";
          document.getElementById('removeStudentButton').style.display = "block";
          document.getElementById('restoreStudentButton').style.display = "none";
          document.getElementById('deleteStudentButton').style.display = "none";
          document.getElementById('emailButton').style.display = "block";
          document.getElementById('archiveButton').style.display = "block";
          document.getElementById('backButton').style.display = "none";
        } 
        else {
          document.getElementById('addStudentButton').style.display = "none";
          document.getElementById('restoreStudentButton').style.display = "block";
          document.getElementById('deleteStudentButton').style.display = "block";
          document.getElementById('removeStudentButton').style.display = "none";
          document.getElementById('emailButton').style.display = "none";
          document.getElementById('archiveButton').style.display = "none";
          document.getElementById('backButton').style.display = "block";
        }

        updateStudentNames();

        // Hide loading indicator and show dashboard page
        loadingIndicator.style.display = 'none';
        toolbar.style.display = 'block';
        page.style.display = 'flex';
      
      }).catch(error => {
        console.error("Error fetching data:", error);
    });
  }
  
  /////////////////////
  // RESTORE STUDENT //
  /////////////////////

  function restoreStudent() {
    if (busyFlag) {
      showError("operationInProgress");
      return;
    }
    
    if (restoreStudentErrorCheck()) {
      busyFlag = false;
      return;
    }

    const studentNameSelectBox = document.getElementById('studentName');
    const message = "Are you sure you want to restore the data for '" + studentNameSelectBox.value + "'?";
    const title = `<i class="bi bi-exclamation-triangle-fill" style="color: var(--warning-color); margin-right: 10px;"></i>Archive Student`
    showModal(title, message, "Cancel", "Restore")
      .then((buttonText) => {
        if (buttonText === "Cancel") {
          return;
        } 
        else {
          busyFlag = true;
          const selectedIndex = studentNameSelectBox.selectedIndex; // Get the index of the selected option

          if (selectedIndex >= 0) {
            let selectedStudent = studentNameSelectBox.options[selectedIndex].value; 
            studentNameSelectBox.remove(selectedIndex); // Remove the selected option from the select box
            studentNameSelectBox.selectedIndex = -1; // Deselect any option after removal
        
            // Temporarily store and then remove the selected student object from the STUDENT_DATA array by name
            let removedStudent = STUDENT_DATA.find(student => student['Student Name'] === selectedStudent);
            STUDENT_DATA = STUDENT_DATA.filter(student => student['Student Name'] !== selectedStudent);

            updateStudentNames();
            
            let toastMessage = "Restoring student...";

            google.script.run.withSuccessHandler(function(response) {
              if (response === "duplicateDatabaseEntry" || response === "missingDatabaseEntry") {
                STUDENT_DATA.push(removedStudent);
                sortStudentData();
                updateStudentNames();
                showError(response);
                busyFlag = false;
                return;
              }
              else {
                toastMessage = "'" + selectedStudent + "' restored successfully!";
                playNotificationSound("success");
                showToast("", toastMessage, 5000);
                busyFlag = false;
              }
            }).restoreStudentData(selectedStudent);

            showToast("", toastMessage, 5000);
          }
        }
      });
  }

  function restoreStudentErrorCheck() {
    const studentNameSelectBox = document.getElementById('studentName');
    
    if (studentNameSelectBox.options.length === 0) {
      showError("missingData");
      return true;
    }

    return false;
  }

  ////////////////////
  // DELETE STUDENT //
  ////////////////////
  
  function deleteStudent() {
    if (busyFlag) {
      showError("operationInProgress");
      return;
    }
    
    if (deleteStudentErrorCheck()) {
      busyFlag = false;
      return;
    } 

    const studentNameSelectBox = document.getElementById('studentName');
    const message = "Are you sure you want to delete the data for '" + studentNameSelectBox.value + "'? This action cannot be undone.";
    const title = `<i class="bi bi-exclamation-triangle-fill" style="color: var(--warning-color); margin-right: 10px;"></i>Delete Student`
    showModal(title, message, "Cancel", "Delete")
      .then((buttonText) => {
        if (buttonText === "Cancel") {
          return;
        } 
        else {
          busyFlag = true;
          const selectedIndex = studentNameSelectBox.selectedIndex; // Get the index of the selected option

          if (selectedIndex >= 0) {
            let selectedStudent = studentNameSelectBox.options[selectedIndex].value; 
            studentNameSelectBox.remove(selectedIndex); // Remove the selected option from the select box
            studentNameSelectBox.selectedIndex = -1; // Deselect any option after removal
        
            // Temporarily store and then remove the selected student object from the STUDENT_DATA array by name
            let removedStudent = STUDENT_DATA.find(student => student['Student Name'] === selectedStudent);
            STUDENT_DATA = STUDENT_DATA.filter(student => student['Student Name'] !== selectedStudent);

            updateStudentNames();

            let toastMessage = "Deleting student...";

            google.script.run.withSuccessHandler(function(response) {
              if (response === "duplicateDatabaseEntry" || response === "missingDatabaseEntry") {
                STUDENT_DATA.push(removedStudent);
                sortStudentData();
                updateStudentNames();
                showError(response);
                busyFlag = false;
                return;
              } 
              else {
                toastMessage = "'" + selectedStudent + "' deleted successfully!";
                playNotificationSound("remove");
                showToast("", toastMessage, 5000);
                busyFlag = false;
              }
            }).deleteStudentData(selectedStudent);

            showToast("", toastMessage, 5000);
          }
        }
      });
  }

  function deleteStudentErrorCheck() {
    const studentNameSelectBox = document.getElementById('studentName');
    
    if (studentNameSelectBox.options.length === 0) {
      showError("missingData");
      return true;
    }

    return false;
  }

  ///////////////////////
  // UTILITY FUNCTIONS //
  ///////////////////////

  // Build the 'studentName' select box with student names
  function updateStudentNames() {
    const studentNameSelectBox = document.getElementById('studentName');
    studentNameSelectBox.innerHTML = ''; // Reset selectbox options
    
    if (Object.keys(STUDENT_DATA).length === 0) {
      console.log ("WARNING: No student data found!")
    }
    
    STUDENT_DATA.forEach(function(item) {
      let option = document.createElement('option');
      option.text = item['Student Name'];
      option.value = item['Student Name'];
      studentNameSelectBox.add(option);
    });

    // Update student information for the first student
    if (studentNameSelectBox.options.length > 0) {
      let selectedStudent = studentNameSelectBox.options[0].value;
      updateStudentData(selectedStudent);
    }
    else {
      updateStudentData();
    }
  }

  function updateStudentData(selectedStudent) {
    const clearAll = !selectedStudent || selectedStudent === "";
    
    let student = clearAll ? {} : STUDENT_DATA.find(function(item) {
      return item['Student Name'] === selectedStudent;
    });

    const elementKeyMappings = {
      'gender': 'Gender', 'dateOfBirth': 'Date Of Birth', 'incomingGradeLevel': 'Incoming Grade Level', 'gradeLevelStatus': 'Grade Level Status', 'enrollmentManager': 'Enrollment Manager',
      'parentGuardianName': 'Parent/Guardian Name', 'parentGuardianPhone': 'Parent/Guardian Phone', 'parentGuardianEmail': 'Parent/Guardian Email',
      'currentSchoolName': 'Current School Name', 'currentTeacherName': 'Current Teacher Name', 'currentTeacherEmail': 'Current Teacher Email',
      'enrolledInEEC': 'Enrolled In EEC',
      'evaluationDueDate': 'Evaluation Due Date', 'evaluationEmail': 'Evaluation Email Sent',
      'studentEvaluationForm': 'Evaluation Form',
      'contactedToSchedule': 'Contacted To Schedule', 'screeningDate': 'Screening Date', 'screeningTime': 'Screening Time', 'screeningEmail': 'Screening Email Sent',
      'reportCard': 'Report Card', 'iepDocumentation': 'IEP Documentation', 'screeningFee': 'Screening Fee',
      'adminSubmissionDate': 'Admin Submission Date', 'adminAcceptance': 'Admin Acceptance', 'acceptanceDueDate': 'Acceptance Due Date', 'acceptanceEmail': 'Acceptance Email Sent', 'familyAcceptance': 'Family Acceptance',
      'blackbaudAccount': 'Blackbaud Account', 'birthCertificatePassport': 'Birth Certificate/Passport',   'immunizationRecords': 'Immunization Records', 'admissionContractForm': 'Admission Contract Form',   'tuitionPaymentForm': 'Tuition Payment Form', 'medicalConsentForm': 'Medical Consent Form',  'emergencyContactsForm': 'Emergency Contacts Form', 'registrationFee': 'Registration Fee',
      'enrollmentNotes': 'Enrollment Notes'
    };

    Object.keys(elementKeyMappings).forEach((id) => {
      const element = document.getElementById(id);
      if (element) {
        if (id === 'enrollmentNotes') {
          element.innerHTML = clearAll || !student || student[elementKeyMappings[id]] === undefined ? "" : student[elementKeyMappings[id]];
        } else {
          element.value = clearAll || !student || student[elementKeyMappings[id]] === undefined ? "" : student[elementKeyMappings[id]];
        }
        
        // Trigger input event for input elements
        if (['dateOfBirth', 'parentGuardianPhone', 'parentGuardianEmail', 'currentTeacherEmail', 'parentGuardianName', 'currentSchoolName', 'currentTeacherName', 'evaluationDueDate', 'screeningDate', 'screeningTime', 'adminSubmissionDate', 'acceptanceDueDate'].includes(id)) {
          element.dispatchEvent(new Event('input', { bubbles: true }));
        }
        
        // Trigger change event for select elements
        if (['gender', 'incomingGradeLevel', 'enrollmentManager', 'gradeLevelStatus', 'enrolledInEEC', 'evaluationEmail', 'studentEvaluationForm', 'contactedToSchedule', 'screeningEmail', 'reportCard', 'iepDocumentation', 'screeningFee', 'adminAcceptance', 'acceptanceEmail', 'familyAcceptance', 'blackbaudAccount', 'birthCertificatePassport', 'immunizationRecords', 'admissionContractForm', 'tuitionPaymentForm', 'medicalConsentForm', 'emergencyContactsForm', 'registrationFee'].includes(id)) {
          element.dispatchEvent(new Event('change', { bubbles: true }));
        }
      }
    });

    const saveChangesButton = document.getElementById('saveChangesButton');
    saveChangesButton.classList.remove('tool-bar-button-unsaved');
    saveFlag = true;
    previousStudent = selectedStudent;
  }

  function sortStudentData() {
    STUDENT_DATA.sort((a, b) => {
      const nameA = a['Student Name'];
      const nameB = b['Student Name'];
      return nameA.localeCompare(nameB);
    });
  }

  function saveAlert() {
    saveFlag = false;
    saveChangesButton.classList.add('tool-bar-button-unsaved');
  }

  ////////////////////
  // ERROR HANDLING //
  ////////////////////

  function showError(errorType, callback = "") {
    const warningIcon = `<i class="bi bi-exclamation-triangle-fill" style="color: var(--warning-color);"></i>`;
    const errorIcon = `<i class="bi bi-x-circle-fill" style="color: var(--error-color);"></i>`;
    let title;
    let message;
    let button1;
    let button2;

    switch (errorType) {
      case "operationInProgress":
        title = warningIcon + "Operation In Progress";
        message = "Please wait until the operation completes and try again.";
        button1 = "Close";
        break;
      
      // Save errors
      case "unsavedChanges":
        title = warningIcon + "Unsaved Changes";
        message = "'" + previousStudent + "' has unsaved changes. Please save the changes and try again.";
        button1 = "Close";
        break;
      
      // Add student errors
      case "missingFirstName":
        title = warningIcon + "Missing First Name";
        message = "Please enter a first name and try again.";
        button1 = "Close";
        break;

      case "missingLastName":
        title = warningIcon + "Missing Last Name";
        message = "Please enter a last name and try again.";
        button1 = "Close";
        break;

      case "missingGender":
        title = warningIcon + "Missing Gender";
        message = "Please select a gender and try again.";
        button1 = "Close";
        break;

      case "missingDateOfBirth":
        title = warningIcon + "Missing Date Of Birth";
        message = "Please enter a date of birth and try again.";
        button1 = "Close";
        break;

      case "missingIncomingGradeLevel":
        title = warningIcon + "Missing Incoming Grade Level";
        message = "Please select an incoming grade level and try again.";
        button1 = "Close";
        break;

      case "missingGradeLevelStatus":
        title = warningIcon + "Missing Grade Level Status";
        message = "Please select a grade level status and try again.";
        button1 = "Close";
        break;

      case "missingEnrollmentManager":
        title = warningIcon + "Missing Enrollment Manager";
        message = "Please select an enrollment manager and try again.";
        button1 = "Close";
        break;

      case "missingParentGuardianName":
        title = warningIcon + "Missing Name";
        message = "Please enter a parent/guardian name and try again.";
        button1 = "Close";
        break;

      case "missingParentGuardianPhone":
        title = warningIcon + "Missing Phone Number";
        message = "Please enter a parent/guardian phone number and try again.";
        button1 = "Close";
        break;

      case "invalidParentGuardianPhone":
        title = warningIcon + "Invalid Phone Number";
        message = "Please check the parent/guardian phone number and try again.";
        button1 = "Close";
        break;

      case "missingParentGuardianEmail":
        title = warningIcon + "Missing Email Address";
        message = "Please enter a parent/guardian email address and try again.";
        button1 = "Close";
        break;

      case "invalidParentGuardianEmail":
        title = warningIcon + "Invalid Email Address";
        message = "Please check the parent/guardian email address and try again.";
        button1 = "Close";
        break;

      case "missingCurrentSchoolName":
        title = warningIcon + "Missing School Name";
        message = "Please enter a current school name and try again.";
        button1 = "Close";
        break;

      case "missingCurrentTeacherName":
        title = warningIcon + "Missing Teacher Name";
        message = "Please enter a current teacher name and try again.";
        button1 = "Close";
        break;

      case "missingCurrentTeacherEmail":
        title = warningIcon + "Missing Email Address";
        message = "Please enter a current teacher email address and try again.";
        button1 = "Close";
        break;

      case "invalidCurrentTeacherEmail":
        title = warningIcon + "Invalid Email Address";
        message = "Please check the current teacher email address and try again.";
        button1 = "Close";
        break;

      case "missingEnrolledInEEC":
        title = warningIcon + "Missing EEC Status";
        message = "Please select the EEC enrollment status and try again.";
        button1 = "Close";
        break;
      
      // Database errors
      case "missingData":
        title = errorIcon + "Data Error";
        message = "No student data found. The operation could not be completed.";
        button1 = "Close";
        break;

      case "missingDatabaseEntry":
        title = errorIcon + "Data Error";
        message = "The student data could not be found in the database. The operation could not be completed.";
        button1 = "Close";
        break;

      case "duplicateDatabaseEntry":
        title = errorIcon + "Data Error";
        message = "Duplicate data was found in the database. The operation could not be completed.";
        button1 = "Close";
        break;

      // Email errors
      case "missingEmailRecipient":
        title = warningIcon + "Missing Email Recipient";
        message = "Please enter an email address and try again.";
        button1 = "Close";
        break;
      
      case "invalidEmail":
        title = warningIcon + "Invalid Email Address";
        message = "Please check the email address and try again.";
        button1 = "Close";
        break;

      case "missingEmailTemplateData":
        title = errorIcon + "Email Error";
        message = "Missing email template data. The operation could not be completed.";
        button1 = "Close";
        break; 

      case "emailQuotaLimit":
        title = errorIcon + "Email Error";
        message = "The daily email limit has been reached. Please wait 24 hours before sending your email and try again.";
        button1 = "Close";
        break;

      case "emailFailure":
        title = errorIcon + "Email Error";
        message = "An unknown error occurred. The operation could not be completed.";
        button1 = "Close";
        break;
      
      // Backup errors
      case "backupFailure":
        title = errorIcon + "Backup Error";
        message = "An unknown error occurred. The operation could not be completed.";
        button1 = "Close";
        break;
    }
    
    playNotificationSound("alert");
    showModal(title, message, button1, button2);
  }
  
  /////////////////////
  // DATA VALIDATION //
  /////////////////////

  // Function to format phone number inputs
  function formatPhoneNumber(input) {
    // Remove all non-digit characters from the input value
    let inputValue = input.value.replace(/\D/g, "");

    // Limit the input value to 10 digits
    inputValue = inputValue.slice(0, 10);

    // Format the input value as '(XXX) XXX-XXXX'
    let formattedValue = '';
    for (let i = 0; i < inputValue.length; i++) {
      if (i === 0) {
        formattedValue += '(';
      } else if (i === 3) {
        formattedValue += ') ';
      } else if (i === 6) {
        formattedValue += '-';
      }
      formattedValue += inputValue[i];
    }

    // Update the input value with the formatted value
    input.value = formattedValue;
  }

  // Function to format date
  function formatDate(dateString) {
    if (!dateString) {
      return '';
    } else {
      // Split the date string into components
      const [year, month, day] = dateString.split('-').map(Number);

      // Create a date object using the local time zone
      const date = new Date(year, month - 1, day);

      // Format the date as 'MM/DD/YYYY'
      const options = { month: 'numeric', day: 'numeric', year: 'numeric' };
      return date.toLocaleDateString('en-US', options);
    }
  }

  // Function to format time
  function formatTime(timeString) {
    if (!timeString) {
      return '';
    } else {
      let hours = parseInt(timeString.split(':')[0]);
      let minutes = timeString.split(':')[1];
      let amPm = hours >= 12 ? 'PM' : 'AM';
      hours = hours % 12 || 12; // Convert hours to 12-hour format
      return `${hours}:${minutes} ${amPm}`;
    }
  }
  
</script>
