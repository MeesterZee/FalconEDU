<script type="text/javascript">
  // Global variables  
  // let USER_SETTINGS; // Defined in HTML
  let SETTINGS;
  let EMAIL_TEMPLATE_DATA;
  let saveFlag = true;
  let busyFlag = false;
  
  // Initialize application
  // Conversion to async allows for parallel data retrieval from Apps Script
  window.onload = async function() {
    console.log("Initializing settings...");

    const toolbar = document.getElementById('toolbar');
    const page = document.getElementById('page');
    const loadingIndicator = document.getElementById('loading-indicator');
    
    try {
      // Show loading indicator
      loadingIndicator.style.display = 'block';
      toolbar.style.display = 'none';
      page.style.display = 'none';

      // Fetch data in parallel (async not needed but allows for future data streams)
      const [allSettings] = await Promise.all([
        new Promise((resolve, reject) => {
          google.script.run.withSuccessHandler(resolve).withFailureHandler(reject).getSettings();
        })
      ]);

      // Set global variables
      SETTINGS = allSettings;

      // Populate elements with data
      await Promise.all([
        setEventListeners(),
        setColorPicker(),
        loadSettings()
      ]);
    
      console.log("Initialization complete!");
    
    } catch (error) {
        console.error("Error during initialization: ", error);
    
    } finally {
      // Hide loading indicator and show page
      loadingIndicator.style.display = 'none';
      toolbar.style.display = 'block';
      page.style.display = 'flex';
    }
  };

  function setEventListeners() {
    console.log("Setting event listeners...");

    const allTextInputs = document.querySelectorAll('.table-input[type=text]');
    const allSelects = document.querySelectorAll('.table-select, #theme');
    const allCurrencyInputs = document.querySelectorAll('.table-input[data-type="currency"]');
    const saveChangesButton = document.getElementById('saveChangesButton');
    const themeSelect = document.getElementById('theme');
    const alertSoundSelect = document.getElementById('alertSound');
    const emailSoundSelect = document.getElementById('emailSound');
    const successSoundSelect = document.getElementById('successSound');
    const removeSoundSelect = document.getElementById('removeSound');
    
    window.addEventListener('beforeunload', function (e) {
      if (!saveFlag) {
        e.preventDefault();
        e.returnValue = '';
      }
    });

    allTextInputs.forEach(input => {
      input.addEventListener('input', saveAlert);
    });

    allSelects.forEach(select => {
      select.addEventListener('change', saveAlert);
    });

    allCurrencyInputs.forEach(input => {
      input.addEventListener('input', saveAlert);
      input.addEventListener('input', function(event) {
        formatCurrency(this);
      });
    });

    themeSelect.addEventListener('change', function() {
      const theme = document.getElementById('theme').value;
      const customTheme = document.getElementById('customTheme');

      if (theme === "custom") {
        customTheme.style.display = 'block';
      } else {
        customTheme.style.display = 'none';
      }
    });

    alertSoundSelect.addEventListener('change', function() {
      USER_SETTINGS.alertSound = alertSoundSelect.value;
      playNotificationSound("alert");
    });

    emailSoundSelect.addEventListener('change', function() {
      USER_SETTINGS.emailSound = emailSoundSelect.value;
      playNotificationSound("email");
    });

    successSoundSelect.addEventListener('change', function() {
      USER_SETTINGS.successSound = successSoundSelect.value;
      playNotificationSound("success");
    });

    removeSoundSelect.addEventListener('change', function() {
      USER_SETTINGS.removeSound = removeSound.value;
      playNotificationSound("remove");
    });

    document.getElementById('silentModeSwitch').addEventListener('change', saveAlert);
    
    document.getElementById('templateSubject').addEventListener('input', saveAlert);
    document.getElementById('templateBody').addEventListener('input', saveAlert);

    saveChangesButton.addEventListener('click', saveSettings);

    console.log("Complete!");
  }

  function setColorPicker() {
    const themeType = document.getElementById('themeTypeSelect');
    const primaryColorPicker = document.getElementById('primaryColorPicker');
    const accentColorPicker = document.getElementById('accentColorPicker');

    themeTypeSelect.value = getComputedStyle(document.documentElement).getPropertyValue('color-scheme').trim();
    primaryColorPicker.value = getComputedStyle(document.documentElement).getPropertyValue('--primary-color').trim();
    accentColorPicker.value = getComputedStyle(document.documentElement).getPropertyValue('--accent-color').trim();
  }

  function loadSettings(settings) {
    console.log("Loading settings...");

    // Appearance
    const themeSelect = document.getElementById('theme');
    const customTheme = document.getElementById('customTheme');

    if (USER_SETTINGS.theme === "custom") {
      customTheme.style.display = 'block';
    } else {
      customTheme.style.display = 'none';
    }

    themeSelect.value = USER_SETTINGS.theme;

    // Sound Effects
    const silentModeChecked = USER_SETTINGS.silentMode === 'true'; // Convert string to boolean
    document.getElementById('silentModeSwitch').checked = silentModeChecked;
    document.getElementById('alertSound').value = USER_SETTINGS.alertSound;
    document.getElementById('emailSound').value = USER_SETTINGS.emailSound;
    document.getElementById('removeSound').value = USER_SETTINGS.removeSound;
    document.getElementById('successSound').value = USER_SETTINGS.successSound;
    
    //School Information
    document.getElementById('schoolName').value = SETTINGS['School Name'];
    document.getElementById('enrollmentYear').value = SETTINGS['School Year'];
    
    //Enrollment Managers
    document.getElementById('managerName1').value = SETTINGS['Enrollment Manager 1'];
    document.getElementById('managerName2').value = SETTINGS['Enrollment Manager 2'];
    document.getElementById('managerName3').value = SETTINGS['Enrollment Manager 3'];
    document.getElementById('managerName4').value = SETTINGS['Enrollment Manager 4'];
    document.getElementById('managerName5').value = SETTINGS['Enrollment Manager 5'];

    //Screening Fees
    document.getElementById('screeningFeeEEC').value = SETTINGS['Developmental Screening Fee (EEC)'];
    document.getElementById('screeningFeeTKK').value = SETTINGS['Developmental Screening Fee (TK/K)'];
    document.getElementById('screeningFee18').value = SETTINGS['Academic Screening Fee'];
    
    //School Fees
    document.getElementById('registrationFee').value = SETTINGS['Registration Fee'];
    document.getElementById('hugFee').value = SETTINGS['HUG Fee'];
    document.getElementById('familyCommitmentFee').value = SETTINGS['Family Commitment Fee'];
    document.getElementById('flashFee').value = SETTINGS['FLASH Processing Fee'];
    document.getElementById('withdrawalFee').value = SETTINGS['Withdrawal Processing Fee'];

    // Email Templates
    EMAIL_TEMPLATE_DATA = {
      waitlist: {
        subject: SETTINGS['Waitlist Subject'],
        body: SETTINGS['Waitlist Body'],
        unsavedSubject: '',
        unsavedBody: ''
      },
      evaluation: {
        subject: SETTINGS['Evaluation Subject'],
        body: SETTINGS['Evaluation Body'],
        unsavedSubject: '',
        unsavedBody: ''
      },
      eecScreening: {
        subject: SETTINGS['Screening (EEC) Subject'],
        body: SETTINGS['Screening (EEC) Body'],
        unsavedSubject: '',
        unsavedBody: ''
      },
      screening: {
        subject: SETTINGS['Screening (School) Subject'],
        body: SETTINGS['Screening (School) Body'],
        unsavedSubject: '',
        unsavedBody: ''
      },
      acceptance: {
        subject: SETTINGS['Acceptance Subject'],
        body: SETTINGS['Acceptance Body'],
        unsavedSubject: '',
        unsavedBody: ''
      },
      acceptanceConditional: {
        subject: SETTINGS['Acceptance (Conditional) Subject'],
        body: SETTINGS['Acceptance (Conditional) Body'],
        unsavedSubject: '',
        unsavedBody: ''
      },
      rejection: {
        subject: SETTINGS['Rejection Subject'],
        body: SETTINGS['Rejection Body'],
        unsavedSubject: '',
        unsavedBody: ''
      },
      completion: {
        subject: SETTINGS['Completion Subject'],
        body: SETTINGS['Completion Body'],
        unsavedSubject: '',
        unsavedBody: ''
      }
    };

    let templateTypeSelect = document.getElementById('templateType');
    let templateSubjectInput = document.getElementById('templateSubject');
    let templateBodyInput = document.getElementById('templateBody');

    // Set the default template type to "waitlist"
    let defaultTemplate = EMAIL_TEMPLATE_DATA['waitlist'];
    defaultTemplate.unsavedSubject = defaultTemplate.subject;
    defaultTemplate.unsavedBody = defaultTemplate.body;
    templateSubjectInput.value = defaultTemplate.subject;
    templateBodyInput.innerHTML = defaultTemplate.body;

    templateSubjectInput.addEventListener('input', function() {
      let selectedTemplate = EMAIL_TEMPLATE_DATA[templateTypeSelect.value];
      if (selectedTemplate) {
        selectedTemplate.unsavedSubject = templateSubjectInput.value;
      }
    });

    templateBodyInput.addEventListener('input', function() {
      let selectedTemplate = EMAIL_TEMPLATE_DATA[templateTypeSelect.value];
      if (selectedTemplate) {
        selectedTemplate.unsavedBody = templateBodyInput.innerHTML;
      }
    });

    templateTypeSelect.addEventListener('change', function() {
      let selectedTemplate = EMAIL_TEMPLATE_DATA[templateTypeSelect.value];
      if (selectedTemplate) {
        templateSubjectInput.value = selectedTemplate.unsavedSubject || selectedTemplate.subject;
        templateBodyInput.innerHTML = selectedTemplate.unsavedBody || selectedTemplate.body;
      }
    });

    console.log("Complete!");
  }

  function saveSettings() {
    if (busyFlag) {
      showError("operationInProgress");
      busyFlag = false;
      return;
    }
    
    busyFlag = true;
    saveChangesButton.classList.remove('tool-bar-button-unsaved');

    const header = document.getElementById('header-text');
    const enrollmentYear = document.getElementById('enrollmentYear').value;
    const themeSetting = document.getElementById('theme').value;
    const silentModeSetting = document.getElementById('silentModeSwitch').checked;
    const alertSound = document.getElementById('alertSound').value;
    const emailSound = document.getElementById('emailSound').value;
    const removeSound = document.getElementById('removeSound').value;
    const successSound = document.getElementById('successSound').value;

    EMAIL_TEMPLATE_DATA.waitlist.subject = EMAIL_TEMPLATE_DATA.waitlist.unsavedSubject;
    EMAIL_TEMPLATE_DATA.waitlist.body = EMAIL_TEMPLATE_DATA.waitlist.unsavedBody;
    
    const schoolSettings = [[
      document.getElementById('schoolName').value,
      enrollmentYear
    ]];

    const managerSettings = [[
      document.getElementById('managerName1').value,
      document.getElementById('managerName2').value,
      document.getElementById('managerName3').value,
      document.getElementById('managerName4').value,
      document.getElementById('managerName5').value
    ]];

    const feeSettings = [[
      parseFloat(document.getElementById('screeningFeeEEC').value.replace(/[^\d.]/g, '')),
      parseFloat(document.getElementById('screeningFeeTKK').value.replace(/[^\d.]/g, '')),
      parseFloat(document.getElementById('screeningFee18').value.replace(/[^\d.]/g, '')),
      parseFloat(document.getElementById('registrationFee').value.replace(/[^\d.]/g, '')),
      parseFloat(document.getElementById('hugFee').value.replace(/[^\d.]/g, '')),
      parseFloat(document.getElementById('familyCommitmentFee').value.replace(/[^\d.]/g, '')),
      parseFloat(document.getElementById('flashFee').value.replace(/[^\d.]/g, '')),
      parseFloat(document.getElementById('withdrawalFee').value.replace(/[^\d.]/g, ''))
    ]];

    Object.keys(EMAIL_TEMPLATE_DATA).forEach(key => {
      if (EMAIL_TEMPLATE_DATA[key].unsavedSubject !== "" || EMAIL_TEMPLATE_DATA[key].unsavedBody !== "") {
        if (EMAIL_TEMPLATE_DATA[key].unsavedSubject !== "") {
          EMAIL_TEMPLATE_DATA[key].subject = EMAIL_TEMPLATE_DATA[key].unsavedSubject;
        }
        if (EMAIL_TEMPLATE_DATA[key].unsavedBody !== "") {
          EMAIL_TEMPLATE_DATA[key].body = EMAIL_TEMPLATE_DATA[key].unsavedBody;
        }
      }
    });

    const emailTemplateSubject = [[
      EMAIL_TEMPLATE_DATA.waitlist.subject, 
      EMAIL_TEMPLATE_DATA.evaluation.subject, 
      EMAIL_TEMPLATE_DATA.eecScreening.subject, 
      EMAIL_TEMPLATE_DATA.screening.subject, 
      EMAIL_TEMPLATE_DATA.acceptance.subject, 
      EMAIL_TEMPLATE_DATA.acceptanceConditional.subject, 
      EMAIL_TEMPLATE_DATA.rejection.subject,
      EMAIL_TEMPLATE_DATA.completion.subject
    ]];
    
    const emailTemplateBody = [[
      EMAIL_TEMPLATE_DATA.waitlist.body,
      EMAIL_TEMPLATE_DATA.evaluation.body,
      EMAIL_TEMPLATE_DATA.eecScreening.body,
      EMAIL_TEMPLATE_DATA.screening.body,
      EMAIL_TEMPLATE_DATA.acceptance.body,
      EMAIL_TEMPLATE_DATA.acceptanceConditional.body, 
      EMAIL_TEMPLATE_DATA.rejection.body,
      EMAIL_TEMPLATE_DATA.completion.body
    ]];
    
    //Update the page header
    saveFlag = true;
    saveTheme();
    setColorPicker();
    saveSound();
    header.innerHTML = "Falcon Enrollment - " + enrollmentYear;
    playNotificationSound("success");
    showToast("", "Settings saved!", 5000);
    
    google.script.run.writeSettings(USER_SETTINGS, schoolSettings, managerSettings, feeSettings, emailTemplateSubject, emailTemplateBody);

    busyFlag = false;
  }

  ///////////////////////
  // UTILITY FUNCTIONS //
  ///////////////////////

  function showError(errorType, callback = "") {
    let icon = `<i class="bi bi-exclamation-triangle-fill" style="color: var(--warning-color);"></i>`;
    let title;
    let message;
    let button1;
    let button2;

    switch (errorType) {
      case "operationInProgress":
        title = icon + "Error";
        message = "Operation currently in progress. Please wait until the operation completes and try again.";
        button1 = "Close";
        break;
    }

    playNotificationSound("alert");
    showModal(title, message, button1, button2);
  }
  
  function saveAlert() {
    saveFlag = false;
    saveChangesButton.classList.add('tool-bar-button-unsaved');
  }

  /////////////////////
  // DATA VALIDATION //
  /////////////////////
  
  function formatCurrency(input) {
    // Remove all non-digit characters except for the first period
    let inputValue = input.value.replace(/[^\d.]+/g, "");

    // Split the input value into dollars and cents
    let [dollars, cents] = inputValue.split(".");

    // Add commas to the dollars part
    let formattedDollars = '';
    let count = 0;
    for (let i = dollars.length - 1; i >= 0; i--) {
      count++;
      formattedDollars = dollars.charAt(i) + formattedDollars;
      if (count % 3 === 0 && i !== 0) {
        formattedDollars = ',' + formattedDollars;
      }
    }

    // Limit cents to two digits
    if (cents !== undefined) {
      cents = cents.slice(0, 2); // Take only the first two characters
    }

    // Combine dollars and cents with a period
    let formattedValue = '$' + formattedDollars;
    if (cents !== undefined) {
      formattedValue += '.' + cents;
    }

    // Update the input value with the formatted value
    input.value = formattedValue;
  }
  
</script>
