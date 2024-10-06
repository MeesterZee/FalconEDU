<script type="text/javascript">

  window.onload = function() {
    const page = document.getElementById('page');
    const loadingIndicator = document.getElementById('loading-indicator');
    const toolbar = document.getElementById('toolbar');

    setScheduleEventListeners();

    google.script.run.withSuccessHandler(function (data) {
      writeEvaluationDates(data.evaluationDates);
      writeScreeningDates(data.screeningDates);
      writeSubmissionDates(data.submissionDates);
      writeAcceptanceDates(data.acceptanceDates);
      
      // Hide the loading indicator and show the content
      loadingIndicator.style.display = 'none';
      toolbar.style.display = 'block';
      page.style.display = 'flex';
    }).getAllDates();
  }

  function setScheduleEventListeners() {
    // Evaluation search functionality
    const evaluationSearchInput = document.getElementById('evaluationSearch');
    const evaluationList = document.getElementById('evaluationDateList');
    const evaluations = evaluationList.getElementsByClassName('evaluation-item');

    evaluationSearchInput.addEventListener('keyup', () => {
      const filter = evaluationSearchInput.value.toLowerCase();

      Array.from(evaluations).forEach(evaluation => {
        const text = evaluation.textContent || evaluation.innerText;
        if (text.toLowerCase().includes(filter)) {
          evaluation.classList.remove('item-hide');
        } else {
          evaluation.classList.add('item-hide');
        }
      });
    });

    // Screening search functionality
    const screeningSearchInput = document.getElementById('screeningSearch');
    const screeningList = document.getElementById('screeningDateList');
    const screenings = screeningList.getElementsByClassName('screening-item');

    screeningSearchInput.addEventListener('keyup', () => {
      const filter = screeningSearchInput.value.toLowerCase();

      Array.from(screenings).forEach(screening => {
        const text = screening.textContent || screening.innerText;
        if (text.toLowerCase().includes(filter)) {
          screening.classList.remove('item-hide');
        } else {
          screening.classList.add('item-hide');
        }
      });
    });

    // Admin submission search functionality
    const submissionSearchInput = document.getElementById('submissionSearch');
    const submissionList = document.getElementById('submissionDateList');
    const submissions = submissionList.getElementsByClassName('submission-item');

    submissionSearchInput.addEventListener('keyup', () => {
      const filter = submissionSearchInput.value.toLowerCase();

      Array.from(submissions).forEach(submission => {
        const text = submission.textContent || submission.innerText;
        if (text.toLowerCase().includes(filter)) {
          submission.classList.remove('item-hide');
        } else {
          submission.classList.add('item-hide');
        }
      });
    });

    // Acceptance search functionality
    const acceptanceSearchInput = document.getElementById('acceptanceSearch');
    const acceptanceList = document.getElementById('acceptanceDateList');
    const acceptances = acceptanceList.getElementsByClassName('acceptance-item');

    acceptanceSearchInput.addEventListener('keyup', () => {
      const filter = acceptanceSearchInput.value.toLowerCase();

      Array.from(acceptances).forEach(acceptance => {
        const text = acceptance.textContent || acceptance.innerText;
        if (text.toLowerCase().includes(filter)) {
          acceptance.classList.remove('item-hide');
        } else {
          acceptance.classList.add('item-hide');
        }
      });
    });
  }

  function writeEvaluationDates(data) {
    const evaluationDateList = document.getElementById('evaluationDateList');

    if (data.length === 0) {
      evaluationDateList.innerHTML = `
        <div class="highlight-none">
        <i class="bi bi-check-circle"></i>No evaluations due!</div>`;
    } else {
      data.forEach((entry, index) => {
        const formattedDate = formatDate(entry.date);

        if (entry.status === 'Requested') {
          evaluationDateList.innerHTML += `
            <div class="evaluation-item highlight-orange">
            <i class="bi bi-person-fill"></i><b>${entry.student}</b>
            <br><i class="bi bi-calendar"></i>${formattedDate}
            <br><i class="bi bi-exclamation-triangle-fill" style="color: var(--warning-color);"></i><b>Evaluation requested</b></div>`;
        } else if (entry.status === 'Received') {
          evaluationDateList.innerHTML += `
            <div class="evaluation-item highlight-green">
            <i class="bi bi-person-fill"></i><b>${entry.student}</b>
            <br><i class="bi bi-calendar"></i>${formattedDate}
            <br><i class="bi bi-file-text"></i>Evaluation received</div>`;
        } else if (entry.status === 'N/A') {
          evaluationDateList.innerHTML += `
            <div class="evaluation-item highlight-gray">
            <i class="bi bi-person-fill"></i><b>${entry.student}</b>
            <br><i class="bi bi-calendar"></i>${formattedDate}
            <br><i class="bi bi-file-text"></i>Evaluation N/A</div>`;
        } else {
          evaluationDateList.innerHTML += `
            <div class="evaluation-item highlight-red">
            <i class="bi bi-person-fill"></i><b>${entry.student}</b>
            <br><i class="bi bi-calendar"></i>${formattedDate}
            <br><i class="bi bi-exclamation-triangle-fill" style="color: var(--warning-color);"></i><b>Evaluation status missing</b></div>`;
        }
      });
    }
  }

  function writeScreeningDates(data) {
    const screeningDateList = document.getElementById('screeningDateList');

    if (data.length === 0) {
      screeningDateList.innerHTML = `
        <div class="highlight-none">
        <i class="bi bi-check-circle"></i>No screenings scheduled!</div>`;
    } else {
      data.forEach((entry, index) => {
        const formattedDate = formatDate(entry.date);
        const formattedTime = formatTime(entry.time);
        
        if (entry.time === 'false') {
          screeningDateList.innerHTML += `
          <div class="screening-item highlight-red">
          <i class="bi bi-person-fill"></i><b>${entry.student}</b>
          <br><i class="bi bi-calendar""></i>${formattedDate}
          <br><i class="bi bi-exclamation-triangle-fill" style="color: var(--warning-color);"></i><b>Time missing</b></div>`;
        } else {
          screeningDateList.innerHTML += `
          <div class="screening-item highlight-green">
          <i class="bi bi-person-fill"></i><b>${entry.student}</b>
          <br><i class="bi bi-calendar"></i>${formattedDate}
          <br><i class="bi bi-alarm"></i>${formattedTime}</div>`;
        }
      });
    }
  }

  function writeSubmissionDates(data) {
    const submissionDateList = document.getElementById('submissionDateList');

    if (data.length === 0) {
      submissionDateList.innerHTML = `
        <div class="highlight-none">
        <i class="bi bi-check-circle"></i>No admin submissions due!`;
    } else {
      data.forEach((entry, index) => {
        const formattedDate = formatDate(entry.date);
        
        if (entry.date === 'false') {
          submissionDateList.innerHTML += `
          <div class="submission-item highlight-red">
          <i class="bi bi-person-fill"></i><b>${entry.student}</b>
          <br><i class="bi bi-exclamation-triangle-fill" style="color: var(--warning-color);"></i><b>Date missing</b></div>`;
        } else {
          submissionDateList.innerHTML += `
          <div class="submission-item highlight-green">
          <i class="bi bi-person-fill"></i><b>${entry.student}</b>
          <br><i class="bi bi-calendar"></i>${formattedDate}</div>`;
        }
      });
    }
  }

  function writeAcceptanceDates(data) {
    const acceptanceDateList = document.getElementById('acceptanceDateList');

    if (data.length === 0) {
      acceptanceDateList.innerHTML = `
        <div class="highlight-none">
        <i class="bi bi-check-circle"></i>No acceptances due!`;
    } else {
      data.forEach((entry, index) => {
        const formattedDate = formatDate(entry.date);
        const documentStatus = Object.values(entry).slice(2); // Exclude student and date properties
        let status;

        const receivedCount = documentStatus.filter(entry => entry === 'Received' || entry === 'N/A').length;
        const requestedCount = documentStatus.filter(entry => entry === 'Requested').length;
        const statusMissingCount = documentStatus.filter(entry => entry === 'Status missing').length;

        // Logic to assign color status
        if (statusMissingCount > 0) {
          status = 'red'; // Some entries are 'Status missing'
        } else if (requestedCount > 0) {
          if (receivedCount === documentStatus.length) {
            status = 'green'; // All entries are 'Received' (including 'N/A')
          } else {
            status = 'yellow'; // Some entries are 'Requested' (even if some are 'Received')
          }
        } else {
          status = 'green'; // All entries are 'Received' (including 'N/A') and no 'Requested' or 'Status missing'
        }

        // Adjusted comparison for generating HTML based on the status variable
        if (status === 'yellow') {
          acceptanceDateList.innerHTML += `
          <div class="acceptance-item highlight-orange">
          <i class="bi bi-person-fill"></i><b>${entry.student}</b>
          <br><i class="bi bi-calendar"></i>${formattedDate}
          <br><i class="bi bi-exclamation-triangle-fill" style="color: var(--warning-color);"></i><b>Documents/fee requested</b></div>`;
        } else if (status === 'green') {
          acceptanceDateList.innerHTML += `
          <div class="acceptance-item highlight-green">
          <i class="bi bi-person-fill"></i><b>${entry.student}</b>
          <br><i class="bi bi-calendar"></i>${formattedDate}
          <br><i class="bi bi-file-text"></i>Documents/fee received</div>`;
        } else if (status === 'gray') {
          acceptanceDateList.innerHTML += `
          <div class="acceptance-item highlight-gray">
          <i class="bi bi-person-fill"></i><b>${entry.student}</b>
          <br><i class="bi bi-calendar"></i>${formattedDate}
          <br><i class="bi bi-file-text"></i>Documents/fee N/A</div>`;
        } else {
          acceptanceDateList.innerHTML += `
          <div class="acceptance-item highlight-red">
          <i class="bi bi-person-fill"></i><b>${entry.student}</b>
          <br><i class="bi bi-calendar"></i>${formattedDate}
          <br><i class="bi bi-exclamation-triangle-fill" style="color: var(--warning-color);"></i><b>Documents/fee status missing</div>`;
        }
      });
    }
  }

  ///////////////////////
  // UTILITY FUNCTIONS //
  ///////////////////////
  
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
