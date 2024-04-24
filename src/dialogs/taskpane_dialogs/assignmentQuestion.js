(function () {
  "use strict";

  var professorsData;
  var templateId;

  // The Office initialize function must be run each time a new page is loaded.
  Office.initialize = function (reason) {
    $(document).ready(function () {
      templateId = getParameterByName("templateId");
      $("#assignment-extension-date").datepicker();
      registerEvents();
      hideInputBasedOnTemplateId();
      populateProfessorsData();
    });
  };

  function getParameterByName(name, url) {
    if (!url) {
      url = window.location.href;
    }
    name = name.replace(/[\[\]]/g, "\\$&");
    const regex = new RegExp("[?&]" + name + "(=([^&#]*)|&|#|$)"),
      results = regex.exec(url);
    if (!results) return null;
    if (!results[2]) return "";
    return decodeURIComponent(results[2].replace(/\+/g, " "));
  }

  function registerEvents() {
    document.getElementById("template-form").onsubmit = validateForm;
    document.getElementById("professor-name").onchange = professorNameChange;
    if (templateId == "template-1" || templateId == "template-2") {
      $("#assignment-name").prop("required", true);
      $("#assignment-question-content").prop("required", true);
    }
    if (templateId == "template-3") {
      $("#onbase-number").prop("required", true);
    }
    if (templateId == "template-4") {
      $("#assignment-name").prop("required", true);
      $("#assignment-extension-date").prop("required", true);
      $("#assignment-extension-reason").prop("required", true);
    }
    if(templateId == "template-5") {
      $("#topic-clarification").prop('required', true);
      $("#topic-quesion").prop('required', true);
    }
  }

  function hideInputBasedOnTemplateId() {
    if (templateId === "template-1" || templateId === "template-2" || templateId === "template-4") {
      let assignmentNameInput = document.getElementById("assignment-name-input");
      assignmentNameInput.classList.remove("hide");
    }
    if (templateId === "template-1" || templateId === "template-2") {
      document.getElementById("assignment-question-input").classList.remove("hide");
    }
    if (templateId === "template-3") {
      document.getElementById("onbase-input").classList.remove("hide");
    }
    if (templateId === "template-4") {
      var dtToday = new Date();

      var month = dtToday.getMonth() + 1;
      var day = dtToday.getDate();
      var year = dtToday.getFullYear();
      if (month < 10) month = "0" + month.toString();
      if (day < 10) day = "0" + day.toString();

      var maxDate = year + "-" + month + "-" + day;

      $("#assignment-extension-date").attr("min", maxDate);
      document.getElementById("extension-date-input").classList.remove("hide");
      document.getElementById("extension-reason").classList.remove("hide");
    }
    if(templateId === "template-5") {
      document.getElementById("clarification-topic-block").classList.remove("hide");
      document.getElementById("topic-quesion-block").classList.remove("hide");
    }
  }

  function populateProfessorsData() {
    fetch("./professors.json")
      .then((response) => {
        if (!response.ok) {
          throw new Error("Network response was not ok");
        }
        return response.json();
      })
      .then((data) => {
        professorsData = data.professors;
        const professors = professorsData.map((professor) => professor.name);

        // Populate the dropdown with professors' names
        const selectElement = document.getElementById("professor-name");
        professors.forEach((professorName) => {
          const option = document.createElement("option");
          option.value = professorName;
          option.textContent = professorName;
          selectElement.appendChild(option);
        });
      })
      .catch((error) => {
        console.error("Error fetching professors data:", error);
      });
  }

  function professorNameChange(event) {
    const selectedProfessor = event.target.value;

    if (selectedProfessor) {
      populateCourseNumbers(selectedProfessor);
    } else {
      // Clear course numbers select if no professor is selected
      const courseNumbersSelect = document.getElementById("course-no");
      courseNumbersSelect.innerHTML = "<option value='' disabled selected>Select Course Number</option>";
    }
  }

  // Function to populate the course numbers based on the selected professor
  function populateCourseNumbers(selectedProfessor) {
    const courseNumbersSelect = document.getElementById("course-no");
    courseNumbersSelect.innerHTML = ""; // Clear existing options
    //Default option building
    courseNumbersSelect.innerHTML = "<option value='' disabled selected>Select Course Number</option>";

    // Search for the professor object by iterating over the array
    for (let i = 0; i < professorsData.length; i++) {
      const professor = professorsData[i];
      if (professor.name === selectedProfessor) {
        professor.courses.forEach((course) => {
          const option = document.createElement("option");
          option.value = course;
          option.textContent = course;
          courseNumbersSelect.appendChild(option);
        });
        break; // Exit loop once the professor is found
      }
    }
  }

  function sendMessage(message) {
    Office.context.ui.messageParent(message);
  }

  function validateForm() {
    let data = {};
    data.professorName = document.forms["assignment-template-form"]["professor-name"].value;
    data.courseNumber = document.forms["assignment-template-form"]["course-no"].value;
    data.assignmentName = document.forms["assignment-template-form"]["assignment-name"].value;
    data.assignmentQuestion = document.forms["assignment-template-form"]["assignment-question-content"].value;
    data.studentId = document.forms["assignment-template-form"]["student-id"].value;
    data.studentName = document.forms["assignment-template-form"]["student-name"].value;
    data.onbaseNumber = document.forms["assignment-template-form"]["onbase-number"].value;
    data.extendDate = document.forms["assignment-template-form"]["assignment-extension-date"].value;
    data.extendReason = document.forms["assignment-template-form"]["assignment-extension-reason"].value;
    data.topicName = document.forms["assignment-template-form"]["topic-clarification"].value;
    data.topicQuestion = document.forms["assignment-template-form"]["topic-quesion"].value;
    sendMessage(JSON.stringify(data));
  }
})();
