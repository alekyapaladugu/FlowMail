(function () {
  'use strict';

  // The Office initialize function must be run each time a new page is loaded.
  Office.initialize = function (reason) {
    registerEvents();
  };

  function registerEvents() {
    document.getElementById("template-form").onsubmit = validateForm;
  }

  function sendMessage(message) {
    Office.context.ui.messageParent(message);
  }

  function validateForm() {
    let data = {}
    data.professorName = document.forms["assignment-template-form"]["professor-name"].value
    data.courseNumber = document.forms["assignment-template-form"]["course-no"].value
    data.assignmentNo = document.forms["assignment-template-form"]["assignment-question-content"].value
    data.studentId = document.forms["assignment-template-form"]["student-id"].value
    data.studentName = document.forms["assignment-template-form"]["student-name"].value

    sendMessage(JSON.stringify(data))
  }

})();