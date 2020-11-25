(function () {
  'use strict';

  // The Office initialize function must be run each time a new page is loaded.
  Office.initialize = function (reason) {

    var myLanguage = Office.context.displayLanguage;
    var UIText = UIStrings.getLocaleStrings(myLanguage);

    jQuery(document).ready(function () {
      $('h1.title').text("Confirm save.");
      $('.ms-MessageBar-text').html("Are you sure you want to save all contacts ?");
      $('#confirm-done .ms-Button-label').text("Save");
      $('#confirm-cancel .ms-Button-label').text("Cancel");

      // When the Done button is selected, send the
      // values back to the caller as a serialized
      // object.
      $('#confirm-done').on('click', done);
      $('#confirm-cancel').on('click', cancel);
    });
  };

  function done() {
    Office.context.ui.messageParent(JSON.stringify({"action": true}));
  }

  function cancel() {
    Office.context.ui.messageParent(JSON.stringify({"action": false}));
  }

})();
