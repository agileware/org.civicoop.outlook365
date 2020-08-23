(function(){
  'use strict';

  // The Office initialize function must be run each time a new page is loaded.
  Office.initialize = function(reason){

    var myLanguage = Office.context.displayLanguage;
    var UIText = UIStrings.getLocaleStrings(myLanguage);

    jQuery(document).ready(function(){
      $('h1.title').text(UIText.SaveContactScreen.Title);
      $('.not-configured-warning .ms-MessageBar-text').html(UIText.SettingsScreen.NotConfigured);
      $('.ms-Label.save-contact-name').text(UIText.SaveContactScreen.ContactName);
      $('.ms-Label.save-contact-email').text(UIText.SaveContactScreen.ContactEmail);
      $('#settings-done-contact .ms-Button-label').text(UIText.SaveContactScreen.Done);

      $('#civicrm-name').change(change);
      $('#civicrm-email').change(change);

      // Check if warning should be displayed.
      var contact = JSON.parse(getParameterByName('config'));
      console.log(contact)
      if (contact) {
        $('#civicrm-name').val(contact.name);
        $('#civicrm-email').val(contact.email);
      }
      change();

      // When the Done button is selected, send the
      // values back to the caller as a serialized
      // object.
      $('#settings-done-contact').on('click', done);
    });
  };

  function change() {
    var contact = {};
    contact.name = $('#civicrm-name').val();
    contact.email = $('#civicrm-email').val();
    $('#settings-done-contact').prop('disabled', false);
  }

  function done() {
    var contact = {};
    contact.name = $('#civicrm-name').val();
    contact.email = $('#civicrm-email').val();
    Office.context.ui.messageParent(JSON.stringify(contact));
  }

  function getParameterByName(name, url) {
    if (!url) {
      url = window.location.href;
    }
    name = name.replace(/[\[\]]/g, "\\$&");
    var regex = new RegExp("[?&]" + name + "(=([^&#]*)|&|#|$)"),
      results = regex.exec(url);
    if (!results) return null;
    if (!results[2]) return '';
    return decodeURIComponent(results[2].replace(/\+/g, " "));
  }
})();
