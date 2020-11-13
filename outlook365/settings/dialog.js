(function(){
  'use strict';

  // The Office initialize function must be run each time a new page is loaded.
  Office.initialize = function(reason){

    var myLanguage = Office.context.displayLanguage;
    var UIText = UIStrings.getLocaleStrings(myLanguage);
    var accessToken;

    jQuery(document).ready(function(){
      $('h1.title').text(UIText.SettingsScreen.Title);
      $('.not-configured-warning .ms-MessageBar-text').html(UIText.SettingsScreen.NotConfigured);
      $('.ms-Label.url').text(UIText.SettingsScreen.URL);
      $('.ms-Label.sitekey').text(UIText.SettingsScreen.SiteKey);
      $('.ms-Label.apikey').text(UIText.SettingsScreen.ApiKey);
      $('#settings-done .ms-Button-label').text(UIText.SettingsScreen.Done);
      $('#civicrm-url').attr("placeholder", UIText.SettingsScreen.URL_Placeholder);



      $('#civicrm-url').change(change);
      $('#site-key').change(change);
      $('#api-key').change(change);

      // Check if warning should be displayed.
      var config = JSON.parse(getParameterByName('config'));
      if (!config) {
        $('.not-configured-warning').show();
      } else {
        $('#civicrm-url').val(config.url);
        $('#site-key').val(config.sitekey);
        $('#api-key').val(config.apikey);
      }

      change();

      // When the Done button is selected, send the
      // values back to the caller as a serialized
      // object.
      $('#settings-done').on('click', done);
    });
  };

  function change() {
    var settings = {};
    settings.url = $('#civicrm-url').val();
    settings.sitekey = $('#site-key').val();
    settings.apikey = $('#api-key').val();
    if (settings.url && settings.sitekey && settings.apikey) {
      $('#settings-done').prop('disabled', false);
    } else {
      $('#settings-done').prop('disabled', true);
    }
  }

  function done() {
    var settings = {};
    settings.url = $('#civicrm-url').val();
    settings.sitekey = $('#site-key').val();
    settings.apikey = $('#api-key').val();
    Office.context.ui.messageParent(JSON.stringify(settings));
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
