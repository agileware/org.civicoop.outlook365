(function(){
  'use strict';

  // The Office initialize function must be run each time a new page is loaded.
  Office.initialize = function(reason){

    var myLanguage = Office.context.displayLanguage;
    var UIText = UIStrings.getLocaleStrings(myLanguage);
    var config = null;

    jQuery(document).ready(async function(){
      $('h1.title').text(UIText.SaveContactInGroupScreen.Title);
      $('.not-configured-warning .ms-MessageBar-text').html(UIText.SettingsScreen.NotConfigured);
      $('#group-select-done .ms-Button-label').text(UIText.SaveContactInGroupScreen.Done);
      config = await getConfig();
      await getGroups(config)

      


      var DropdownHTMLElements = document.querySelectorAll('.ms-Dropdown');
      for (var i = 0; i < DropdownHTMLElements.length; ++i) {
        var Dropdown = new fabric['Dropdown'](DropdownHTMLElements[i]);
      }

      // Check if warning should be displayed.
      // var contact = JSON.parse(getParameterByName('config'));
      // console.log(contact)
      // if (contact) {
      //   $('#civicrm-name').val(contact.name);
      //   $('#civicrm-email').val(contact.email);
      // }
      // change();

      // // When the Done button is selected, send the
      // // values back to the caller as a serialized
      // // object.
      $('#group-select-done').on('click', done);
    });
  };

  async function getGroups(config){
    var url = config.url + '?';
      var data = {
        "entity": "Group",
        "action": "get",
        "api_key": config.apikey,
        "key": config.sitekey,
        "json": {
          "sequential": 1,
          "return": ["id","name"],
          "options": {
            "limit": 0,
          }
        }
      };
      for(var prop in data) {
        if (prop == 'json') {
          url = url + '&' + prop + '=' + encodeURI(JSON.stringify(data[prop]));
        } else {
          url = url + '&' + prop + '=' + data[prop];
        }
      } 
      await $.getJSON(url, {}, addGroups);
  }

  function addGroups(data){
    console.log(data)
    for(var key in data['values']){
      $('.group-select-dropdown').append('<option value="' + data['values'][key]["id"] + '">' + data['values'][key]['name'] + '</option>'); 
    }
  }

  function done() {
    let selectedGroup = $(".group-select-dropdown").val()
    let fieldName = $("#civicrm-group-field").val()
    let exist = true
    if(fieldName!==""){
      exist = false
    }
    Office.context.ui.messageParent(JSON.stringify({"exist":exist,"selectedGroup":selectedGroup,"groupName":fieldName}));
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

  function getConfig() {
      var config = {};

      config.url = Office.context.roamingSettings.get('civicrm_url');
      config.sitekey = Office.context.roamingSettings.get('civicrm_sitekey');
      config.apikey = Office.context.roamingSettings.get('civicrm_apikey');
      if (config.url && config.apikey && config.sitekey) {
        return config;
      }
      return null;
  }
})();
