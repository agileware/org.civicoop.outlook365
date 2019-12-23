(function() {
  'use strict';
  Office.initialize = function (reason) {
    var item = Office.context.mailbox.item;
    var currentOffset = 0;
    var moreAvailable = true;
    var search = false;
    var config = false;

    var myLanguage = Office.context.displayLanguage;
    var UIText = UIStrings.getLocaleStrings(myLanguage);

    var fields = [];


    jQuery(document).ready(function() {
      // Set localized text for UI elements.
      $("#loadingContacts .ms-Spinner-label").text(UIText.Loading);
      $("#searchField .ms-SearchBox-text").text(UIText.Search);
      $('#settings-prompt').html(UIText.NotConfigured);
      $('#settings-icon .label').text(UIText.Settings);

      reset();
      console.log(fields);
      if (config) {
        loadNextContacts();
      } else {
        openSettingsDialog();
      }

      $('#settings-icon').on('click', openSettingsDialog);
    });

    $('#searchField').on("keypress", function(e) {
      if (e.keyCode == 13) {
        reset();
        search = $(this).val();
        loadNextContacts();
        return false; // prevent the button click from happening
      }
    });

    $(window).scroll(function() {
      if($(window).scrollTop() == $(document).height() - $(window).height()) {
        // ajax call get data from server and append to the div
        loadNextContacts();
      }
    });

    /**
     * Reset the contact list
     */
    function reset() {
      config = getConfig();
      if (!config) {
        $('.not-configured-warning').show();
        $('#search-form').hide();
      } else {
        $('.not-configured-warning').hide();
        $('#search-form').show();
      }
      $('#loadingContacts').hide();
      $('#contacts').html('');
      currentOffset = 0;
      moreAvailable = true;
      search = false;
      fields = [];
      getFields();
    }

    /**
     * Retrieve next batch of contacts.
     */
    function getFields() {
      if (!config) {
        return;
      }
      $('#loadingContacts').show();
      var url = config.url + '?';
      var data = {
        "entity": "Outlook365Contact",
        "action": "getfields",
        "api_key": config.apikey,
        "key": config.sitekey,
        "json": {
          "sequential": 1,
          "api_action": "get"
        }
      };
      for(var prop in data) {
        if (prop == 'json') {
          url = url + '&' + prop + '=' + JSON.stringify(data[prop]);
        } else {
          url = url + '&' + prop + '=' + data[prop];
        }
      }
      $.getJSON(url, {}, function(result) {
        console.log(result);
        fields = [];
        if (!result.is_error) {
          for (var i in result.values) {
            var field = result.values[i];
            if (field.name != 'display_name') {
              fields.push({
                "name": field.name,
                "title": field.title
              });
            }
          }
          console.log(fields);
        }
      });
    }

    /**
     * Retrieve next batch of contacts.
     */
    function loadNextContacts() {
      if (!moreAvailable || !config) {
        return;
      }
      $('#loadingContacts').show();
      var url = config.url + '?';
      var data = {
        "entity": "Outlook365Contact",
        "action": "get",
        "api_key": config.apikey,
        "key": config.sitekey,
        "json": {
          "sequential": 1,
          "options": {
            "offset": currentOffset,
            "limit": 25,
          }
        }
      };
      if (search) {
        data.json.display_name = {"LIKE": '%'+search+'%'};
      }
      for(var prop in data) {
        if (prop == 'json') {
          url = url + '&' + prop + '=' + JSON.stringify(data[prop]);
        } else {
          url = url + '&' + prop + '=' + data[prop];
        }
      }
      $.getJSON(url, {}, addContacts);
    }

    /**
     * Add the contacts to the list.
     *
     * @param data
     */
    function addContacts(data) {
      $('#loadingContacts').hide();
      if (data.is_error == 0) {
        for(var i in data.values) {
          currentOffset ++;
          var contact = data.values[i];
          var name = contact.display_name;
          var email = '';
          var buttons = '';
          if (contact.email) {
            buttons =
              '<button class="ms-Button ms-Button--small to"><span class="ms-Button-label">'+UIText.To+'</span></button>' +
              '<button class="ms-Button ms-Button--small cc"><span class="ms-Button-label">'+UIText.Cc+'</span></button>' +
              '<button class="ms-Button ms-Button--small bcc"><span class="ms-Button-label">'+UIText.Bcc+'</span></button>';
            buttons = '<div class="CiviCRM-Email" data-civicrm-name="'+contact.display_name+'" data-civicrm-email="'+contact.email+'">'+buttons+'</div>';
          }

          var secondaryFields = '';
          for(var fieldI in fields) {
            var fieldName = fields[fieldI].name;
            var value = contact[fieldName];
            if (value) {
              secondaryFields = secondaryFields + '<div class="ms-Persona-secondaryText"><strong>' + fields[fieldI].title + ':</strong>&nbsp;' + value + '</div>';
            }
          }

          var html = '' +
            '<div class="ms-Persona">'+
            '<div class="ms-Persona-details">' +
            '<div class="ms-Persona-primaryText">' + name + '</div>' +
            '<div class="ms-Persona-secondaryText">' + email + '</div>' +
            secondaryFields +
            buttons +
            '</div>' +
            '</div>';
          $('#contacts').append(html);
        }

        $("#contacts .ms-Button.to").click(function() {
          var email = $(this).parent('.CiviCRM-Email').data('civicrm-email');
          var name = $(this).parent('.CiviCRM-Email').data('civicrm-name');
          var recipients = item.to;
          addReiever(recipients, email, name);
        });
        $("#contacts .ms-Button.cc").click(function() {
          var email = $(this).parent('.CiviCRM-Email').data('civicrm-email');
          var name = $(this).parent('.CiviCRM-Email').data('civicrm-name');
          var recipients = item.cc;
          addReiever(recipients, email, name);
        });
        $("#contacts .ms-Button.bcc").click(function() {
          var email = $(this).parent('.CiviCRM-Email').data('civicrm-email');
          var name = $(this).parent('.CiviCRM-Email').data('civicrm-name');
          var recipients = item.bcc;
          addReiever(recipients, email, name);
        });
      }
      if (data.count < 25) {
        moreAvailable = false;
      }
    }

    /**
     * Add the contact to the To, CC or BCC part of the compose e-mail.
     *
     * @param recipients
     * @param email
     * @param name
     */
    function addReiever(recipients, email, name) {
      var data = [
        {
          "displayName": name,
          "emailAddress": email
        }
      ];
      // Use asynchronous method getAsync to get each type of recipients
      // of the composed item. Each time, this example passes an anonymous
      // callback function that doesn't take any parameters.
      recipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
          write(asyncResult.error.message);
        }
        else {
          // Async call to get to-recipients of the item completed.
          // Display the email addresses of the to-recipients.
          var emailAlreadyExists = false;
          for (var i=0; i<asyncResult.value.length; i++) {
            if (asyncResult.value[i].emailAddress == email) {
              emailAlreadyExists = true;
              break;
            }
          }
          if (!emailAlreadyExists) {
            recipients.addAsync(data, function (asyncResult) {});
          }
        }
      });
    }

    /**
     * Function to open the settings dialog.
     */
    function openSettingsDialog() {
      // Display settings dialog.
      var url = new URI('../settings/dialog.html').absoluteTo(window.location).toString();
      if (config) {
        // If the add-in has already been configured, pass the existing values
        // to the dialog.
        url = url + '?config='+JSON.stringify(config);
      }

      var dialogOptions = { width: 30, height: 40, displayInIframe: true };

      Office.context.ui.displayDialogAsync(url, dialogOptions, function(result) {
        var settingsDialog = result.value;
        settingsDialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, function(message){
          config = JSON.parse(message.message);
          setConfig(config, function(result) {
            settingsDialog.close();
            settingsDialog = null;
            reset();
            loadNextContacts();
          });
        });
      });
    }

    /**
     * Load the configuration.
     *
     * @returns {null}
     */
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

    /**
     * Save the configuration.
     *
     * @param config
     * @param callback
     */
    function setConfig(config, callback) {
      Office.context.roamingSettings.set('civicrm_url', config.url);
      Office.context.roamingSettings.set('civicrm_sitekey', config.sitekey);
      Office.context.roamingSettings.set('civicrm_apikey', config.apikey);

      Office.context.roamingSettings.saveAsync(callback);
    }

  };

})();
