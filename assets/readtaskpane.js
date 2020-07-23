(function() {
  'use strict';
  Office.initialize = function (reason) {
        var requiredAttendees = Office.context.mailbox.item;
        var config = false;

        jQuery(document).ready(function() {
          // Set localized text for UI elements.
          reset();
          if (config) {
            showContacts();
          } else {
            openSettingsDialog();
          }

          $('#settings-icon').on('click', openSettingsDialog);

        });

        function reset() {
          config = getConfig();
          if (!config) {
            $('.not-configured-warning').show();
            $('#search-form').hide();
          } else {
            $('.not-configured-warning').hide();
          }
          $('#loadingContacts').hide();
        }

        function openSettingsDialog() {
          // Display settings dialog.
          var url = settingsDialogUrl;
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
                showContacts()
              });
            });
          });
        }

        function getListItem(item){
            let html =''
            html += '<li class="ms-ListItem">'+
                    '<span class="ms-ListItem-secondaryText">' + item.displayName +'</span>'+
                    '<span class="ms-ListItem-tertiaryText">' + item.emailAddress +'</span>'+
                    '<div class="ms-ListItem-actions">'+
                        '<div class="ms-ListItem-action" data-civicrm-name="' + item.displayName+'" data-civicrm-email="' + item.emailAddress+'" >'+
                        '<i class="ms-Icon ms-Icon--Save save-contact"></i>'+
                        '</div>'+
                    '</div>'+
                    '</li>'
            return html
        }

        function saveContactToCRM(contact){
            console.log(contact)
            var url = config.url + '?';
            var data = {
                "entity": "Contact",
                "action": "create",
                "api_key": config.apikey,
                "key": config.sitekey,
                "json": {
                    "display_name":contact.name,
                    "contact_type":"Individual",
                }
              };
            for(var prop in data) {
                if (prop == 'json') {
                  url = url + '&' + prop + '=' + JSON.stringify(data[prop]);
                } else {
                  url = url + '&' + prop + '=' + data[prop];
                }
            }
            $.post(url, function(result) {
                console.log(result)
            })
        }

        function saveContact(event) {
            let name = $(event.target).parent().data('civicrm-name')
            let email = $(event.target).parent().data('civicrm-email')
            var dialogOptions = { width: 30, height: 40, displayInIframe: true };
            let contactObject = {"name":name,"email":email}
            var url = saveContactDialogUrl;
            url = url + '?config='+JSON.stringify(contactObject);
            Office.context.ui.displayDialogAsync(url, dialogOptions, function(result) {
                var settingsDialog = result.value;
                settingsDialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, function(message){
                  let contact = JSON.parse(message.message);
                  saveContactToCRM(contact)
                });
              });
        }


        function showContacts() {

            let html = '<ul class="ms-List">'

            //adding sender
            html += getListItem(requiredAttendees.from)

            //adding to
            for(const [key,val] of Object.entries(requiredAttendees.to)){
                html += getListItem(val)
            }

            //adding CC
            for(const [key,val] of Object.entries(requiredAttendees.cc)){
                html += getListItem(val)
            }

            //adding BCC
            for(const [key,val] of Object.entries(requiredAttendees.bcc)){
                html += getListItem(val)
            }
            html += '</ul>'

            html += '<button class="ms-Button ms-Button--medium save-contact-all">'+
                      '<span class="ms-Button-label">Save all</span>'+
                      '<span class="ms-Button-description">Save all above listed contacts to CiviCRM</span>'+
                    '</button>'


            $("#contacts").html(html)
            $(".save-contact").on('click',saveContact)

            // requiredAttendees.body.getAsync(
            // "text",
            // { asyncContext: "This is passed to the callback" },
            // function callback(result) {
            //     console.log(result)
            // });

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
  }
})();
