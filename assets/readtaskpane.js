(function () {
  'use strict';
  Office.initialize = function (reason) {

    var requiredAttendees = Office.context.mailbox.item;
    var config = false;
    var saveDialog = null
    var accessToken = null;

    jQuery(document).ready(async function () {
      // Set localized text for UI elements.
      await reset();

      Office.context.mailbox.getCallbackTokenAsync({isRest: true}, function (result) {
        accessToken = result.value;
        var getMessageUrl = Office.context.mailbox.restUrl + '/v2.0/me/MailFolders';

        $.ajax({
          url: getMessageUrl,
          dataType: 'json',
          headers: {'Authorization': 'Bearer ' + accessToken}
        }).done(function (data) {
          // Message is passed in `item`.
          $('#target').empty();
          console.log(data);
          for (const item of data.value) {
            let html = $(`<li class="ms-ListItem is-selectable" name="${item.Id}" id="${item.Id}" tabindex="0">` +
              `<span class="ms-ListItem-secondaryText">${item.DisplayName}</span>` +
              '<div class="ms-ListItem-selectionTarget"></div>' +
              '</li>');
            html.appendTo('#target');
          }
          var ListItemElements = document.querySelectorAll(".ms-ListItem");
          for (var i = 0; i < ListItemElements.length; i++) {
            new fabric['ListItem'](ListItemElements[i]);
          }
        }).fail(function (error) {
          // Handle error.
          console.log(error);
        });

        $("#send-submit").on('click', saveEmailsInFolder);
      });

      // fetch all folders

      if (config) {
        showContacts();
      } else {
        openSettingsDialog();
      }

      $('#settings-icon').on('click', openSettingsDialog);

      $("body").delegate(".selectAll", "click", function (event) {
        $(".ms-ListItem").addClass("is-selected")
      });

      $("body").delegate(".unselectAll", "click", function (event) {
        $(".ms-ListItem").removeClass("is-selected")
      });

      $("body").delegate(".ms-ListItem", "click", function (event) {
        if ($(event.target).hasClass('is-selected')) {
          $(event.target).removeClass('is-selected');
        } else {
          $(event.target).addClass('is-selected');
        }
      });

    });

    /**
     * An event handler for save emails in the selected folders
     * @param event
     */
    function saveEmailsInFolder(event) {
      $(this).prop('disabled', true);
      savingEmailInFolderInfo();
      $('#folder-form li.is-selected').each((index, element) => {
        let folderID = element.getAttribute('name');
        let getMessageUrl = Office.context.mailbox.restUrl +
          '/v2.0/me/MailFolders/' + folderID + '/messages?$expand=SingleValueExtendedProperties';
        $.ajax({
          url: getMessageUrl,
          dataType: 'json',
          headers: {'Authorization': 'Bearer ' + accessToken}
        }).done(async function (data) {
          console.log(data);
          // Message is passed in `item`.
          if (!data.value.length) return;
          for (const email of data.value) {
            if (email.Categories.includes('Saved in CiviCRM')) {
              continue;
            }
            $.ajax({
              url: Office.context.mailbox.restUrl + "/v2.0/me/messages/" + email.Id,
              dataType: 'json',
              contentType: 'application/json',
              method: 'PATCH',
              headers: {'Authorization': 'Bearer ' + accessToken},
              data: JSON.stringify({
                Categories: [
                  "Saved in CiviCRM"
                ]
              })
            }).done(result => {
              console.log(result);
            });
            await pushEmailActivity(email.Subject, email.Body.Content, new Date());
          }
          emailInFolderSavedInfo();
        }).fail(function (error) {
          // Handle error.
          console.log(error);
        });
      });
    }

    function savingEmailInFolderInfo() {
      $('#saving-email-notice').show();
      $('#saved-email-notice').hide();
    }

    function emailInFolderSavedInfo() {
      $('#saving-email-notice').hide();
      $('#saved-email-notice').show();
    }

    async function pushEmailActivity(subject, body, date, from, to) {
      let emailData = {
        "source_contact_id": "user_contact_id",
        "activity_type_id": "Email",
        "subject": subject,
        "details": body,
      }

      let url = config.url + '?';
      let data = {
        "entity": "Activity",
        "action": "create",
        "api_key": config.apikey,
        "key": config.sitekey,
        "json": 1
      };
      for (var prop in data) {
        if (prop == 'json') {
          url = url + '&' + prop + '=' + JSON.stringify(data[prop]);
        } else {
          url = url + '&' + prop + '=' + data[prop];
        }
      }

      await $.post(url, emailData, function (result, status) {
        console.log(result)
      }).fail(response => {
        openDialog(dialogComponent);
      });
    }

    async function reset() {
      config = await getConfig();
      if (!config) {
        $('.not-configured-warning').show();
        $('#search-form').hide();
      } else {
        $('.not-configured-warning').hide();
      }
      $('#loadingContacts').hide();
    }

    async function openSettingsDialog() {
      // Display settings dialog.
      var url = settingsDialogUrl;
      if (config) {
        // If the add-in has already been configured, pass the existing values
        // to the dialog.
        url = url + '?config=' + JSON.stringify(config);
      }

      var dialogOptions = {width: 30, height: 40, displayInIframe: true};

      Office.context.ui.displayDialogAsync(url, dialogOptions, function (result) {
        var settingsDialog = result.value;
        settingsDialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, async function (message) {
          config = JSON.parse(message.message);
          await setConfig(config, function (result) {
            settingsDialog.close();
            settingsDialog = null;
            reset();
            showContacts()
          });
        });
      });
    }

    async function getListItem(item, res) {
      let html = ''
      let name
      let class_to_add
      let contact_id
      if (res['exist']) {
        name = res['contact_name']
        class_to_add = "not_to_save"
        contact_id = res['contact_id']
      } else {
        name = item.displayName
        class_to_add = "to_save"
      }
      html += '<li class="ms-ListItem is-selectable ' + class_to_add + '">' +
        '<span class="ms-ListItem-secondaryText">' + name + '</span>' +
        '<span class="ms-ListItem-tertiaryText">' + item.emailAddress + '</span>' +
        '<div class="ms-ListItem-actions">' +
        '<div class="ms-ListItem-action" data-civicrm-name="' + name + '" data-civicrm-email="' + item.emailAddress + '"'

      if (res['exist']) {
        html += ' data-civicrm-id="' + String(contact_id) + '">'
      } else {
        html += '>';
      }

      if (res['exist']) {
        html += '<a href="' + res.contact_url + '" target="_blank"><i class="ms-Icon ms-Icon--Contact" title="View Contact in CiviCRM"></i></a>'
      } else {
        html += '<i class="ms-Icon ms-Icon--Save save-contact" title="Save Contact to CiviCRM"></i>'
      }

      html += '</div>' +
        '</div>' +
        '</li>'
      return html
    }

    async function saveContactToCRM(contact) {
      let contactName = splitContactName(contact.name);
      var url = config.url + '?';
      var data = {
        "entity": "Contact",
        "action": "create",
        "api_key": config.apikey,
        "key": config.sitekey,
        "json": {
          "first_name": contactName.firstName,
          'last_name': contactName.lastName,
          "contact_type": 'Individual',
        }
      };
      for (var prop in data) {
        if (prop == 'json') {
          url = url + '&' + prop + '=' + encodeURI(JSON.stringify(data[prop]));
        } else {
          url = url + '&' + prop + '=' + data[prop];
        }
      }
      var contact_info
      await $.post(url, function (result, status) {
        contact_info = result
      }).fail(response => {
        openDialog(dialogComponent);
      });
      var contact_id = contact_info.id

      url = config.url + '?';
      data = {
        "entity": "Email",
        "action": "create",
        "api_key": config.apikey,
        "key": config.sitekey,
        "json": {
          "contact_id": contact_id,
          "email": contact.email,
        }
      };
      for (var prop in data) {
        if (prop == 'json') {
          url = url + '&' + prop + '=' + JSON.stringify(data[prop]);
        } else {
          url = url + '&' + prop + '=' + data[prop];
        }
      }
      await $.post(url, function (result, status) {
        console.log(result)
      }).fail(response => {
        openDialog(dialogComponent);
      });
      return contact_id
    }

    async function saveContact(event) {
      $(this).prop('disabled', true);
      showSavingContactInfo();
      let name = $(event.target).parent().data('civicrm-name')
      let email = $(event.target).parent().data('civicrm-email');
      let contact = {name: name, email: email};
      await saveContactToCRM(contact);
      $(this).prop('disabled', false);
      showContactSavedInfo();
      showContacts();
    }

    async function confirmSaveAllContact(event) {
      $(this).prop('disabled', true);
      showSavingContactInfo();
      await saveAllContact();
      $(this).prop('disabled', false);
      showContactSavedInfo();
      location.reload();
    }

    function showSavingContactInfo() {
      $('#saving-contact-help-text').toggle(true);
      $('#saved-contact-help-text').toggle(false);
    }

    function showContactSavedInfo() {
      $('#saving-contact-help-text').toggle(false);
      $('#saved-contact-help-text').toggle(true);
    }

    async function saveAllContact(event) {
      let toSave = []
      $('.ms-List').children().each(function (index) {
        if ($(this).hasClass('to_save') && $(this).hasClass('is-selected')) {
          let contact = $(this).children(".ms-ListItem-actions").children(".ms-ListItem-action")
          toSave.push([contact.data('civicrm-name'), contact.data('civicrm-email')])
        }
      });


      for (const [key, val] of Object.entries(toSave)) {
        await saveContactToCRM({"name": val[0], "email": val[1]})
      }
    }

    async function confirmSaveAllContactInGroup(event) {
      let toSave = [];
      $('.ms-List').children().each(function (index) {
        let contact = $(this).children(".ms-ListItem-actions").children(".ms-ListItem-action")
        // name,email,already saved
        if ($(this).hasClass('is-selected')) {
          if ($(this).hasClass('to_save')) {
            toSave.push([contact.data('civicrm-name'), contact.data('civicrm-email'), $(this).hasClass('to_save')])
          } else {
            toSave.push([contact.data('civicrm-name'), contact.data('civicrm-email'), $(this).hasClass('to_save'), contact.data('civicrm-id')])
          }
        }
      });
      if (toSave.length == 0) {
        console.log("None selected")
        return
      }
      var dialogOptions = {width: 30, height: 40, displayInIframe: true};
      var url = saveContactInGroupDialogUrl;
      Office.context.ui.displayDialogAsync(url, dialogOptions, async function (result) {
        var confirmDialog = result.value;
        confirmDialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, async function (message) {
          let groupInfo = JSON.parse(message.message);
          console.log(groupInfo)
          var groupId
          var groupResult
          if (groupInfo['exist'] === true) {
            groupId = groupInfo["selectedGroup"]
          } else {
            groupResult = await addGroupToCRM(groupInfo['groupName'])
            groupId = groupResult["id"]
          }

          for (var key in toSave) {
            if (toSave[key][2] === true) {
              console.log("tosave in CRM")
              let contactId = await saveContactToCRM({"name": toSave[key][0], "email": toSave[key][1]})
              await saveContactGroup(contactId, groupId)
            } else {
              console.log("nottosave in CRM")
              await saveContactGroup(toSave[key][3], groupId)
            }
          }
          confirmDialog.close();
          confirmDialog = null;
          location.reload();
        });

      });

    }

    async function saveContactGroup(contactId, groupId) {
      var url = config.url + '?';
      var data = {
        "entity": "GroupContact",
        "action": "create",
        "api_key": config.apikey,
        "key": config.sitekey,
        "json": {
          "group_id": groupId,
          "contact_id": contactId,
        }
      };
      for (var prop in data) {
        if (prop == 'json') {
          url = url + '&' + prop + '=' + encodeURI(JSON.stringify(data[prop]));
        } else {
          url = url + '&' + prop + '=' + data[prop];
        }
      }
      await $.post(url, function (result, status) {
        console.log(result)
      }).fail(response => {
        openDialog(dialogComponent);
      });

    }

    async function addGroupToCRM(name) {
      var url = config.url + '?';
      var data = {
        "entity": "Group",
        "action": "create",
        "api_key": config.apikey,
        "key": config.sitekey,
        "json": {
          "title": name,
          "name": name,
        }
      };
      for (var prop in data) {
        if (prop == 'json') {
          url = url + '&' + prop + '=' + JSON.stringify(data[prop]);
        } else {
          url = url + '&' + prop + '=' + data[prop];
        }
      }
      let groupInfo
      await $.post(url, function (result, status) {
        groupInfo = result
      }).fail(response => {
        openDialog(dialogComponent);
      });

      return groupInfo
    }

    async function checkContact(val) {
      var url = config.url + '?';
      var data = {
        "entity": "Email",
        "action": "get",
        "api_key": config.apikey,
        "key": config.sitekey,
        "json": {
          "email": val.emailAddress,
        }
      };
      for (var prop in data) {
        if (prop == 'json') {
          url = url + '&' + prop + '=' + JSON.stringify(data[prop]);
        } else {
          url = url + '&' + prop + '=' + data[prop];
        }
      }
      var contact_info
      var exist = null
      var contact_url = CRMContactURL
      var contact_name = null
      var contact_id
      for (const [key, val] of Object.entries(config.url.split("/"))) {
        if (val == "sites") {
          break
        }
      }
      await $.post(url, function (result, status) {
        exist = false
        if (result["count"] > 0) {
          exist = true
          let keys = Object.keys(result["values"])
          contact_id = result["values"][keys[0]]["contact_id"]
          contact_url += "&cid=" + String(contact_id)
        }
      }).fail(response => {
        openDialog(dialogComponent);
      });
      if (exist) {
        url = config.url + '?'
        data = {
          "entity": "Contact",
          "action": "getsingle",
          "api_key": config.apikey,
          "key": config.sitekey,
          "json": {
            "id": contact_id,
          }
        };
        for (var prop in data) {
          if (prop == 'json') {
            url = url + '&' + prop + '=' + JSON.stringify(data[prop]);
          } else {
            url = url + '&' + prop + '=' + data[prop];
          }
        }
        await $.post(url, function (result, status) {
          contact_name = result["display_name"]
        }).fail(response => {
          openDialog(dialogComponent);
        });
      }

      return {"exist": exist, "contact_url": contact_url, "contact_name": contact_name, "contact_id": contact_id}
    }

    async function showContacts() {

      let html = '<ul class="ms-List">'

      //adding sender
      let res = await checkContact(requiredAttendees.from)
      html += await getListItem(requiredAttendees.from, res)

      //adding to
      for (const [key, val] of Object.entries(requiredAttendees.to)) {
        let res = await checkContact(val)
        html += await getListItem(val, res)
      }

      //adding CC
      for (const [key, val] of Object.entries(requiredAttendees.cc)) {
        let res = await checkContact(val)
        html += await getListItem(val, res)
      }

      //adding BCC
      for (const [key, val] of Object.entries(requiredAttendees.bcc)) {
        let res = await checkContact(val)
        html += await getListItem(val, res)
      }
      html += '</ul>'

      html += '<button class="ms-Button ms-Button--small save-contact-all">' +
        '<span class="ms-Button-label">Save Contacts</span>' +
        '</button>'
      html += '<button class="ms-Button ms-Button--small save-contact-all-group">' +
        '<span class="ms-Button-label">Save Contacts to Group</span>' +
        '</button>'
      html += '<p id="saving-contact-help-text" style="display: none;">Contacts are being saved to CiviCRM, please wait...</p>';
      html += '<p id="saved-contact-help-text" style="display: none;">All selected contacts have been saved to CiviCRM</p>';

      html += '<br><br>';


      $("#contacts").html(html)
      $(".save-contact").on('click', saveContact)
      $(".save-contact-all-group").on('click', confirmSaveAllContactInGroup)
      $(".save-contact-all").on('click', confirmSaveAllContact)
      $(".save-email").on('click', saveEmail)

    }

    async function saveEmail() {
      var emailBody = null
      var datetime = new Date(requiredAttendees.dateTimeCreated)
      var hours = String(datetime.getHours())
      var seconds = String(datetime.getSeconds())
      var minutes = String(datetime.getMinutes())
      datetime = datetime.getFullYear() + "-" + (datetime.getMonth() + 1) + "-" + datetime.getDate()

      datetime = datetime + " " + hours.padStart(2, "0") + ":" + minutes.padStart(2, "0") + ":" + seconds.padStart(2, "0")
      await requiredAttendees.body.getAsync(
        "html",
        async function callback(result) {
          emailBody = result['value']

          var data = {
            "source_contact_id": "user_contact_id",
            "activity_type_id": "Email",
            "subject": requiredAttendees.subject,
            "details": emailBody,
            "activity_date_time": datetime.toString(),
            "target_id": [],
          }

          $('.ms-List').children().each(function (index) {
             let contact = $(this).children(".ms-ListItem-actions").children(".ms-ListItem-action")
             // name,email,already saved
             if ($(this).hasClass('is-selected')) {
               if ($(this).hasClass('to_save')) {
                 let contactId = saveContactToCRM({"name": contact.data('civicrm-name'), "email": contact.data('civicrm-email')})
                 data.target_id.push(contactId);
               } else {
                 data.target_id.push(contact.data('civicrm-id'));
               }
             }
          });
      
          var url = config.url + '?'
          var param = {
            "entity": "Activity",
            "action": "create",
            "api_key": config.apikey,
            "key": config.sitekey,
            "json": 1
          };
          for (var prop in param) {
            if (prop == 'json') {
              url = url + '&' + prop + '=' + JSON.stringify(param[prop]);
            } else {
              url = url + '&' + prop + '=' + param[prop];
            }
          }
          await $.post(url, data, function (result, status) {
            console.log(result);
            addSavedCategoryToCurrentMessage();
          }).fail(response => {
            openDialog(dialogComponent);
          });
        });
    }

    function addSavedCategoryToCurrentMessage() {
      // skip if the message has sent to CiviCRM
      Office.context.mailbox.item.categories.getAsync(function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.log("Action failed with error: " + asyncResult.error.message);
        } else {
          var categories = asyncResult.value;
          var hasSaved = false;
          categories.forEach(function (item) {
            if (item.displayName === 'Saved in CiviCRM') {
              hasSaved = true;
            }
          });
          if (!hasSaved) {
            Office.context.mailbox.item.categories.addAsync(['Saved in CiviCRM'], result => {
              if (result.status === Office.AsyncResultStatus.Succeeded) {
                console.log("Successfully added categories");
              } else {
                console.log("categories.addAsync call failed with error: " + result.error.message);
                // create the category if it is not exist and try again
                if (result.error.code === 9044) {
                  var masterCategoriesToAdd = [
                    {
                      "displayName": "Saved in CiviCRM",
                      "color": Office.MailboxEnums.CategoryColor.Preset0
                    }
                  ];

                  Office.context.mailbox.masterCategories.addAsync(masterCategoriesToAdd, function (asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                      console.log("Successfully added categories to master list");
                      addSavedCategoryToCurrentMessage();
                    } else {
                      console.log("masterCategories.addAsync call failed with error: " + asyncResult.error.message);
                    }
                  });
                }
              }
            });
          }
        }
      });

    }

    async function getConfig() {
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
    async function setConfig(config, callback) {
      await Office.context.roamingSettings.set('civicrm_url', config.url);
      await Office.context.roamingSettings.set('civicrm_sitekey', config.sitekey);
      await Office.context.roamingSettings.set('civicrm_apikey', config.apikey);
      Office.context.roamingSettings.saveAsync(callback);
    }

    function splitContactName(name) {
      let separatorIndex = name.indexOf(' ');
      if (separatorIndex === -1) {
        separatorIndex = name.indexOf(',');
      }
      let contact = {};
      if (separatorIndex === -1) {
        contact.lastName = 'Unknown';
      } else {
        contact.lastName = name.substring(separatorIndex + 1);
      }
      contact.firstName = name.substring(0, separatorIndex);
      return contact;
    }
  }
})();
