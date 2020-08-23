(function() {
  'use strict';
  Office.initialize = function (reason) {

        var requiredAttendees = Office.context.mailbox.item;
        console.log(Office)
        Office.context.auth.getAccessTokenAsync(function(result) {
            if (result.status === "succeeded") {
                var token = result.value;
                // ...
            } else {
                console.log("Error obtaining token", result.error);
            }
        });
        var config = false;
        var saveDialog = null

        jQuery(document).ready(async function() {
          // Set localized text for UI elements.
          await reset();

          // fetch all folders

          if (config) {
            showContacts();
          } else {
            openSettingsDialog();
          }

          $('#settings-icon').on('click', openSettingsDialog);

          $("body").delegate(".selectAll", "click", function(event){
              $(".ms-ListItem").addClass("is-selected")
          });

          $("body").delegate(".unselectAll", "click", function(event){
              $(".ms-ListItem").removeClass("is-selected")
          });

          $("body").delegate(".ms-ListItem", "click", function(event){
              if($(event.target).hasClass('is-selected')){
                $(event.target).removeClass('is-selected');
              } 
              else{
                $(event.target).addClass('is-selected');
              }
          });

        });

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
            url = url + '?config='+JSON.stringify(config);
          }

          var dialogOptions = { width: 30, height: 40, displayInIframe: true };

          Office.context.ui.displayDialogAsync(url, dialogOptions, function(result) {
            var settingsDialog = result.value;
            settingsDialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, async function(message){
              config = JSON.parse(message.message);
              await setConfig(config, function(result) {
                settingsDialog.close();
                settingsDialog = null;
                reset();
                showContacts()
              });
            });
          });
        }

        async function getListItem(item,res){
            let html =''
            let name
            let class_to_add 
            let contact_id
            if(res['exist']){
              name = res['contact_name']
              class_to_add = "not_to_save"
              contact_id = res['contact_id'] 
            }
            else{
              name = item.displayName
              class_to_add = "to_save"
            }
            html += '<li class="ms-ListItem is-selectable ' + class_to_add +'">'+
                    '<span class="ms-ListItem-secondaryText">' + name +'</span>'+
                    '<span class="ms-ListItem-tertiaryText">' + item.emailAddress +'</span>'+
                    '<div class="ms-ListItem-actions">'+
                        '<div class="ms-ListItem-action" data-civicrm-name="' + name+'" data-civicrm-email="' + item.emailAddress+'"'

            if(res['exist']){
              html+= ' data-civicrm-id="' + String(contact_id) +'">'
            }

            if(res['exist']){
              html+= '<a href="'+res.contact_url+'" ><i class="ms-Icon ms-Icon--Contact"></i></a>'
            }
            else{
              html += '<i class="ms-Icon ms-Icon--Save save-contact"></i>'
            }

            html +='</div>'+
                    '</div>'+
                    '</li>'
            return html
        }

        async function saveContactToCRM(contact){
            var url = config.url + '?';
            var data = {
                "entity": "Contact",
                "action": "create",
                "api_key": config.apikey,
                "key": config.sitekey,
                "json": {
                    "display_name":contact.name,
                    "contact_type":config.contacttype,
                }
              };
            for(var prop in data) {
                if (prop == 'json') {
                  url = url + '&' + prop + '=' + JSON.stringify(data[prop]);
                } else {
                  url = url + '&' + prop + '=' + data[prop];
                }
            }
            var contact_info
            await $.post(url, function(result) {
                contact_info = result
            })
            var contact_id = contact_info.id

            url = config.url + '?';
            data = {
                "entity": "Email",
                "action": "create",
                "api_key": config.apikey,
                "key": config.sitekey,
                "json": {
                    "contact_id":contact_id,
                    "email":contact.email,
                }
              };
            for(var prop in data) {
                if (prop == 'json') {
                  url = url + '&' + prop + '=' + JSON.stringify(data[prop]);
                } else {
                  url = url + '&' + prop + '=' + data[prop];
                }
            }
            await $.post(url, function(result) {
                console.log(result)
            })
            return contact_id
        }

        async function saveContact(event) {
            let name = $(event.target).parent().data('civicrm-name')
            let email = $(event.target).parent().data('civicrm-email')
            var dialogOptions = { width: 30, height: 40, displayInIframe: true };
            let contactObject = {"name":name,"email":email}
            var url = saveContactDialogUrl;
            url = url + '?config='+JSON.stringify(contactObject);
            Office.context.ui.displayDialogAsync(url, dialogOptions,async function(result) {
                var saveDialog = result.value;
                saveDialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, async function(message){
                  let contact = JSON.parse(message.message);
                  await saveContactToCRM(contact)
                  saveDialog.close();
                  saveDialog = null;
                  showContacts()
                });

              });
        }

        async function confirmSaveAllContact(event) {
            var dialogOptions = { width: 30, height: 40, displayInIframe: true };
            var url = confirmDialogUrl;
            Office.context.ui.displayDialogAsync(url, dialogOptions,async function(result) {
                var confirmDialog = result.value;
                confirmDialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, async function(message){
                  let action = JSON.parse(message.message);
                  if(action['action']===true){
                    saveAllContact()
                    confirmDialog.close();
                    confirmDialog = null;
                  }
                });

              });
        }

        async function saveAllContact(event){
          let toSave = []
          $('.ms-List').children().each(function (index) {
            if($(this).hasClass('to_save') && $(this).hasClass('is-selected')){
                let contact = $(this).children(".ms-ListItem-actions").children(".ms-ListItem-action")
                toSave.push([contact.data('civicrm-name'),contact.data('civicrm-email')])
              }
          });


          for(const [key,val] of Object.entries(toSave)){
            await saveContactToCRM({"name":val[0],"email":val[1]})
          }
        }

        async function confirmSaveAllContactInGroup(event){
          let toSave = []
          $('.ms-List').children().each(function (index) {
              let contact = $(this).children(".ms-ListItem-actions").children(".ms-ListItem-action")
              // name,email,already saved
              if($(this).hasClass('is-selected')){
                if($(this).hasClass('to_save')){
                  toSave.push([contact.data('civicrm-name'),contact.data('civicrm-email'),$(this).hasClass('to_save')])
                } else{
                  toSave.push([contact.data('civicrm-name'),contact.data('civicrm-email'),$(this).hasClass('to_save'),contact.data('civicrm-id')])
                }
              }
          });
          if (toSave.length == 0){
            console.log("None selected")
            return
          }
          var dialogOptions = { width: 30, height: 40, displayInIframe: true };
            var url = saveContactInGroupDialogUrl;
            Office.context.ui.displayDialogAsync(url, dialogOptions,async function(result) {
                var confirmDialog = result.value;
                confirmDialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, async function(message){
                let groupInfo = JSON.parse(message.message);
                console.log(groupInfo)
                var groupId 
                var groupResult
                if(groupInfo['exist']===true){
                  groupId = groupInfo["selectedGroup"]
                } else{
                  groupResult = await addGroupToCRM(groupInfo['groupName'])
                  groupId = groupResult["id"]
                }

                for(var key in toSave){
                  if(toSave[key][2]===true){
                    console.log("tosave in CRM")
                    let contactId = await saveContactToCRM({"name":toSave[key][0],"email":toSave[key][1]})
                    await saveContactGroup(contactId,groupId)
                  }else{
                    console.log("nottosave in CRM")
                    await saveContactGroup(toSave[key][3],groupId)
                  }
                }
                // await setConfig(config, function(result) {
                //   confirmDialog.close();
                //   confirmDialog = null;
                // });

              });

            });

        }

        async function saveContactGroup(contactId,groupId){
          var url = config.url + '?';
          var data = {
              "entity": "GroupContact",
              "action": "create",
              "api_key": config.apikey,
              "key": config.sitekey,
              "json": {
                  "group_id":groupId,
                  "contact_id":contactId,
              }
            };
          for(var prop in data) {
              if (prop == 'json') {
                url = url + '&' + prop + '=' + JSON.stringify(data[prop]);
              } else {
                url = url + '&' + prop + '=' + data[prop];
              }
          }
          await $.post(url, function(result) {
            console.log(result)
          }) 

        }

        async function addGroupToCRM(name){
          var url = config.url + '?';
          var data = {
              "entity": "Group",
              "action": "create",
              "api_key": config.apikey,
              "key": config.sitekey,
              "json": {
                  "title":name,
                  "name":name,
              }
            };
          for(var prop in data) {
              if (prop == 'json') {
                url = url + '&' + prop + '=' + JSON.stringify(data[prop]);
              } else {
                url = url + '&' + prop + '=' + data[prop];
              }
          }
          let groupInfo
          await $.post(url, function(result) {
            console.log(result)
              groupInfo = result
          }) 

          return groupInfo
        }

        async function checkContact(val){
            var url = config.url + '?';
            var data = {
                "entity": "Email",
                "action": "get",
                "api_key": config.apikey,
                "key": config.sitekey,
                "json": {
                    "email":val.emailAddress,
                }
              };
            for(var prop in data) {
                if (prop == 'json') {
                  url = url + '&' + prop + '=' + JSON.stringify(data[prop]);
                } else {
                  url = url + '&' + prop + '=' + data[prop];
                }
            }
            var contact_info
            var exist = null
            var contact_url = ''
            var contact_name = null
            var contact_id
            for(const [key,val] of Object.entries(config.url.split("/"))){
              if(val == "sites"){
                break
              }
              contact_url += val
              contact_url += "/"
            }
            contact_url += "civicrm/contact/view/?reset=1&cid="
            await $.post(url, function(result) {
                  exist = false
                  if( result["count"] > 0 ){
                    exist = true
                    let keys = Object.keys(result["values"])
                    contact_id = result["values"][keys[0]]["contact_id"]
                    contact_url += String(contact_id)
                  }
            })
            if(exist){
              url = config.url + '?'
              data = {
                  "entity": "Contact",
                  "action": "getsingle",
                  "api_key": config.apikey,
                  "key": config.sitekey,
                  "json": {
                      "id":contact_id,
                  }
                };
              for(var prop in data) {
                  if (prop == 'json') {
                    url = url + '&' + prop + '=' + JSON.stringify(data[prop]);
                  } else {
                    url = url + '&' + prop + '=' + data[prop];
                  }
              }
              await $.post(url, function(result) {
                contact_name = result["display_name"]
              })
            }

            return {"exist":exist,"contact_url":contact_url,"contact_name":contact_name,"contact_id":contact_id}
        } 

        async function showContacts() {

            let html = '<ul class="ms-List">'

            //adding sender
            let res = await checkContact(requiredAttendees.from)
            html += await getListItem(requiredAttendees.from,res)

            //adding to
            for(const [key,val] of Object.entries(requiredAttendees.to)){
                let res = await checkContact(val)
                html += await getListItem(val,res)
            }

            //adding CC
            for(const [key,val] of Object.entries(requiredAttendees.cc)){
                let res = await checkContact(val)
                html += await getListItem(val,res)
            }

            //adding BCC
            for(const [key,val] of Object.entries(requiredAttendees.bcc)){
                let res = await checkContact(val)
                html += await getListItem(val,res)
            }
            html += '</ul>'

            html += '<button class="ms-Button ms-Button--small save-contact-all">'+
                      '<span class="ms-Button-label">Save all</span>'+
                    '</button>'
            html += '<button class="ms-Button ms-Button--small save-contact-all-group">'+
                      '<span class="ms-Button-label">Save all in Group</span>'+
                    '</button>'

            html += '<br><br>'

            html += '<button class="ms-Button ms-Button--medium save-email">'+
                      '<span class="ms-Button-label">Save Email</span>'+
                    '</button>'

            html += '<br><br>'

            html += '<button class="ms-Button ms-Button--medium save-folder-emails">'+
                      '<span class="ms-Button-label">Save Email</span>'+
                    '</button>'


            $("#contacts").html(html)
            $(".save-contact").on('click',saveContact)
            $(".save-contact-all-group").on('click',confirmSaveAllContactInGroup)
            $(".save-contact-all").on('click',confirmSaveAllContact)
            $(".save-email").on('click',saveEmail)
            
        }

        async function saveEmail() {
          var emailBody = null
          var datetime = new Date(requiredAttendees.dateTimeCreated)
          var hours = String(datetime.getHours())
          var seconds = String(datetime.getSeconds())
          var minutes = String(datetime.getMinutes())
          datetime = datetime.getFullYear()+"-"+(datetime.getMonth()+1) + "-" + datetime.getDate()

          datetime = datetime + " " + hours.padStart(2,"0") + ":" + minutes.padStart(2,"0") + ":" + seconds.padStart(2,"0")
          await requiredAttendees.body.getAsync(
            "html",
            async function callback(result) {
                emailBody = result['value']

                var data = {
                  "source_contact_id":202,
                  "activity_type_id": "Email",
                  "subject":requiredAttendees.subject,
                  "details":emailBody,
                  "activity_date_time":datetime.toString(),
                }

                var url = config.url + '?'
                data = {
                    "entity": "Activity",
                    "action": "create",
                    "api_key": config.apikey,
                    "key": config.sitekey,
                    "json": {...data}
                  };
                for(var prop in data) {
                    if (prop == 'json') {
                      url = url + '&' + prop + '=' + JSON.stringify(data[prop]);
                    } else {
                      url = url + '&' + prop + '=' + data[prop];
                    }
                }
                await $.post(url, function(result) {
                  console.log(result)
                })
          });
        }

        async function getConfig() {
          var config = {};

          config.url = Office.context.roamingSettings.get('civicrm_url');
          config.sitekey = Office.context.roamingSettings.get('civicrm_sitekey');
          config.apikey = Office.context.roamingSettings.get('civicrm_apikey');
          config.contacttype = Office.context.roamingSettings.get('civicrm_contacttype');
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
          await Office.context.roamingSettings.set('civicrm_contacttype', config.contacttype);
          Office.context.roamingSettings.saveAsync(callback);
        }
  }
})();
