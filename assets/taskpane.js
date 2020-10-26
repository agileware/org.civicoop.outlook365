(function() {
  'use strict';
  Office.initialize = function (reason) {
    var item = Office.context.mailbox.item;
    var currentOffset = 0;
    var currentTab = "contacts"
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
      if (config) {
        loadNextContacts();
      } else {
        openSettingsDialog();
      }

      $('#settings-icon').on('click', openSettingsDialog);
      $('#reset').on('click', clearRecipients);
      $('.ms-CommandButton--pivot span').on('click', handleTabChange);
    });

    $("body").delegate(".ms-ListItem", "click", function(event){
        if($(event.target).hasClass('is-selected')){
          $(event.target).removeClass('is-selected');
        }
        else{
          $(event.target).addClass('is-selected');
        }
    });

    $("body").delegate(".selectAll", "click", function(event){
        let parentElement = $(event.target).closest(".allData").parent()
        parentElement.children("ul").children(".ms-ListItem").addClass("is-selected")
    });

    $("body").delegate(".unselectAll", "click", function(event){
        let parentElement = $(event.target).closest(".allData").parent()
        parentElement.children("ul").children(".ms-ListItem").removeClass("is-selected")
    });

    $('#searchField').on("keypress", function(e) {
      if (e.keyCode == 13) {
        reset();
        search = $(this).val();
        if(currentTab==="contacts"){
          loadNextContacts();
        }
        else if (currentTab==="groups"){
          loadNextGroups();
        }
        return false; // prevent the button click from happening
      }
    });

    $(window).scroll(function() {
      if($(window).scrollTop() == $(document).height() - $(window).height()) {
        // ajax call get data from server and append to the div
        if(currentTab==="contacts"){
          loadNextContacts();
        }
        else if (currentTab==="groups"){
          loadNextGroups();
        }
      }
    });

    /**
      * Handle Tab Change
    */
    function handleTabChange(event){
      let classes = event.target.className.split(" ")
      let targetTab = classes[classes.length - 1]
      if(targetTab === "settings"){
        return
      }
      currentTab = targetTab
      let parentselector = $($(event.target).parent()).parent()
      $(".ms-CommandButton.ms-CommandButton--pivot").removeClass("is-active")
      parentselector.addClass("is-active")
      $(".dataclass").empty()
      reset()
      if(currentTab === "contacts"){
        $("#search-form").show()
        loadNextContacts()
      }
      else if(currentTab === "groups"){
        $("#search-form").show()
        loadNextGroups()
      }

    }

    function getSearchForm(id){
      let html = '<div id="group_search_' + id + '">'+
        '<div class="ms-SearchBox ms-SearchBox--commandBar">'+
          '<input class="ms-SearchBox-field" id="groupSearchField" type="text" value="">'+
          '<label class="ms-SearchBox-label">'+
            '<i class="ms-SearchBox-icon ms-Icon ms-Icon--Search"></i>'+
            '<span class="ms-SearchBox-text">Search</span>'+
          '</label>'+
          '<div class="ms-CommandButton ms-SearchBox-clear ms-CommandButton--noLabel">'+
            '<button class="ms-CommandButton-button">'+
              '<span class="ms-CommandButton-icon"><i class="ms-Icon ms-Icon--Clear"></i></span>'+
              '<span class="ms-CommandButton-label"></span>'+
            '</button>'+
          '</div>'+
          '<div class="ms-CommandButton ms-SearchBox-exit ms-CommandButton--noLabel">'+
            '<button class="ms-CommandButton-button">'+
              '<span class="ms-CommandButton-icon"><i class="ms-Icon ms-Icon--ChromeBack"></i></span>'+
              '<span class="ms-CommandButton-label"></span>'+
            '</button>'+
          '</div>'+
          '<div class="ms-CommandButton ms-SearchBox-filter ms-CommandButton--noLabel">'+
            '<button class="ms-CommandButton-button">'+
              '<span class="ms-CommandButton-icon"><i class="ms-Icon ms-Icon--Filter"></i></span>'+
              '<span class="ms-CommandButton-label"></span>'+
            '</button>'+
          '</div>'+
        '</div>'+
      '</div>'

      return html
    }

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
        if(currentTab!=="settings"){
          $('#search-form').show();
        }
      }
      $('#loadingContacts').hide();
      $('#contacts').html('');
      $('#groups').html('');
      currentOffset = 0;
      moreAvailable = true;
      search = false;
      fields = [];
      if(currentTab === "contacts"){
        getContactFields();
      }
    }

    /**
     * Retrieve next batch of contacts.
     */
    function getContactFields() {
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
        }
      });
    }

    /**
     * Retrieve next batch of Groups.
     */
    function loadNextGroups() {

      if (!moreAvailable || !config) {
        return;
      }
      var url = config.url + '?';
      var data = {
        "entity": "Group",
        "action": "get",
        "api_key": config.apikey,
        "key": config.sitekey,
        "json": {
          "sequential": 1,
          "return": ["id","name", "title"],
          "options": {
            "offset": currentOffset,
            "limit": 25,
          }
        }
      };
      if (search) {
        data.json.name = {"LIKE": '%'+search+'%'};
      }
      for(var prop in data) {
        if (prop == 'json') {
          url = url + '&' + prop + '=' + encodeURI(JSON.stringify(data[prop]));
        } else {
          url = url + '&' + prop + '=' + data[prop];
        }
      }
      $.getJSON(url, {}, addGroups);
    }

    /**
     * Add the group to the list.
     *
     * @param data
     */
    function addGroups(data) {

      if (data.is_error == 0) {
        for(var i in data.values) {
          var group = data.values[i];
          var idname = group.name.replace(" ","-").toLowerCase();
          var name = group.title;
          var id = group.id;

          var html = '' +
            '<div class="ms-Persona">'+
            '<div class="ms-Persona-details">' +
            '<div class="ms-Persona-primaryText CiviCRM-Group-Email"  data-civicrm-id="'+id+'" data-civicrm-name="'+name+'">' +
            name +
            '<i class="ms-Icon ms-Icon--ChevronRight" id="'+
            idname +
            '-expand-groups" style="padding: 6px"></i>'+
            '</div>' +
            '</div>' +
            '</div>';
          $('#groups').append(html);

        $('#'+idname+'-expand-groups').on('click', expandGroups);
        }

        if (data.count < 25) {
          moreAvailable = false;
        }

      }
    }

    async function expandGroups(event){
      if($(this).parent().hasClass("expanded")){
          let name = $(this).parent('.CiviCRM-Group-Email').data('civicrm-name').replace(" ","-").toLowerCase();
          $(this).parent().removeClass('expanded')
          $(this).parent().children(".allData").remove()
          $(this).parent().children(".CiviCRM-Group-Email").remove()
          $(this).parent().children("#group_search_"+name).remove()
          $(this).parent().children("ul").remove()
          $(this).attr('class', 'ms-Icon ms-Icon--ChevronRight');
          return;
      }
      $(this).parent().addClass('expanded')
      let name = $(this).parent('.CiviCRM-Group-Email').data('civicrm-name').replace(" ","-").toLowerCase();
      let id = $(this).parent('.CiviCRM-Group-Email').data('civicrm-id');
      // $(this).parent().append('<div id="' + name + '-expanded-groups"></div>');
      var url = config.url + '?';
      var data = {
          "entity": "Outlook365Group",
          "action": "get",
          "api_key": config.apikey,
          "key": config.sitekey,
          "json": {
            "sequential": 1,
            "options": {
              "limit": 0,
            },
            "group_id":id,
          }
        };
        for(var prop in data) {
          if (prop == 'json') {
            url = url + '&' + prop + '=' + JSON.stringify(data[prop]);
          } else {
            url = url + '&' + prop + '=' + data[prop];
          }
        }
        let html = await getGroupContacts(url)

        var buttons = '<button class="ms-Button ms-Button--small to"><span class="ms-Button-label">'+UIText.To+'</span></button>' +
          '<button class="ms-Button ms-Button--small cc"><span class="ms-Button-label">'+UIText.Cc+'</span></button>' +
          '<button class="ms-Button ms-Button--small bcc"><span class="ms-Button-label">'+UIText.Bcc+'</span></button>';
        buttons = '<div class="CiviCRM-Group-Email" data-civicrm-id="'+id+'" data-civicrm-name="'+name+'">'+buttons+'</div>';
        $(event.target).parent().append(buttons);
        $(event.target).parent().append('<div class="allData"><button class="ms-Button ms-Button--small selectAll"><span class="ms-Button-label">Select All</span></button>' +
                                        '<button class="ms-Button ms-Button--small"><span class="ms-Button-label unselectAll">Unselect All</span></button></div>')
        $(event.target).parent().append(getSearchForm(name))
        $("#group_search_"+name+" .ms-SearchBox-text").text("");

        $(event.target).parent().append('<ul class="ms-List ' +name +'-list-email">'+html+'</ul>')



        // to get smoothness change class after query
        $(this).attr('class', 'ms-Icon ms-Icon--ChevronDown');

        $("#groups .ms-Button.to").click(function(event) {
          let toSend = []
          var recipients = item.to
          $('.'+name+'-list-email').children().each(function (index) {
              if($(this).hasClass('is-selected')){
                let contact = $(this)
                toSend.push([contact.data('civicrm-name'),contact.data('civicrm-email')])
              }

          });
          for(var iter in toSend){
            addReiever(recipients, toSend[iter][1], toSend[iter][0]);
          }
        });

        $("#groups .ms-Button.cc").click(function(event) {
          let toSend = []
          var recipients = item.cc
          $('.'+name+'-list-email').children().each(function (index) {
              if($(this).hasClass('is-selected')){
                let contact = $(this)
                toSend.push([contact.data('civicrm-name'),contact.data('civicrm-email')])
              }

          });
          for(var iter in toSend){
            addReiever(recipients, toSend[iter][1], toSend[iter][0]);
          }
        });

        $("#groups .ms-Button.bcc").click(function(event) {
          let toSend = []
          var recipients = item.bcc
          $('.'+name+'-list-email').children().each(function (index) {
              if($(this).hasClass('is-selected')){
                let contact = $(this)
                toSend.push([contact.data('civicrm-name'),contact.data('civicrm-email')])
              }

          });
          for(var iter in toSend){
            addReiever(recipients, toSend[iter][1], toSend[iter][0]);
          }
        });

        $("#group_search_"+name+" #groupSearchField").on("keypress", async function(e) {
            if (e.keyCode == 13) {
              let newUrl = config.url + '?';
              let groupSearchVal = $(this).val();
              console.log(groupSearchVal)
              if (groupSearchVal!='') {
                data.json.display_name = {"LIKE": '%'+groupSearchVal+'%'};
              }
              for(var prop in data) {
                if (prop == 'json') {
                  newUrl = newUrl + '&' + prop + '=' + encodeURI(JSON.stringify(data[prop]));
                } else {
                  newUrl = newUrl + '&' + prop + '=' + data[prop];
                }
              }
              let newHtml = await getGroupContacts(newUrl)
              $("."+name +'-list-email').html(newHtml)
              console.log(newHtml)
              return false; // prevent the button click from happening
            }
          });
    }

    async function getGroupContacts(url){
      var html =''
      await $.getJSON(url, {}, function(data){
          for(var i in data.values){
            var contact = data.values[i]
            html += '<li class="ms-ListItem is-selectable" data-civicrm-name="'+contact.display_name+
                        '" data-civicrm-email="'+contact.email+'">'+
                        '<span class=""ms-ListItem-primaryText">'+ contact.display_name+ '</span>' +
                        '<div class="ms-ListItem-selectionTarget"></div></li>'
            // break
          }
        });
      html += '<script type="text/javascript">\n' +
        '  var ListItemElements = document.querySelectorAll(".ms-ListItem");\n' +
        '  for (var i = 0; i < ListItemElements.length; i++) {\n' +
        '    new fabric[\'ListItem\'](ListItemElements[i]);\n' +
        '  }\n' +
        '</script>';
      return html
    }

    /**
     * Retrieve next batch of contacts.
     */
    function loadNextContacts() {
      if (!moreAvailable || !config || !search) {
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
      url = encodeURI(url);
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
          // for(var fieldI in fields) {
          //   var fieldName = fields[fieldI].name;
          //   var value = contact[fieldName];
          //   if (value) {
          //     secondaryFields = secondaryFields + '<div class="ms-Persona-secondaryText"><strong>' + fields[fieldI].title + ':</strong>&nbsp;' + value + '</div>';
          //   }
          // }
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
     * Clear Recipients
     */
    function clearRecipients() {
      var toRecipients, ccRecipients, bccRecipients;
      if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
          toRecipients = item.requiredAttendees;
          ccRecipients = item.optionalAttendees;
      }
      else {
          toRecipients = item.to;
          ccRecipients = item.cc;
          bccRecipients = item.bcc;
      }

      toRecipients.setAsync([],
          function (asyncResult) {
              if (asyncResult.status == Office.AsyncResultStatus.Failed){
                  write(asyncResult.error.message);
              }
              else {
              }
      });

      ccRecipients.setAsync([],
          function (asyncResult) {
              if (asyncResult.status == Office.AsyncResultStatus.Failed){
                  write(asyncResult.error.message);
              }
              else {
              }
      });

      bccRecipients.setAsync([],
          function (asyncResult) {
              if (asyncResult.status == Office.AsyncResultStatus.Failed){
                  write(asyncResult.error.message);
              }
              else {
              }
      });

    }

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
            if(currentTab==="contacts"){
              loadNextContacts();
            }
            else if(currentTab==="groups"){
              loadNextGroups();
            }
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
