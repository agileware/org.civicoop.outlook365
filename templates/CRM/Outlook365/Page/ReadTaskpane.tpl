{crmScope extensionKey='outlook365'}
<!DOCTYPE html>
<html>

<head>
  <meta charset="UTF-8" />
  <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>CiviCRM</title>
  <script type="text/javascript">
    var confirmDialogUrl = '{$baseurl}outlook365/settings/confirm.html';
    var settingsDialogUrl = '{$baseurl}outlook365/settings/dialog.html';
    var saveContactDialogUrl = '{$baseurl}outlook365/settings/saveContact.html';
    var saveContactInGroupDialogUrl = '{$baseurl}outlook365/settings/saveContactInGroup.html';
  </script>

  <!-- Office JavaScript API -->
  <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>
  <script type="text/javascript" src="{$baseurl}assets/jquery-3.4.1.min.js"></script>

  <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.min.css" />
  <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.components.min.css" />
  <script type="text/javascript" src="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/js/fabric.min.js"></script>

  <link rel="stylesheet" href="{$baseurl}assets/taskpane.css"/>
</head>

<body class="ms-font-m ms-Fabric ms-landing-page">
<main class="ms-landing-page__main">
  <section class="ms-landing-page__content ms-font-m ms-fontColor-neutralPrimary">
    <div class="not-configured-warning ms-MessageBar ms-MessageBar--warning" style="display: none;">
      <div class="ms-font-l" id="settings-prompt"></div>
    </div>

    <div class="ms-CommandBar-mainArea">
      <div class="ms-CommandButton ms-CommandButton--pivot is-active">
          <!-- <a class="ms-CommandButton-button"> <span class="ms-CommandButton-label contacts">Contacts</span>  </a>  -->
          <i class="ms-Icon ms-Icon--Contact" aria-hidden="true" ><span class="ms-font-l contacts">{ts}Contacts{/ts}</span></i>
      </div>
      <div class="ms-CommandButton ms-CommandButton--pivot" id="settings-icon">
          <i class="ms-Icon ms-Icon--Settings" aria-hidden="true"> <span class="ms-font-l settings">{ts}Settings{/ts}</span></i>
      </div>
      </div>
    <div class="dataclass">
      <br>
    </div>
    <div>
      <button class="ms-Button ms-Button--small selectAll">
        <span class="ms-Button-label">{ts}Select All{/ts}</span>
      </button>
      <button class="ms-Button ms-Button--small">
        <span class="ms-Button-label unselectAll">{ts}Unselect All{/ts}</span>
      </button>
    </div>

    <div class="dataclass" id="contacts">
    </div>
    <div class="ms-Spinner" id="loadingContacts">
      <div class="ms-Spinner-label">
        <br>
        {ts}Loading...{/ts}
      </div>
    </div>
    <hr>

    <div class="ms-CommandBar-mainArea">
      <h2>{ts}Save Email{/ts}</h2>
      <p>{ts}Click the following button to save this email in CiviCRM as an Activity{/ts}</p>
      <button class="ms-Button ms-Button--medium save-email">
        <span class="ms-Button-label">{ts}Save Email{/ts}</span>
      </button>
      <p>{ts}Select one or more email folders to save all the emails to CiviCRM as an Activity.{/ts}</p>
          <form id="folder-form">

            <ul class="ms-List" id="target">
            </ul>
            <button class="ms-Button ms-Button--medium" id="send-submit">
              <span class="ms-Button-label">{ts}Save Folder in CiviCRM{/ts}</span>
            </button>
            <p id="saving-email-notice" style="display: none;">{ts}Emails are being saved to CiviCRM, please wait...{/ts}</p>
            <p id="saved-email-notice" style="display: none;">{ts}All emails have been saved to CiviCRM and assigned the "Saved in CiviCRM" category{/ts}</p>
          </form>
      </div>

  </section>
</main>
<!-- <footer class="ms-landing-page__footer ms-bgColor-neutralLighter ms-bgColor-neutralLight--hover">
  <div id="settings-icon" class="ms-landing-page__footer--left ms-bgColor-neutralLight--hover ms-fontColor-neutralDark ms-fontColor-neutralDarker--hover" aria-label="Settings" tabindex=0>
    <i class="ms-Icon enlarge ms-Icon--Settings "></i><span class="label"></span>
  </div>
</footer> -->

<div class="civicrm-notice">
  <div class="ms-Dialog">
    <div class="ms-Dialog-title">CiviCRM</div>
    <div class="ms-Dialog-content" id="civicrm-notice-div">
      <p class="ms-Dialog-subText" id="civicrm-notice-text">{ts}An error occurred communicating with CiviCRM. This action could not be completed{/ts}</p>
    </div>
    <div class="ms-Dialog-actions">
      <button class="ms-Button ms-Dialog-action ms-Button--primary">
        <span class="ms-Button-label">OK</span>
      </button>
    </div>
  </div>
</div>
<script type="text/javascript">
  {literal}
  var noticeDialog = document.querySelector(".ms-Dialog");
  var dialogComponent = new fabric['Dialog'](noticeDialog);
  function openDialog(dialog) {
    // Open the dialog
    dialog.open();
  }
  {/literal}
</script>

<script>
  var CRMContactURL = "{$contactURL}";
  {literal}
  var SpinnerElements = document.querySelectorAll(".ms-Spinner");
  for (var i = 0; i < SpinnerElements.length; i++) {
    new fabric['Spinner'](SpinnerElements[i]);
  }
  var SearchBoxElements = document.querySelectorAll(".ms-SearchBox");
  for (var i = 0; i < SearchBoxElements.length; i++) {
    new fabric['SearchBox'](SearchBoxElements[i]);
  }
  {/literal}
</script>
<script type="text/javascript" src="{$baseurl}assets/UIStrings.js"></script>
<script type="text/javascript" src="{$baseurl}assets/readtaskpane.js"></script>
</body>

</html>
{/crmScope}
