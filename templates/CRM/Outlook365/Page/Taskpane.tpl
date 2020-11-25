<!DOCTYPE html>
<html>

<head>
  <meta charset="UTF-8" />
  <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>CiviCRM</title>
  <script type="text/javascript">
    var settingsDialogUrl = '{$baseurl}outlook365/settings/dialog.html';
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
          <i class="ms-Icon ms-Icon--Contact" aria-hidden="true" ><span class="ms-font-l contacts">Contacts</span></i>
      </div>
      <div class="ms-CommandButton ms-CommandButton--pivot">
          <i class="ms-Icon ms-Icon--Group" aria-hidden="true"><span class="ms-font-l groups">Groups</span></i>
      </div>
      <div class="ms-CommandButton ms-CommandButton--pivot" id="settings-icon">
          <i class="ms-Icon ms-Icon--Settings" aria-hidden="true"> <span class="ms-font-l settings">Settings</span>  </i>
      </div>
      <div id="search-form">
        <div class="ms-SearchBox ms-SearchBox--commandBar">
          <input class="ms-SearchBox-field" id="searchField" type="text" value="">
          <label class="ms-SearchBox-label">
            <i class="ms-SearchBox-icon ms-Icon ms-Icon--Search"></i>
            <span class="ms-SearchBox-text">Search</span>
          </label>
          <div class="ms-CommandButton ms-SearchBox-clear ms-CommandButton--noLabel">
            <button class="ms-CommandButton-button">
              <span class="ms-CommandButton-icon"><i class="ms-Icon ms-Icon--Clear"></i></span>
              <span class="ms-CommandButton-label"></span>
            </button>
          </div>
          <div class="ms-CommandButton ms-SearchBox-exit ms-CommandButton--noLabel">
            <button class="ms-CommandButton-button">
              <span class="ms-CommandButton-icon"><i class="ms-Icon ms-Icon--ChromeBack"></i></span>
              <span class="ms-CommandButton-label"></span>
            </button>
          </div>
          <div class="ms-CommandButton ms-SearchBox-filter ms-CommandButton--noLabel">
            <button class="ms-CommandButton-button">
              <span class="ms-CommandButton-icon"><i class="ms-Icon ms-Icon--Filter"></i></span>
              <span class="ms-CommandButton-label"></span>
            </button>
          </div>
        </div>
      </div>
    </div>
    <div class="dataclass">
      <br>
    </div>
    <div class="dataclass" id="contacts">
    </div>
    <div class="dataclass" id="groups">
    </div>

  </section>
</main>

<script>
  var CRMContactURL = "{$contactURL}";
  var CRMGroupURL = "{$groupURL}";
  {literal}
  var SearchBoxElements = document.querySelectorAll(".ms-SearchBox");
  for (var i = 0; i < SearchBoxElements.length; i++) {
    new fabric['SearchBox'](SearchBoxElements[i]);
  }
  {/literal}
</script>

<script type="text/javascript" src="{$baseurl}assets/UIStrings.js"></script>
<script type="text/javascript" src="{$baseurl}assets/taskpane.js"></script>
</body>

</html>
