/* Store the locale-specific strings */

var UIStrings = (function () {
  "use strict";

  var UIStrings = {};

  // JSON object for English strings
  UIStrings.default =
    {
      "Search": "Search",
      "Loading": "Retrieving contacts...",
      "To": "To",
      "Cc": "Cc",
      "Bcc": "Bcc",
      "Settings": "Settings",
      "NotConfigured": "This add in requires additional configuration. Please choose the <strong>Settings</strong> button at the bottom of this window.",
      "SettingsScreen": {
        "Title": "Settings",
        "NotConfigured": "Please enter the URL to the REST endpoint of your CiviCRM and your Site key and API key.<br>If you don't know what this is ask your administrator of CiviCRM.",
        "URL": "CiviCRM REST URL",
        "URL_Placeholder": "E.g. http://your-site/sites/all/modules/civicrm/extern/rest.php",
        "SiteKey": "CiviCRM Site Key",
        "ApiKey": "CiviCRM API Key",
        "ContactType": "Contact Type",
        "Done": "Save"
      },
      "SaveContactScreen": {
        "Title": "Save Contact",
        "ContactName": "Contact Name",
        "ContactEmail": "Contact Email",
        "ContactType": "Contact Type",
        "Done": "Done"
      }
      ,
      "SaveContactInGroupScreen": {
        "Title": "Save Contacts to Group",
        "Done": "Save Contacts to Group"
      }
    };

  // JSON object for Spanish strings
  UIStrings.nl_NL =
    {
      "Search": "Zoeken",
      "Loading": "Contacten ophalen...",
      "To": "Aan",
      "Cc": "Cc",
      "Bcc": "Bcc",
      "Settings": "Instellingen",
      "NotConfigured": "Deze Add is nog niet geconfigureerd. Klik op de <strong>Instellingen</strong> knop onderaan dit scherm.",
      "SettingsScreen": {
        "Title": "Instelling",
        "NotConfigured": "Geef de URL van de REST interface van je CiviCRM en je sitekey en API key..<br>Mocht je niet weten wat dit is vraag dan de beheerder van je CiviCRM voor deze gegevens.",
        "URL": "URL naar je CiviCRM REST interface",
        "SiteKey": "Site key",
        "ApiKey": "Api Key",
        "Done": "Opslaan"
      }
    };

  UIStrings.getLocaleStrings = function (locale) {
    var text;
    var localeProp = locale.replace("-", "_");
    if (UIStrings[localeProp]) {
      text = UIStrings[localeProp];
    } else {
      text = UIStrings.default;
    }
    return text;
  };

  return UIStrings;
})();
