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
    "ContactScreen": {
      "Placeholder": "Search",
      "TitleURL": "View Contact in CiviCRM",
    },
    "GroupScreen": {
      "Placeholder": "Search Group",
      "TitleURL": "View Group Settings in CiviCRM",
      "SearchContact": "Search contact in group",
      "SelectAll": "Select All",
      "UnselectAll": "Unselect All",
    },
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
    },
    "SaveContactInGroupScreen": {
      "Title": "Save Contacts to Group",
      "Done": "Save Contacts to Group",
      "Save": "Save Contacts",
      "SaveContact": "Save Contact to CiviCRM",
      "SavingText": "Contacts are being saved to CiviCRM, please wait...",
      "SavedText": "All selected contacts have been saved to CiviCRM",
    }
  };

  // JSON object for Dutch strings
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

  // JSON object for Spanish strings
  UIStrings.es_ES =
  {
    "Search": "Buscar",
    "Loading": "Recuperación de contactos...",
    "To": "Para",
    "Cc": "Cc",
    "Bcc": "Cco",
    "Settings": "Ajustes",
    "NotConfigured": "Este complemento requiere una configuración adicional. Por favor, elija el botón <strong>Configuración</strong> en la parte inferior de esta ventana.",
    "ContactScreen": {
      "Placeholder": "Buscar",
      "TitleURL": "Ver el contacto en CiviCRM",
    },
    "GroupScreen": {
      "Placeholder": "Buscar grupos",
      "TitleURL": "Ver la configuración del grupo en CiviCRM",
      "SearchContact": "Buscar contactos en el grupo",
      "SelectAll": "Selecionar Todo",
      "UnselectAll": "Deseleccionar todo",
    },
    "SettingsScreen": {
      "Title": "Ajustes",
      "NotConfigured": "Introduzca la URL del punto final REST de su CiviCRM y su clave de sitio y clave de API.<br>Si no sabe qué es, pregunte a su administrador de CiviCRM.",
      "URL": "CiviCRM REST URL",
      "URL_Placeholder": "Por ejemplo: http://your-site/sites/all/modules/civicrm/extern/rest.php",
      "SiteKey": "Clave del sitio de CiviCRM",
      "ApiKey": "Clave de la API CiviCRM",
      "ContactType": "Tipo de contacto",
      "Done": "Guardar"
    },
    "SaveContactScreen": {
      "Title": "Guarda el contacto",
      "ContactName": "Nombre del contacto",
      "ContactEmail": "Correo electrónico de contacto",
      "ContactType": "Tipo de contacto",
      "Done": "Guardar"
    },
    "SaveContactInGroupScreen": {
      "Title": "Guarda los contactos en el grupo",
      "Done": "Guarda los contactos en el grupo",
      "Save": "Guardar",
      "SaveContact": "Guarda el contacto en CiviCRM",
      "SavingText": "Los contactos se están guardando en CiviCRM, por favor espere...",
      "SavedText": "Todos los contactos seleccionados se han guardado en CiviCRM",
    }
  };

  // JSON object for Catalan strings
  UIStrings.ca_ES =
  {
    "Search": "Cerca",
    "Loading": "Recuperació de contactes...",
    "To": "Per a",
    "Cc": "A/c",
    "Bcc": "C/o",
    "Settings": "Configuració",
    "NotConfigured": "Aquest complement requereix una configuració addicional. Si us plau, trieu el botó <strong>Configuració</strong> a la part inferior d'aquesta finestra.",
    "ContactScreen": {
      "Placeholder": "Cerca",
      "TitleURL": "Veure el contacte a CiviCRM",
    },
    "GroupScreen": {
      "Placeholder": "Cerca de grups",
      "TitleURL": "Veure la configuració del grup a CiviCRM",
      "SearchContact": "Cerca de contactes al grup",
      "SelectAll": "Seleccionar tot",
      "UnselectAll": "Desseleccionar tot",
    },
    "SettingsScreen": {
      "Title": "Configuració",
      "NotConfigured": "Introduïu l'URL del punt final REST del vostre CiviCRM i la vostra clau de lloc i clau d'API.<br>Si no sabeu què és, pregunteu al vostre administrador de CiviCRM.",
      "URL": "CiviCRM REST URL",
      "URL_Placeholder": "Per exemple: http://your-site/sites/all/modules/civicrm/extern/rest.php",
      "SiteKey": "Clau del lloc de CiviCRM",
      "ApiKey": "Clau de l'API CiviCRM",
      "ContactType": "Tipus de contacte",
      "Done": "Desa"
    },
    "SaveContactScreen": {
      "Title": "Desa el contacte",
      "ContactName": "Nom de contacte",
      "ContactEmail": "Correu electrònic de contacte",
      "ContactType": "Tipus de contacte",
      "Done": "Desa"
    },
    "SaveContactInGroupScreen": {
      "Title": "Desa els contactes al grup",
      "Done": "Desa els contactes al grup",
      "Save": "Desa",
      "SaveContact": "Desa el contacte a CiviCRM",
      "SavingText": "Els contactes s'estan desant a CiviCRM, espereu...",
      "SavedText": "Tots els contactes seleccionats s'han desat a CiviCRM",
    }
  };
  // JSON for Canada Francais
  UIStrings.fr_CA =
  {
    "Search": "Recherche",
    "Loading": "Récupérer des contacts...",
    "To": "To",
    "Cc": "Cc",
    "Bcc": "Bcc",
    "Settings": "Paramètres",
    "NotConfigured": "Cet ajout nécessite une configuration supplémentaire. Veuillez choisir l'option <strong>Paramètres</strong> button at the bottom of this win$
    "ContactScreen": {
      "Placeholder": "Recherche",
      "TitleURL": "Voir le contact dans CiviCRM",
    },
    "GroupScreen": {
      "Placeholder": "Groupe de recherche",
      "TitleURL": "Afficher les paramètres du groupe dans CiviCRM",
      "SearchContact": "Rechercher un contact dans le groupe",
      "SelectAll": "Tout sélectionner",
      "UnselectAll": "Tout déselectionner",
    },
    "SettingsScreen": {
      "Title": "Paramètres",
      "NotConfigured": "Veuillez entrer l'URL REST de CiviCRM et vos clés 'Site key' et 'API key'.<br>Au besoin, demandez à votre administrateur CiviCRM.",
      "URL": "CiviCRM URL REST",
      "URL_Placeholder": "ex: http://votre-site/sites/all/modules/civicrm/extern/rest.php",
      "SiteKey": "'Site Key' de CiviCRM",
      "ApiKey": "'API Key' de CiviCRM",
      "ContactType": "Type de contact",
      "Done": "Enregistrer"
    },
    "SaveContactScreen": {
      "Title": "Sauvegarder le contact",
      "ContactName": "Nom du contact",
      "ContactEmail": "Courriel du contact",
      "ContactType": "Type de contact",
      "Done": "Enregistrer"
    },
    "SaveContactInGroupScreen": {
      "Title": "Enregistrer les contacts dans un groupe",
      "Done": "Enregistrer les contacts dans un groupe",
      "Save": "Sauvegarder les contacts",
      "SaveContact": "Enregistrer un contact dans CiviCRM",
      "SavingText": "Les contacts sont en train d'être enregistrés dans CiviCRM, veuillez patienter...",
      "SavedText": "Tous les contacts sélectionnés ont été enregistrés dans CiviCRM",
    }
  };

  UIStrings.getLocaleStrings = function (locale) {
    var text;
    var localeProp = locale.replace("-", "_");
    if (UIStrings[localeProp]) {
      text = $.extend( true, UIStrings.default, UIStrings[localeProp] );
    } else {
      text = UIStrings.default;
    }
    return text;
  };

  return UIStrings;
})();
