<?php
use CRM_Outlook365_ExtensionUtil as E;

class CRM_Outlook365_Page_Manifest extends CRM_Core_Page {

  public function run() {
    // Example: Set the page-title dynamically; alternatively, declare a static title in xml/Menu/*.xml
    CRM_Utils_System::setTitle(E::ts('Manifest'));

    $site_url = CRM_Utils_System::baseURL();
    $domainContactID = civicrm_api3('Domain', 'getvalue', ['return' => "contact_id", 'current_domain' => $site_url]);
    $domainContact = civicrm_api3('Contact', 'getvalue', ['id' => $domainContactID, 'return' => 'display_name']);
    $domainContactName = CRM_Utils_String::convertStringToCamel($domainContact);
    $guid = CRM_Outlook365_Utils_Uuid::v5('3b44a0ed-311f-4f35-ba77-d8467a3624f6', $site_url);

    $this->_print = CRM_Core_Smarty::PRINT_SNIPPET;
    $baseUrl = E::url('');
    $this->assign('baseurl', $baseUrl);
    $this->assign('default_contact', $domainContact);
    $this->assign('default_contact_name', $domainContactName);
    $this->assign('guid', $guid);

    self::$_template->assign('mode', $this->_mode);
    $pageTemplateFile = $this->getHookedTemplateFileName();
    self::$_template->assign('tplFile', $pageTemplateFile);
    // invoke the pagRun hook, CRM-3906
    CRM_Utils_Hook::pageRun($this);

    $content = self::$_template->fetch($pageTemplateFile);
    CRM_Utils_System::appendTPLFile($pageTemplateFile, $content);

    //its time to call the hook.
    CRM_Utils_Hook::alterContent($content, 'page', $pageTemplateFile, $this);


    CRM_Utils_System::setHttpHeader('Content-Type', 'application/xml');
    CRM_Utils_System::setHttpHeader('Content-Disposition', 'attachment; filename="manifest.xml"');

    echo $content;

    CRM_Utils_System::civiExit();
  }

}
