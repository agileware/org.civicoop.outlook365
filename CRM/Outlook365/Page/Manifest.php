<?php
use CRM_Outlook365_ExtensionUtil as E;

class CRM_Outlook365_Page_Manifest extends CRM_Core_Page {

  public function run() {
    // Example: Set the page-title dynamically; alternatively, declare a static title in xml/Menu/*.xml
    CRM_Utils_System::setTitle(E::ts('Manifest'));


    $defaultContact = civicrm_api3('Contact', 'getvalue', ['id' => 1, 'return' => 'display_name']);
    $defaultContactName = CRM_Utils_String::convertStringToCamel($defaultContact);

    $this->_print = CRM_Core_Smarty::PRINT_SNIPPET;
    $baseUrl = E::url('');
    $this->assign('baseurl', $baseUrl);
    $this->assign('default_contact', $defaultContact);
    $this->assign('default_contact_name', $defaultContactName);

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
