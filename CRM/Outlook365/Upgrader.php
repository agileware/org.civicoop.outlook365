<?php
use CRM_Outlook365_ExtensionUtil as E;

/**
 * Collection of upgrade steps.
 */
class CRM_Outlook365_Upgrader extends CRM_Outlook365_Upgrader_Base {

  /**
   * Add the data processor from this extension
   *
   * @throws \CiviCRM_API3_Exception
   */
  public function postInstall() {
    CRM_Dataprocessor_Utils_Extensions::updateDataProcessorsFromExtension('outlook365');
  }

  /**
   * Look up extension dependency error messages and display as Core Session Status
   *
   * @param array $unmet
   */
  public static function displayDependencyErrors(array $unmet){
    foreach ($unmet as $ext) {
      $message = self::getUnmetDependencyErrorMessage($ext);
      CRM_Core_Session::setStatus($message, E::ts('Prerequisite check failed.'), 'error');
    }
  }

  /**
   * Mapping of extensions names to localized dependency error messages
   *
   * @param string $unmet an extension name
   */
  public static function getUnmetDependencyErrorMessage($unmet) {
    switch ($unmet[0]) {
      case 'dataprocessor':
        return ts('Outlook 365 was installed successfully, but you must also install and enable the <a href="%1">dataprocessor Extension</a> version %2 or newer.', array(1 => 'https://lab.civicrm.org/extensions/dataprocessor', 2=>$unmet[1]));
    }

    CRM_Core_Error::fatal(ts('Unknown error key: %1', array(1 => $unmet)));
  }

  /**
   * Extension Dependency Check
   *
   * @return Array of names of unmet extension dependencies; NOTE: returns an
   *         empty array when all dependencies are met.
   */
  public static function checkExtensionDependencies() {
    $manager = CRM_Extension_System::singleton()->getManager();

    $dependencies = array(
      ['dataprocessor', '1.1']
    );

    $unmet = array();
    foreach($dependencies as $ext) {
      if (!self::checkExtensionVersion($ext[0], $ext[1])) {
        array_push($unmet, $ext);
      }
    }
    return $unmet;
  }

  public static function checkExtensionVersion($extension, $version) {
    try {
      static $extensions = null;
      if (!$extensions) {
        $extensions = civicrm_api3('Extension', 'get', array('options' => array('limit' => 0)));
      }
      foreach($extensions['values'] as $ext) {
        if ($ext['key'] == $extension && $ext['status'] == 'installed') {
          if (version_compare($ext['version'], $version, '>=')) {
            return true;
          }
        }
      }
    }
    catch (Exception $e) {
      return FALSE;
    }
    return FALSE;
  }

}
