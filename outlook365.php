<?php

require_once 'outlook365.civix.php';
use CRM_Outlook365_ExtensionUtil as E;

/**
 * Implementation of hook_civicrm_pageRun()
 *
 * @link https://docs.civicrm.org/dev/en/latest/hooks/hook_civicrm_pageRun/
 */
function outlook365_civicrm_pageRun(&$page) {
  if ($page instanceof CRM_Admin_Page_Extensions) {
    $unmet = CRM_Outlook365_Upgrader::checkExtensionDependencies();
    CRM_Outlook365_Upgrader::displayDependencyErrors($unmet);
  }
}

/**
 * Implements hook_civicrm_config().
 *
 * @link https://docs.civicrm.org/dev/en/latest/hooks/hook_civicrm_config/
 */
function outlook365_civicrm_config(&$config) {
  _outlook365_civix_civicrm_config($config);
}

/**
 * Implements hook_civicrm_xmlMenu().
 *
 * @link https://docs.civicrm.org/dev/en/latest/hooks/hook_civicrm_xmlMenu
 */
function outlook365_civicrm_xmlMenu(&$files) {
  _outlook365_civix_civicrm_xmlMenu($files);
}

/**
 * Implements hook_civicrm_install().
 *
 * @link https://docs.civicrm.org/dev/en/latest/hooks/hook_civicrm_install
 */
function outlook365_civicrm_install() {
  _outlook365_civix_civicrm_install();
}

/**
 * Implements hook_civicrm_postInstall().
 *
 * @link https://docs.civicrm.org/dev/en/latest/hooks/hook_civicrm_postInstall
 */
function outlook365_civicrm_postInstall() {
  _outlook365_civix_civicrm_postInstall();
}

/**
 * Implements hook_civicrm_uninstall().
 *
 * @link https://docs.civicrm.org/dev/en/latest/hooks/hook_civicrm_uninstall
 */
function outlook365_civicrm_uninstall() {
  _outlook365_civix_civicrm_uninstall();
}

/**
 * Implements hook_civicrm_enable().
 *
 * @link https://docs.civicrm.org/dev/en/latest/hooks/hook_civicrm_enable
 */
function outlook365_civicrm_enable() {
  _outlook365_civix_civicrm_enable();
}

/**
 * Implements hook_civicrm_disable().
 *
 * @link https://docs.civicrm.org/dev/en/latest/hooks/hook_civicrm_disable
 */
function outlook365_civicrm_disable() {
  _outlook365_civix_civicrm_disable();
}

/**
 * Implements hook_civicrm_upgrade().
 *
 * @link https://docs.civicrm.org/dev/en/latest/hooks/hook_civicrm_upgrade
 */
function outlook365_civicrm_upgrade($op, CRM_Queue_Queue $queue = NULL) {
  return _outlook365_civix_civicrm_upgrade($op, $queue);
}

/**
 * Implements hook_civicrm_managed().
 *
 * Generate a list of entities to create/deactivate/delete when this module
 * is installed, disabled, uninstalled.
 *
 * @link https://docs.civicrm.org/dev/en/latest/hooks/hook_civicrm_managed
 */
function outlook365_civicrm_managed(&$entities) {
  _outlook365_civix_civicrm_managed($entities);
}

/**
 * Implements hook_civicrm_caseTypes().
 *
 * Generate a list of case-types.
 *
 * Note: This hook only runs in CiviCRM 4.4+.
 *
 * @link https://docs.civicrm.org/dev/en/latest/hooks/hook_civicrm_caseTypes
 */
function outlook365_civicrm_caseTypes(&$caseTypes) {
  _outlook365_civix_civicrm_caseTypes($caseTypes);
}

/**
 * Implements hook_civicrm_angularModules().
 *
 * Generate a list of Angular modules.
 *
 * Note: This hook only runs in CiviCRM 4.5+. It may
 * use features only available in v4.6+.
 *
 * @link https://docs.civicrm.org/dev/en/latest/hooks/hook_civicrm_angularModules
 */
function outlook365_civicrm_angularModules(&$angularModules) {
  _outlook365_civix_civicrm_angularModules($angularModules);
}

/**
 * Implements hook_civicrm_alterSettingsFolders().
 *
 * @link https://docs.civicrm.org/dev/en/latest/hooks/hook_civicrm_alterSettingsFolders
 */
function outlook365_civicrm_alterSettingsFolders(&$metaDataFolders = NULL) {
  _outlook365_civix_civicrm_alterSettingsFolders($metaDataFolders);
}

/**
 * Implements hook_civicrm_entityTypes().
 *
 * Declare entity types provided by this module.
 *
 * @link https://docs.civicrm.org/dev/en/latest/hooks/hook_civicrm_entityTypes
 */
function outlook365_civicrm_entityTypes(&$entityTypes) {
  _outlook365_civix_civicrm_entityTypes($entityTypes);
}

/**
 * Implements hook_civicrm_thems().
 */
function outlook365_civicrm_themes(&$themes) {
  _outlook365_civix_civicrm_themes($themes);
}

// --- Functions below this ship commented out. Uncomment as required. ---

/**
 * Implements hook_civicrm_preProcess().
 *
 * @link https://docs.civicrm.org/dev/en/latest/hooks/hook_civicrm_preProcess
 *
function outlook365_civicrm_preProcess($formName, &$form) {

} // */

/**
 * Implements hook_civicrm_navigationMenu().
 *
 * @link https://docs.civicrm.org/dev/en/latest/hooks/hook_civicrm_navigationMenu
 */
function outlook365_civicrm_navigationMenu(&$menu) {
  _outlook365_civix_insert_navigation_menu($menu, 'Administer', array(
    'label' => E::ts('Download Outlook 365 Manifest.xml'),
    'name' => 'outlook365_manifest',
    'url' => 'civicrm/outlook365/manifest.xml',
    'permission' => 'access CiviCRM',
    'operator' => 'OR',
    'separator' => 0,
  ));
  _outlook365_civix_navigationMenu($menu);
} // */
