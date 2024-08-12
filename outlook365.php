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

function outlook365_civicrm_permission(&$permissions) {
  $permissions['access outlook 365 pages'] = [
    'label' => E::ts('Access Outlook 365 pages'),
    'description' => E::ts('Give this permission to anonymous users.'),
  ];
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
 * Implements hook_civicrm_install().
 *
 * @link https://docs.civicrm.org/dev/en/latest/hooks/hook_civicrm_install
 */
function outlook365_civicrm_install() {
  _outlook365_civix_civicrm_install();
}

/**
 * Implements hook_civicrm_enable().
 *
 * @link https://docs.civicrm.org/dev/en/latest/hooks/hook_civicrm_enable
 */
function outlook365_civicrm_enable() {
  _outlook365_civix_civicrm_enable();
}

// --- Functions below this ship commented out. Uncomment as required. ---

/**
 * Implements hook_civicrm_preProcess().
 *
 * @link https://docs.civicrm.org/dev/en/latest/hooks/hook_civicrm_preProcess
 *

 // */

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
