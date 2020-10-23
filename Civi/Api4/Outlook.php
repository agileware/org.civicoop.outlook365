<?php

namespace Civi\Api4;

use Civi\Api4\Generic\BasicGetFieldsAction;

/**
 * Example of an _ad-hoc_ API.
 *
 * This demonstrates how to create your own API for any arbitrary data source,
 * implementing all the usual actions + a few extras.
 *
 * Try creating some Example records in the _API Explorer_ using `Create` or `Save`.
 * Then try `Get` (with `select`, `where`, `orderBy`, etc.),
 * `Update`, `Replace` and `Delete` too!
 *
 * The API entity `Example` is "declared" simply by the presence of our `Civi\Api4\Example` class.
 * Just that one file is all that's needed for an API to exist; but non-trivial APIs will usually
 * be organized into different files, as we've done in this extension.
 *
 * _The "@method" annotation helps IDEs with the virtual action provided by our `Random` class._
 * @method static Action\Example\Random random()
 * _Annotations for virtual methods are helpful but not required_
 * _(when mixing an action into an entity outside this extension, it wouldn't be possible)._
 *
 * **Note:** this docblock will appear in the _API Explorer_, as will these links:
 * @see https://lab.civicrm.org/extensions/api4example
 * @see https://docs.civicrm.org/dev/en/latest/api/v4/architecture/
 *
 * @package Civi\Api4
 */
class Outlook extends Generic\AbstractEntity {


	/**
	 * @param bool $checkPermissions
	 *
	 * @return \Civi\Api4\Generic\BasicGetFieldsAction
	 */
	public static function getFields($checkPermissions = TRUE) {
		return ( new BasicGetFieldsAction( __CLASS__, __FUNCTION__, function ( $getFieldsAction ) {
			return [
			];

		} ) )->setCheckPermissions( $checkPermissions );
	}

	public static function get($checkPermissions = TRUE) {
		return (new Generic\BasicGetAction(__CLASS__, __FUNCTION__, function($actionObject) {
			return [ 1 => 'something' ];
		}))->setCheckPermissions($checkPermissions);
	}
}