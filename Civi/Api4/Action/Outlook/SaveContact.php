<?php

namespace Civi\Api4\Action\Outlook;

use Civi\Api4\Generic\AbstractAction;
use Civi\Api4\Generic\Result;

class SaveContact extends AbstractAction {

	/**
	 * @var string
	 */
	protected $email;

	/**
	 * @var string
	 */
	protected $full_name;


	public static function fields() {
		return [
			['name' => 'email'],
			['name' => 'full_name']
		];
	}

	/**
	 * @param \Civi\Api4\Generic\Result $result
	 */
	public function _run( Result $result ) {
		$result[] = $this->findOrRecordContact();
	}

	protected function findOrRecordContact() {
		$result = \Civi\Api4\Email::get()
		                        ->setSelect([
		                        	'contact_id',
			                        'contact.display_name',
			                        'email',
		                        ])
		                        ->addWhere('email', '=', $this->email)
		                        ->setCheckPermissions(FALSE)
		                        ->execute();
		if ($result->count() == 0 ){
			return $this->recordEmail();
		}
		return $result->first();
	}

	protected function recordEmail() {
		$result = \Civi\Api4\Contact::create()
			->addValue( 'display_name', $this->full_name)
			->addChain( 'email', \Civi\Api4\Email::create()
				->addValue( 'contact_id', '$id')
				->addValue( 'email', $this->email) )
			->execute();

		return $result->first();
	}
}