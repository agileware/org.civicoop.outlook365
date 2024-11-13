<?php

namespace Civi\Api4\Action\Outlook;

use Civi\Api4\Contact;
use Civi\Api4\Email;
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

    protected function findOrRecordContact()
    {
        $result = Email::get()
            ->setSelect([
                'contact_id',
                'contact.display_name',
                'email',
            ])
            ->addWhere('email', '=', trim($this->email))
            ->addWhere('contact_id.is_deleted', '=', FALSE)
            ->addWhere('contact_id.contact_type', '=', 'Individual')
            ->setCheckPermissions(FALSE)
            ->execute();
        if ($result->count() == 0) {
            return $this->recordEmail();
        }
        return $result->first();
    }

    protected function recordEmail()
    {
        $result = Contact::create()
            ->addValue('display_name', $this->full_name)
            ->addValue('contact_type', 'Individual')
            ->addChain('email', Email::create()
                ->addValue('contact_id', '$id')
                ->addValue('email', strtolower(trim($this->email))))
            ->execute();

        return $result->first();
    }
}