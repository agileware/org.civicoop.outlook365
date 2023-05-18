# Getting started

This page explains how to set up CiviCRM and Outlook 365 integration.

## Enable CiviCRM Extension, Outlook 365 Integration

Go to the CiviCRM Extensions page and enable the CiviCRM extension, **Outlook 365 Integration**.

## Obtain the CiviCRM REST URL, CiviCRM Site Key and CiviCRM API Key

### CiviCRM REST URL

The CiviCRM REST URL is (typically), https://your.domain.name/civicrm/ajax/rest

For more information see [APIv3 REST, End-Point URL](https://docs.civicrm.org/dev/en/latest/api/v3/rest/#end-point-url)

### CiviCRM Site Key

The CiviCRM Site Key is **unique** for each CiviCRM site. To obtain this information, see https://docs.civicrm.org/sysadmin/en/latest/setup/secret-keys/

### CiviCRM API Key

Each user that will be using the Outlook integration will need to have a CiviCRM API Key generated for their CiviCRM Contact. Locate each Contact, click on the API Key tab and click the Add Key API button. (If you have many users, then you will probably want to find a faster way to do this, because life is short!). Each user can also login to CiviCRM and generate their own API Key for their CiviCRM Contact.  

## Provide the CiviCRM REST URL, Site Key and API Key to the end user

Provide the CiviCRM REST URL, CiviCRM Site Key and CiviCRM API Key to each user, as this is required for setting up their Outlook integration. Ensure each user receives the correct CiviCRM API Key.

## CiviCRM Permissions 

For **WordPress**, enable the **Access Outlook 365 pages** permission to **Anonymous User**.

For **Drupal**, enable the **Access Outlook 365 pages** permission to both **Anonymous User** and **Authenticated User** roles.

## Outlook 365 Manifest file (manifest.xml)

In the CiviCRM Menu under *Administer*. Select the **Download Outlook 365 Manifest file.xml** option and download the **manifest.xml** file. This file is used to install the Outlook integration as an **Add-In** into Outlook 365.

## Setting up Outlook 365

There are **two options** for deploying the Outlook 365 Add-Ins, either for all users in the Organisation or on a per user basis.

*Note: Screenshots are not shown for these steps as Microsoft regularly changes the Microsoft 365 admin center.*

### Deploying for the Organisation

1. Go to the [Microsoft 365 admin center](https://admin.microsoft.com/AdminPortal)
2. Click **Settings**
3. Click **Integrated apps**
4. On the Integrated apps page, click the [Add-ins link](https://admin.microsoft.com/Adminportal#/Settings/AddIns)
5. Click the **Deploy Add-in** button
6. Click on Next
7. On the **Deploy a new add-in** page, under the **Deploy a custom add-in** click the **Upload custom apps** button
8. Select **manifest.xml** file from your computer and click **Upload**
9. On the **Configure add-in** page, select options as appropriate for you
10. Click on **Deploy**

The Add-in will be deployed in your Organisation for Outlook 365 users. See the **First use** section below.

### Deploying per user

1. Login into Outlook 365
2. **Open an email** or click on **New message**
3. On the **email page** click on the **...** action button
4. Select the option, **Get Add-Ins**
5. On the **Add-Ins for Outlook** page, select **My add-ins** and scroll to the bottom of the page
6. Select **Add a custom add-in** and select the **Add From file...** option
8. Select **manifest.xml** file from your computer
9. Click **Install** when prompted

The Add-in will now be available in this users Outlook 365. See the *First use configuration* section below.

## First use configuration

The Add-in must be configured for *each user* with the following settings:
* CiviCRM REST URL
* Site Key
* API Key

See the *Obtain the CiviCRM REST URL, Site Key and API Key* section above for more details.

1. Login into Outlook 365
2. **Open an email** or click on **New message**
3. On the **email page** click on the **...** action button
4. Click on icon with the **CiviCRM logo**, your Organisation name should be shown next to this icon
5. The **Add-in pane** will open. Click on the Settings in the **Add-in pane** 
6. Enter the following details: CiviCRM REST URL, CiviCRM Site Key, CiviCRM API Key
6. Click **Save**

If the details have been entered correctly, then the Add-in should now be ready to use.
