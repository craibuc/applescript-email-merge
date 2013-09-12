applescript-email-merge
========================

AppleScripts to send personalized email messages to contacts using OS X Mail/Address Book or Microsoft Outlook.

Installation
------------
- Download source and decompress
- Copy or move 'Email Merge to Contact.scptd' bundle to ~/Library/Scripts/ folder.

Usage
-----
- Create new mail message and save it to the 'Drafts' folder; include replacement variables as desired
- Choose 'Email Merge to Contacts' from OS X's AppleScript menu
- Follow instructions

###Variables
 - %%FirstName%% - Contact's first name (e.g. 'Foo')
 - %%LastName%% - Contact's last name (e.g. 'Bar')
 - %%DisplayName%% - Contact's full name (e.g. 'Foo Bar')
 - %%Company%% - Contact's (e.g. 'Acme')
 - %%EmailAddress%% - Contact's email address (e.g. 'foobar@acme.com')
 - %%Date%% - Date the message was sent
 - %%Time%% - Time the message was sent
 
Features
--------
 - Multiple email addresses/contact are supported; types (e.g. 'work') are chosen during the process
 - No duplicate messages are sent

Issues
------
 - Attachments to templates aren't sent by OS X Mail plugin
 
Enhancements
------------
 - Fix OS X plugins handling of template's attachments
 - Log activity to CSV
 - Filter Outlook Contacts by Category
 - Use Smart * as Lists when/if they are exposed by respective APIs
 - Installation script?
 
Contributors
------------
- [Craig Buchanan](https://github.com/craibuc)

You're welcome to fork this project and send pull requests.
