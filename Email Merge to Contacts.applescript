-- --------------------------------------------------------------------------------
-- Author:	Craig Buchanan <craig.buchanan@cogniza.com>
-- Created:	2013-09-07
-- Purpose:	Send a personalize email message to each Contact in the selected OS X Address Book group or Outlook Category
-- Installation: 
--			1) Drag bundle (Email Merge to Contacts.scptd) to ~/Library/Scripts/ folder
-- Usage:		
--			1) Select AppleScrip menu, then 'Email Merge to Contacts'
-- Todos:
--			1) Handle (Mail) attachments
--			2) Save 'theResult' variables to a CSV file
--			3) Convert theVariables to a Record and adjust dependent script accordingly; move to helper script?
-- --------------------------------------------------------------------------------

-- --------------------------------------------------------------------------------
-- default string localizations
-- --------------------------------------------------------------------------------
set DIALOG_TITLE to "Email Merge"

set BUTTON_OK to "OK"
set BUTTON_CANCEL to "Cancel"
set BUTTON_YES to "Yes"
set BUTTON_NO to "No"

set BUTTON_MAIL to "Mail"
set BUTTON_OUTLOOK to "Outlook"

set PROMPT_TESTING_MODE to "Do you want to run in testing mode?"
set PROMPT_WHAT_ADDRESS to "Send all messages to:"
set ANSWER_EMAIL to "email.merge@mailinator.com"

set PROMPT_WHICH_PROVIDER to "Which provider?"
set PROMPT_WHICH_GROUP to "Which group shall be used?"
set PROMPT_NO_VALID_EMAILS to "The group of contacts contained no valid email addresses."
set PROMPT_WHICH_ADDRESS_TYPE to "Which email-address type(s) to use?"
set PROMPT_WHICH_TEMPLATE to "Which template to use?"
set PROMT_PAUSING to "Pausing between batches.  Processing will resume at %s. (this alert will be automatically dismissed)"
set PROMPT_DONE to "%d emails were sent to %d contacts."

set REASON_NO_ADDRESS to "No email addresses"
set REASON_DUPLICATE_ADDRESS to "Duplicate address"
set REASON_WRONG_ADDRESS_TYPE to "Wrong address type"

-- --------------------------------------------------------------------------------
-- settings
-- --------------------------------------------------------------------------------

set theVariables to {"%%FirstName%%", "%%LastName%%", "%%DisplayName%%", "%%Company%%", "%%EmailAddress%%", "%%Date%%", "%%Time%%"}
-- prevent messages from being perceived as 'spam'
set batchSize to 50 -- # of messages to deliver sequentially; recommended: 50
set thePause to 1 -- # of minutes between batches; recommended: 5

-- --------------------------------------------------------------------------------
-- external libraries
-- --------------------------------------------------------------------------------

--set userHome to (POSIX path of (path to home folder as text) as text)
--set userScriptHome to userHome & "Library/Scripts/"
--set outlookHome to userHome & "Application Support/Microsoft/Office/Outlook Script Menu Items/"

--property Helper : load script (path to resource "HelperLibrary.scpt" in directory "Scripts"
set Helper to load script (path to resource "HelperLibrary.scpt" in directory "Scripts")

-- --------------------------------------------------------------------------------
-- choose plugin
-- --------------------------------------------------------------------------------
display dialog PROMPT_WHICH_PROVIDER with title DIALOG_TITLE buttons {BUTTON_MAIL, BUTTON_OUTLOOK} default button BUTTON_MAIL

if result = {button returned:BUTTON_MAIL} then
	set thePlugin to load script (path to resource "MailPlugin.scpt" in directory "Scripts")
	
else if result = {button returned:BUTTON_OUTLOOK} then
	set thePlugin to load script (path to resource "OutlookPlugin.scpt" in directory "Scripts")
	
end if

-- --------------------------------------------------------------------------------
-- testing mode
-- --------------------------------------------------------------------------------

set testingMode to (display dialog PROMPT_TESTING_MODE with title DIALOG_TITLE buttons {BUTTON_YES, BUTTON_NO} default button BUTTON_YES)
if result = {button returned:BUTTON_NO} then
	set testingMode to false
	
else if result = {button returned:BUTTON_YES} then
	set testingMode to true
	
	set testingAddress to text returned of (display dialog PROMPT_WHAT_ADDRESS default answer ANSWER_EMAIL with title DIALOG_TITLE buttons {BUTTON_OK} default button BUTTON_OK)
	if result is false then return
	
end if

-- --------------------------------------------------------------------------------
-- select group/contacts
-- --------------------------------------------------------------------------------

--generate a List of group names
set groupNames to thePlugin's getGroups()

-- sort the list alphabetically
set groupNames to Helper's simple_sort(the groupNames)

-- prompt to select a group name from the List
set groupName to (choose from list groupNames with title DIALOG_TITLE with prompt PROMPT_WHICH_GROUP)
if result is false then return

set theContacts to thePlugin's getContacts(groupName)
if result is false then return

-- promtp to select which email-address type (e.g. 'work') to use
if (count of theTypes of theContacts) = 0 then
	display dialog PROMPT_NO_VALID_EMAILS with title DIALOG_TITLE buttons {BUTTON_OK} default button BUTTON_OK
	return
end if

set theTypes to (choose from list (theTypes of theContacts) with title DIALOG_TITLE with prompt PROMPT_WHICH_ADDRESS_TYPE with multiple selections allowed)
if result is false then return

-- --------------------------------------------------------------------------------
-- choose template to send
-- --------------------------------------------------------------------------------

-- generate a List of drafts in the 'Drafts' folder
set theSubjectIds to thePlugin's getTemplates()

-- sort the list alphabetically
set theSubjectIds to Helper's simple_sort(the theSubjectIds)

-- prompt to select a group name from the List
set theChoice to (choose from list theSubjectIds with title DIALOG_TITLE with prompt PROMPT_WHICH_TEMPLATE)
if result is false then return

-- extract ID from string
set n to Helper's find("[", theChoice)
set m to Helper's find("]", theChoice)
set theID to (text (n + 1) through (m - 1) of (theChoice as text)) as text

-- --------------------------------------------------------------------------------
-- iterate and send
-- --------------------------------------------------------------------------------

set startTime to (current date)
set iterator to 0 -- tracks batch size

set recordsProcessed to 0 -- the number of contacts
set messagesSent to 0 -- the number of messages actually sent

set theAddressesProcessed to {} -- track email address
set theResults to {} -- track results

repeat with theContact in (theRecords of theContacts)
	repeat 1 times -- fake loop 0
		
		set recordsProcessed to recordsProcessed + 1
		
		-- if the contact doesn't have any email addresses, skip
		if length of (emailAddresses of theContact) = 0 then
			
			set theResult to {displayName:(displayName of theContact), companyName:(companyName of theContact), emailAddress:null, type:null, sent:false, reason:REASON_NO_ADDRESS}
			
			--display dialog Helper's format_string("%s - %s: %s", {(displayName of theContact), "NOT SENT", REASON_NO_ADDRESS})
			
			exit repeat --loop 0
			
		end if
		
		-- otherwise process each email address
		repeat with theEmailAddress in emailAddresses of theContact
			repeat 1 times -- fake loop 1
				
				-- skip if the address type (e.g. 'work') is not one of the chosen values
				if theTypes does not contain (theType of theEmailAddress) then
					
					set theResult to {displayName:(displayName of theContact), companyName:(companyName of theContact), emailAddress:(theAddress of theEmailAddress), type:(theType of theEmailAddress), sent:false, reason:REASON_WRONG_ADDRESS_TYPE}
					
					--display dialog Helper's format_string("%s (%s:%s) - %s: %s", {(displayName of theContact), (theType of theEmailAddress), (theAddress of theEmailAddress), "NOT SENT", REASON_WRONG_ADDRESS_TYPE})
					
					exit repeat --loop 1
					
					-- skip if the address is a duplicate
				else if theAddressesProcessed contains (theAddress of theEmailAddress) then
					
					set theResult to {displayName:(displayName of theContact), companyName:(companyName of theContact), emailAddress:(theAddress of theEmailAddress), type:(theType of theEmailAddress), sent:false, reason:REASON_DUPLICATE_ADDRESS}
					
					--display dialog Helper's format_string("%s (%s:%s) - %s: %s", {(displayName of theContact), (theType of theEmailAddress), (theAddress of theEmailAddress), "NOT SENT", REASON_DUPLICATE_ADDRESS})
					
					exit repeat --loop 1
					
				end if
				
				-- copy the message template's fields
				set theContent to thePlugin's getContent(theID)
				
				-- personalize the fields
				set theContent to Helper's personalize(theContent, theVariables, theContact, (theAddress of theEmailAddress))
				
				-- create the recipient (use testing email address in test mode)
				set theRecipient to {thename:null, theAddress:null}
				if testingMode then
					set theRecipient to {thename:(displayName of theContact), theAddress:testingAddress}
				else
					set theRecipient to {thename:(displayName of theContact), theAddress:(theAddress of theEmailAddress)}
				end if
				
				-- send the message
				thePlugin's copyAndSendMessage(theID, theContent, theRecipient)
				
				set theResult to {displayName:(displayName of theContact), companyName:(companyName of theContact), emailAddress:(theAddress of theEmailAddress), type:(theType of theEmailAddress), sent:true, reason:null}
				
				--display dialog Helper's format_string("%s (%s:%s) - %s", {(displayName of theContact), (theType of theEmailAddress), (theAddress of theEmailAddress), "SENT"})
				
				-- add address to unique list
				if theAddressesProcessed does not contain (theAddress of theEmailAddress) then
					copy (theAddress of theEmailAddress) to the end of theAddressesProcessed
				end if
				
				-- increment
				set iterator to iterator + 1
				-- update statistics
				set messagesSent to messagesSent + 1
				
				-- manage batches
				if ((count of theRecords of theContacts) - recordsProcessed) > 0 and (iterator is batchSize) then
					set iterator to 0
					
					display dialog Helper's format_string(PROMT_PAUSING, {((current date) + (60 * thePause) as text)}) with title DIALOG_TITLE buttons {BUTTON_OK} giving up after 5
					delay (60 * thePause)
				end if
				
			end repeat --/fake 1
		end repeat -- email addresses
		
	end repeat --/fake 0
end repeat --contacts

display dialog Helper's format_string(PROMPT_DONE, {messagesSent, recordsProcessed}) with title DIALOG_TITLE buttons {BUTTON_OK} default button BUTTON_OK

-- {
-- 	firstName: Jane, lastName: Doe, displayName: Jane Doe, companyName: Acme, 
--	emailAddresses:
--	{
--		{theType:Home, theAddress:jane@doe.name},
--		{theType:Work, theAddress:jane.doe@acme.com}
--	}
--}
