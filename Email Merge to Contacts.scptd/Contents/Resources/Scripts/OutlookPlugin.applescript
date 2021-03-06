-- --------------------------------------------------------------------------------
-- Author:	Craig Buchanan <craig.buchanan@cogniza.com>
-- Created:	2013-09-07
-- Purpose:	Useful Microsoft Outlook functions
-- Installation: 
--			1) Drag script (Contacts Library.applescript) to ~/Library/Scripts/Applications/Outlook folder
--			2) Open with AppleScript Editor and compile (⌘K)
--			3) Export as Script (.scpt)
-- --------------------------------------------------------------------------------

-- --------------------------------------------------------------------------------
-- unit testing
-- --------------------------------------------------------------------------------
(*

set theGroups to my getGroups()
(choose from list theGroups with prompt "Which group to use?")
set theChoice to result as text

set theContacts to my getContacts(theChoice)

repeat with theType in (theTypes of theContacts)
	display dialog theType
end repeat

repeat with theContact in (theRecords of theContacts)
	display dialog "FN: " & (firstName of theContact) & "; LN: " & (lastName of theContact) & "; DN: " & (displayName of theContact) & "; CN: " & (companyName of theContact)
	repeat with emailAddress in (emailAddresses of theContact)
		display dialog (theType of emailAddress) & ":" & (theAddress of emailAddress)
	end repeat
end repeat

set theMessage to my getTemplates()
(choose from list theMessage with prompt "Which template to use?")
set theChoice to result as text
display dialog of theChoice

--sendMessage("Outlook - subject", "Message sent by Outlook.", {thename:"First Last", theAddress:"first.last@mailinator.com"})

display dialog (getContent(15002))
*)

-- --------------------------------------------------------------------------------
-- / unit testing
-- --------------------------------------------------------------------------------

-- --------------------------------------------------------------------------------
-- default string localizations
-- --------------------------------------------------------------------------------

--
-- Purpose: return the plugin's name
-- Returns:	Text
--
on getName()
	return "Outlook Plugin"
end getName

--
-- Purpose:	Create a new message with the properties (including attachments) of the selected template, altering the Content and the Recipient
-- Parameters:
--			theID: the ID of the message template
--			theContent: the message's content; string
--			theRecipient: the message's recipient; Record; format: {theName:Jane Doe, theAddress:jane@doe.name}
-- Returns:	nothing
--
on copyAndSendMessage(theID, theContent, theRecipient)
	
	set ADVERTISEMENT to "<br/>Sent with <a href='https://github.com/craibuc/applescript-email-merge'>AppleScript Email Merge (" & my getName() & ")<a/>"
	
	tell application "Microsoft Outlook"
		
		--find the template
		set theTemplate to message id theID
		-- create a new message with the some of the same properties as the template
		set theMessage to make new outgoing message with properties {subject:(subject of theTemplate), content:theContent, attachment:(attachments of theTemplate)}
		-- add plug
		set (content of theMessage) to ((content of theMessage) & ADVERTISEMENT)
		-- add recipient
		make new recipient at theMessage with properties {email address:{name:(thename of theRecipient), address:(theAddress of theRecipient)}}
		-- send the copy
		send theMessage
		
	end tell
	
end copyAndSendMessage

--
-- Purpose:	Open a message given its ID
-- Parameters:
--			theId: message ID
-- Returns:	a Message
--
on getContent(theID)
	
	tell application "Microsoft Outlook"
		return (the content of message id theID) as text
	end tell
	
end getContent

--
-- Purpose:	Assemble a list of subjects of each message in the Drafts folder
-- Returns:	List of Message subjects and IDs. The ID can be used to open the message
--
on getTemplates()
	
	tell application "Microsoft Outlook"
		
		-- generate a List of drafts in the 'Drafts' folder
		set theSubjectIds to {}
		
		set theMessages to every message in the drafts
		repeat with theMessage in theMessages
			-- Subject [ID]
			set theSubjectId to (subject of theMessage as text) & " [" & (id of theMessage as text) & "]"
			copy theSubjectId to end of theSubjectIds
		end repeat
		
		return theSubjectIds
		
	end tell
	
end getTemplates

--
-- Purpose:	Convert selected Contacts to a List of Records
-- Parameters:
--			* categoryName - unused; matches interface of similar function in Contacts Library.scpt
-- Returns:	A Record that contains the distinct list of address types (e.g. 'work') and a List (unsorted) of Records:
--	{
--		theTypes: {'home','work'}
--		theRecords:{
-- 			{ firstName: Jane, lastName: Doe, displayName: Jane Doe, companyName: Acme, emailAddresses: {{theType:"home", theAddress: jane@doe.name},{theType:"work", theAddress: jane.doe@acme.com}} },
--			…
-- 			{ firstName: John, lastName: Doe, displayName: John Doe, companyName: Acme, emailAddresses: {{theType:"home", theAddress: john@doe.name},{theType:"work", theAddress: john.doe@acme.com}} }
--			}
--	}
-- TODOs:	
--			* Filter Contacts by categoryName
--			* Identify primary email address; not exposed by API?

on getContacts(categoryName)
	
	tell application "Microsoft Outlook"
		
		set theContacts to {}
		
		if (categoryName as text) is equal to ("<<Selected Objects>>" as text) then
			set theContacts to selected objects
		else
			display dialog "Search by Category has not been implemented."
			-- TODO: add logic to filter by categoryName
			return false
		end if
		
		-- if there are no contacts selected, warn then quit
		if theContacts is {} then
			display dialog "No contacts have been selected.  Please select one or more contacts in the Contacts pane (⌘3)." with title "Outlook Library" with icon 1
			return false
		end if
		
		--test type of selected items to ensure that they are Contacts
		if class of item 1 of theContacts is not contact then
			display dialog "The items that have been selected are NOT contacts.  Please select one or more contacts in the Contacts pane (⌘3)." with title "Outlook Library" with icon 1
			return false
		end if
		
		--unique list of address types (e.g. 'work')
		set theTypesX to {}
		set theRecordsX to {}
		
		repeat with theContact in theContacts
			
			-- store relevant properties in Record for easy access
			set theRecord to {firstName:null, lastName:null, displayName:null, companyName:null, emailAddresses:null}
			
			set firstName of theRecord to first name of theContact
			set lastName of theRecord to last name of theContact
			set displayName of theRecord to display name of theContact
			set companyName of theRecord to company of theContact
			
			set addressList to {}
			set emailAddresses to {}
			
			set theAddresses to email addresses of theContact
			repeat with theAddress in the theAddresses
--				display dialog theAddress
				-- build unique type list (for theContacts)
				if theTypesX does not contain (type of theAddress as text) then
					--set theTypesX to theTypesX & (type of theAddress as text)
					--copy emailAddress to end of emailAddresses
					copy (type of theAddress as text) to end of theTypesX
				end if
				
				--build unique address list (for theContact)
				--if addressList does not contain (address of theAddress) then
				
				-- {theType:Home, theAddress:jane@doe.name}				
				set emailAddress to {theType:(type of theAddress as text), theAddress:(address of theAddress as text)}
				copy emailAddress to end of emailAddresses
				
				set addressList to addressList & (address of theAddress)
				
				--end if
				
			end repeat
			
			set emailAddresses of theRecord to emailAddresses
			copy theRecord to end of theRecordsX
			
		end repeat
		
		--		return theRecords
		return {theTypes:theTypesX, theRecords:theRecordsX}
		
	end tell
	
end getContacts

--
-- Purpose:	Collect names of Categories; name of function matches equivalent in OS X Contacts Library.scpt (pseudo-interface)
-- Returns:	List (unsorted) of Categories
-- Notes:		"<<Selected Objects>>" allows any artibratry list (e.g. the contents of a Smart Folder) of contacts to be used as a data source
--
on getGroups()
	
	tell application "Microsoft Outlook"
		
		--generate a List of group names
		set groupNames to {"<<Selected Objects>>"}
		
		--	repeat with groupItem in every category
		--		copy (name of groupItem) as text to end of groupNames
		--	end repeat
		
		return groupNames
		
	end tell
	
end getGroups