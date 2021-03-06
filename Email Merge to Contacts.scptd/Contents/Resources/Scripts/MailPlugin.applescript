-- --------------------------------------------------------------------------------
-- Author:	Craig Buchanan <craig.buchanan@cogniza.com>
-- Created:	2013-09-07
-- Purpose:	Useful OS X Mail/Address Book (Contacts) functions
-- Installation: 
--			1) Drag script (Contacts Library.applescript) to ~/Library/Scripts/Applications/Mail folder
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
		display dialog (theType of emailAddress as text) & ":" & (theAddress of emailAddress as text)
	end repeat
end repeat

set theTemplates to my getTemplates()
(choose from list theTemplates with prompt "Which template to use?")
set theChoice to result as text
display dialog theChoice
*)

--set theContent to getContent(100977)
--display dialog theContent

-- with
--copyAndSendMessage(100977, "this message has an attachment", {thename:"Craig Buchanan", theAddress:"craig.buchanan@cogniza.com"})
-- without
--copyAndSendMessage(100978, "this message does NOT have an attachment", {thename:"Craig Buchanan", theAddress:"craig.buchanan@cogniza.com"})
redir(100978)
-- --------------------------------------------------------------------------------
-- / unit testing
-- --------------------------------------------------------------------------------

-- --------------------------------------------------------------------------------
-- default string localizations
-- --------------------------------------------------------------------------------
(*
on redir(theID)	
	tell application "Mail"
		set theRedirectedEmail to redirect the (first message of drafts mailbox whose id = theID)
		tell theRedirectedEmail
			make new to recipient at end of to recipients with properties {name:"Foo Bar", address:"craig.buchanan@cogniza.com"}
			send
		end tell
	end tell
end redir
*)

--
-- Purpose: return the plugin's name
-- Returns:	Text
--

on getName()
	return "OS X Mail/Address Book Plugin"
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
	
	set ADVERTISEMENT to "<br/>Sent with <a href='https://github.com/craibuc/applescript-email-merge'>AppleScript Email Merge (" & my getName() & ") <a/>"
	
	tell application "Mail"
		
		--find the template
		set theTemplate to the (first message of drafts mailbox whose id = theID)
		-- subject
		set theSubject to (subject of theTemplate)
		-- add ad
		set theContent to (theContent & ADVERTISEMENT)
		-- create new message
		set theMessage to make new outgoing message with properties {subject:theSubject, content:theContent}
		-- add recipient
		tell theMessage
			-- add the recipient
			make new to recipient at end of to recipients with properties {name:(thename of theRecipient), address:(theAddress of theRecipient)}
			--
			--			repeat with theAttachment in (mail attachments of theTemplate)
			repeat with theAttachment in theTemplate's mail attachments
				set originalName to name of theAttachment
				display dialog originalName
				--				display dialog (properties of theAttachment)
				--make new attachment at after the last paragraph with properties {theAttachment}
				--				make new attachment with properties {file name:theAttachment} at after the last paragraph
			end repeat
			-- send the copy
			--			send theMessage
		end tell
		
	end tell
	
end copyAndSendMessage

--
-- Purpose:	Open a message given its ID
-- Parameters:
--			theId: message ID
-- Returns:	a Message
--
on getContent(theID)
	
	tell application "Mail"
		
		try
			return the content of (first message of drafts mailbox whose id = theID)
		on error errorText number errorNumber
			display dialog errorText & " [" & errorNumber & "]"
		end try
		
	end tell
	
end getContent

--
-- Purpose:	Assemble a list of subjects of each message in the Drafts folder
-- Returns:	List of Message subjects and IDs; The ID can be used to open the message
--
on getTemplates()
	
	tell application "Mail"
		
		-- generate a List of drafts in the 'Drafts' folder
		set theSubjectIds to {}
		
		set theMessages to every message in drafts mailbox
		repeat with theMessage in theMessages
			-- Subject [ID]
			set theSubjectId to (subject of theMessage as rich text) & " [" & (id of theMessage as rich text) & "]"
			copy theSubjectId to end of theSubjectIds
		end repeat
		
		return theSubjectIds
		
	end tell
	
end getTemplates

--
-- Purpose:	Convert selected Contacts to a List of Records
-- Parameters:
--			* categoryName - unused; matches interface of similar function in Outlook Library.scpt
-- Returns:	A Record that contains the distinct list of address types (e.g. 'work') and a List (unsorted) of Records
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
--			* Filter Contacts by groupName
--			* Identify primary email address; not exposed by API?
--
on getContacts(groupName)
	
	tell application "Contacts"
		
		set theGroup to group (groupName as text)
		set theContacts to people of theGroup
		
		--unique list of address types (e.g. 'work')
		set theTypesX to {}
		set theRecordsX to {}
		
		repeat with theContact in theContacts
			
			-- store relevant properties in Record for easy access
			set theRecord to {firstName:null, lastName:null, displayName:null, companyName:null, emailAddresses:null}
			
			set firstName of theRecord to first name of theContact
			set lastName of theRecord to last name of theContact
			set displayName of theRecord to name of theContact
			set companyName of theRecord to organization of theContact
			
			-- get unique list of email addresses; TODO: identify primary email address
			set addressList to {}
			set emailAddresses to {}
			
			set theAddresses to emails of theContact
			repeat with theAddress in the theAddresses
				
				-- build unique type list (for theContacts)
				if theTypesX does not contain (label of theAddress) then
					set theTypesX to theTypesX & (label of theAddress)
				end if
				
				--build unique address list (for theContact)
				--if addressList does not contain (value of theAddress) then
				
				-- {theType:Home, theAddress:jane@doe.name}
				set emailAddress to {theType:(label of theAddress), theAddress:(value of theAddress)}
				copy emailAddress to end of emailAddresses
				
				set addressList to addressList & (value of theAddress)
				--end if
				
			end repeat
			
			set emailAddresses of theRecord to emailAddresses
			copy theRecord to end of theRecordsX
			
		end repeat
		
		--return theRecords
		return {theTypes:theTypesX, theRecords:theRecordsX}
		
	end tell
	
end getContacts

--
-- Purpose:	Collect names of Groups
-- Returns:	List (unsorted) of Groups
--
on getGroups()
	
	tell application "Contacts"
		
		--generate a List of group names
		set groupNames to {}
		repeat with groupItem in every group
			copy (name of groupItem) as text to end of groupNames
		end repeat
		
		return groupNames
		
	end tell
	
end getGroups