
--
--
--
on thename()
	return "Helper Library"
end thename

--
-- Purpose:	list of substitution variables to be used in messages
--
on variable_list()
	
	--	return {"%%FirstName%%", "%%LastName%%", "%%DisplayName%%", "%%Company%%", "%%EmailAddress%%", "%%Date%%", "%%Time%%"}
	
	-- key/value Record use for personalization
	return {{key:"%%FirstName%%", name:"First Name", abbr:"FN"}, Â
		{key:"%%LastName%%", name:"Last Name", abbr:"LN"}, Â
		{key:"%%DisplayName%%", name:"Display Name", abbr:"DN"}, Â
		{key:"%%Company%%", name:"Company", abbr:"C"}, Â
		{key:"%%EmailAddress%%", name:"Email Address", abbr:"EA"}, Â
		{key:"%%Date%%", name:"Date", abbr:"D"}, Â
		{key:"%%Time%%", name:"Time", abbr:"T"}}
	
end variable_list

--
-- Purpose: replace variables in subject/content text with contact's details
--
on personalize(theText, theVariables, theRecipient, theAddress)
	
	set theValue to null
	
	repeat with theVariable in the theVariables
		
		if theVariable as string is equal to "%%FirstName%%" then
			set theValue to (firstName of theRecipient)
			
		else if theVariable as string is equal to "%%LastName%%" then
			set theValue to (lastName of theRecipient)
			
		else if theVariable as string is equal to "%%DisplayName%%" then
			set theValue to (displayName of theRecipient)
			
		else if theVariable as string is equal to "%%Company%%" then
			set theValue to (companyName of theRecipient)
			
		else if theVariable as string is equal to "%%EmailAddress%%" then
			set theValue to theAddress
			
		else if theVariable as string is equal to "%%Date%%" then
			set theValue to (date string of (current date) as text)
			
		else if theVariable as string is equal to "%%Time%%" then
			set theValue to (time string of (current date) as text)
			
		else
			display dialog theVariable
		end if
		
		set theText to replace(theText, theVariable, theValue)
		
	end repeat
	
	return theText
	
end personalize

--
-- find one character within a set of characters
--
on find(char, source)
	offset of char in source
end find

--
-- Purpose: replace one set of characters with another set, found within a third
--
on replace(input, find, replacement)
	
	set text item delimiters to find
	set ti to text items of input
	set text item delimiters to replacement
	ti as text
	
end replace

--
-- Purpose: sort a List
-- Source: http://macosxautomation.com/applescript/sbrt/sbrt-05.html
--
on simple_sort(my_list)
	
	set the index_list to {}
	set the sorted_list to {}
	repeat (the number of items in my_list) times
		set the low_item to ""
		repeat with i from 1 to (number of items in my_list)
			if i is not in the index_list then
				set this_item to item i of my_list as text
				if the low_item is "" then
					set the low_item to this_item
					set the low_item_index to i
				else if this_item comes before the low_item then
					set the low_item to this_item
					set the low_item_index to i
				end if
			end if
		end repeat
		set the end of sorted_list to the low_item
		set the end of the index_list to the low_item_index
	end repeat
	return the sorted_list
	
end simple_sort

--
-- Purpose:	a printf function for AppleScript
-- Parameters:
--	key_string: text with tokens
--	parametes: List of values
-- Usage:	formatted_string(Ò%d image annotations were copied from %s to %sÓ, {1,"A","B"})
-- Returns: formated text
-- Source:	http://www.tow.com/2006/10/12/applescript-stringwithformat/
--
on format_string(key_string, parameters)
	
	set cmd to "printf " & quoted form of key_string
	
	repeat with i from 1 to count parameters
		set cmd to cmd & space & quoted form of ((item i of parameters) as string)
	end repeat
	return do shell script cmd
	
end format_string
