-- Open Kimble in Safari (assumes you are logged in)
-- How this works...
-- Use the csv template to enter the details of the travel you wish to request in Kimble
-- Do NOT muck around with the headings, and don't use any commas. Make sure you save it as a csv, NOT xlsx.  
-- Save that file (don't rename it) in the same folder that this app will live in.  It expects it to be co-located.


tell application "Finder" to set containerFolder to POSIX path of (container of (path to me) as alias)
set csv to "kimble-travel-batch.csv"
log "Importing - " & (containerFolder & csv) as POSIX file

set batchPath to (containerFolder & csv) as POSIX file

-- Read in the csv file
set csvText to read batchPath

-- replace any apostrophes that were put into the csv, otherwise all hell will break loose.
set the csvText to replaceText(csvText, "'", «data utxt02BC» as Unicode text)

set listOfRequests to csvToList(csvText, {separator:","}, {trimming:true})
log (count of listOfRequests)


set numberOfCols to count of (item 1 of listOfRequests)
if numberOfCols is not equal to 28 then error "Oh dear. It looks like you've done something iffy to the csv file!!" -- I TOLD YOU NOT TO ADD ANY COLUMNS!!
set requestHeaders to item 1 of listOfRequests
set listOfRequestsNoHeads to items 2 thru -1 of listOfRequests
set numberOfRequests to count of listOfRequestsNoHeads


log "numberOfRequests - " & numberOfRequests
log "numberOfCols - " & numberOfCols


repeat with theRequest in listOfRequestsNoHeads
	set mainTraveller to item 1 of theRequest
	if mainTraveller is "" then return -- if people edit the csv with Excel, it has a habit of adding a row of ,,,,,,,,, if you ever delete a row out
	set travelSummary to item 2 of theRequest
	set fromDate to item 3 of theRequest
	set toDate to item 4 of theRequest
	set activity to item 5 of theRequest
	set reason to item 6 of theRequest
	set flightBaggage to item 7 of theRequest as boolean
	set flightNeed to item 8 of theRequest as boolean
	set hotelNeed to item 9 of theRequest as boolean
	set trainNeed to item 10 of theRequest as boolean
	set taxiNeed to item 11 of theRequest as boolean
	set carNeed to item 12 of theRequest as boolean
	set OTNeed to item 13 of theRequest as boolean
	set returnFlight to item 14 of theRequest as boolean
	set depAirport to item 15 of theRequest
	set destAirport to item 16 of theRequest
	set retAirport to item 17 of theRequest
	set depTime to item 18 of theRequest
	set retTime to item 19 of theRequest
	set airNotes to item 20 of theRequest
	set hotelPref to item 21 of theRequest
	set hotelNotes to item 22 of theRequest
	set returnTrain to item 23 of theRequest as boolean
	set depStation to item 24 of theRequest
	set destStation to item 25 of theRequest
	set depTrainTime to item 26 of theRequest
	set retTrainTime to item 27 of theRequest
	set trainNotes to item 28 of theRequest
	
	
	
	
	
	tell application "Safari" -- because let's face it, this is hardly going to be robust now, is it?
		try
			tell window 1
				activate
				set current tab to (make new tab with properties {URL:"https://eu1.salesforce.com/a32/o"})
			end tell
			
			delay 3
			--	DatePicker.datePicker.selectDate(this);
			do JavaScript "document.forms['hotlist']['new'].click()" in document 1
			delay 3
			do JavaScript "document.forms['j_id0:j_id1:TheForm']['j_id0:j_id1:TheForm:j_id104:j_id108:j_id109:j_id112'].value = '" & mainTraveller & "'" in document 1
			--	do JavaScript "document.forms['j_id0:j_id1:TheForm'][''].value = '" & mainTraveller & "'" in document 1
			do JavaScript "document.forms['j_id0:j_id1:TheForm']['j_id0:j_id1:TheForm:j_id104:j_id108:j_id117:j_id119'].value = '" & travelSummary & "'" in document 1
			
			do JavaScript "DatePicker.insertDate('" & fromDate & "', 'j_id0:j_id1:TheForm:j_id104:j_id108:j_id120:j_id123',true);" in document 1
			do JavaScript "DatePicker.insertDate('" & toDate & "', 'j_id0:j_id1:TheForm:j_id104:j_id108:j_id125:j_id128',false);" in document 1
			
			delay 3
			
			-- Handling text selectors!
			set theElement to "document.forms['j_id0:j_id1:TheForm']['j_id0:j_id1:TheForm:j_id104:j_id108:j_id130:ActivityList']"
			set textToFind to activity
			do JavaScript "var objSelect = " & theElement & ";setSelectedValue(objSelect, '" & textToFind & "');
function setSelectedValue(selectObj, valueToSet) { for (var i = 0; i < selectObj.options.length; i++) { if (selectObj.options[i].text== valueToSet) { selectObj.options[i].selected = true; return; } }}" in document 1
			
			
			
			--do JavaScript "document.forms['j_id0:j_id1:TheForm']['j_id0:j_id1:TheForm:j_id104:j_id108:j_id130:ActivityList'].selectedIndex = '5';" in document 1
			do JavaScript "document.forms['j_id0:j_id1:TheForm']['j_id0:j_id1:TheForm:j_id104:j_id108:j_id138:j_id140'].value = '" & reason & "';" in document 1
			if flightBaggage is true then
				do JavaScript "document.forms['j_id0:j_id1:TheForm']['j_id0:j_id1:TheForm:j_id104:j_id141:j_id142:0:j_id143'].checked = true;" in document 1
			end if
			-- addRequisitionItem('Flight')  addRequisitionItem('Hotel')  addRequisitionItem('Train')
			
			-- Flights! FLIGHTS! FLIGHTS!!!
			if flightNeed then
				do JavaScript "addRequisitionItem('Flight')" in document 1
				delay 2
				
				-- Handling text selectors! Departure Airport
				set theElement to "document.forms['j_id0:j_id1:TheForm']['j_id0:j_id1:TheForm:j_id104:j_id150:0:j_id170']"
				set textToFind to depAirport
				do JavaScript "var objSelect = " & theElement & ";setSelectedValue(objSelect, '" & textToFind & "');
function setSelectedValue(selectObj, valueToSet) { for (var i = 0; i < selectObj.options.length; i++) { if (selectObj.options[i].text== valueToSet) { selectObj.options[i].selected = true; return; } }}" in document 1
				
				-- Destination Airport
				set theElement to "j_id0:j_id1:TheForm:j_id104:j_id150:0:j_id172"
				set textToFind to destAirport
				do JavaScript "var objSelect = document.forms['j_id0:j_id1:TheForm']['" & theElement & "'];setSelectedValue(objSelect, '" & textToFind & "');
function setSelectedValue(selectObj, valueToSet) { for (var i = 0; i < selectObj.options.length; i++) { if (selectObj.options[i].text== valueToSet) { selectObj.options[i].selected = true; return; } }}" in document 1
				
				
				-- Departure time
				set theElement to "j_id0:j_id1:TheForm:j_id104:j_id150:0:j_id176"
				set textToFind to depTime
				do JavaScript "var objSelect = document.forms['j_id0:j_id1:TheForm']['" & theElement & "'];setSelectedValue(objSelect, '" & textToFind & "');
function setSelectedValue(selectObj, valueToSet) { for (var i = 0; i < selectObj.options.length; i++) { if (selectObj.options[i].text== valueToSet) { selectObj.options[i].selected = true; return; } }}" in document 1
				
				-- Requisition Notes
				set theElement to "j_id0:j_id1:TheForm:j_id104:j_id150:0:j_id189"
				do JavaScript "document.getElementById('" & theElement & "').value='" & airNotes & "';" in document 1
				
				-- Only execute this if we need a return flight			
				if returnFlight is true then
					
					-- Return airport
					set theElement to "j_id0:j_id1:TheForm:j_id104:j_id150:0:j_id181"
					set textToFind to retAirport
					do JavaScript "var objSelect = document.forms['j_id0:j_id1:TheForm']['" & theElement & "'];setSelectedValue(objSelect, '" & textToFind & "');
function setSelectedValue(selectObj, valueToSet) { for (var i = 0; i < selectObj.options.length; i++) { if (selectObj.options[i].text== valueToSet) { selectObj.options[i].selected = true; return; } }}" in document 1
					
					-- Return date	 (use same toDate as the main form - can't think why you wouldn't want to?)
					set theElement to "j_id0:j_id1:TheForm:j_id104:j_id150:0:j_id183"
					do JavaScript "DatePicker.insertDate('" & toDate & "', '" & theElement & "',false);" in document 1
					
					-- Return time
					set theElement to "j_id0:j_id1:TheForm:j_id104:j_id150:0:j_id185"
					set textToFind to retTime
					do JavaScript "var objSelect = document.forms['j_id0:j_id1:TheForm']['" & theElement & "'];setSelectedValue(objSelect, '" & textToFind & "');
function setSelectedValue(selectObj, valueToSet) { for (var i = 0; i < selectObj.options.length; i++) { if (selectObj.options[i].text== valueToSet) { selectObj.options[i].selected = true; return; } }}" in document 1
					
					
					
					-- Single leg - no return?
				else
					set radio1 to "j_id0:j_id1:TheForm:j_id104:j_id150:0:j_id164:0"
					do JavaScript "document.getElementById('" & radio1 & "').click();" in document 1
				end if
			end if
			
			
			-- Hotels! HOTELS! HOTELS!!!
			if hotelNeed then
				do JavaScript "addRequisitionItem('Hotel')" in document 1
				delay 2
				
				-- set the check out date to the toDate!!
				set theElement to "j_id0:j_id1:TheForm:j_id104:j_id194:0:j_id210"
				do JavaScript "DatePicker.insertDate('" & toDate & "', '" & theElement & "',false);" in document 1
				
				-- set hotel pref
				set theElement to "j_id0:j_id1:TheForm:j_id104:j_id194:0:j_id213"
				do JavaScript "document.getElementById('" & theElement & "').value='" & hotelPref & "';" in document 1
				
				-- set hotel notes
				set theElement to "j_id0:j_id1:TheForm:j_id104:j_id194:0:j_id216"
				do JavaScript "document.getElementById('" & theElement & "').value='" & hotelNotes & "';" in document 1
				
				
			end if
			
			-- Trains! TRAINS! TRAINS!!!   returnTrain	depStation	destStation	depTrainTime	retTrainTime	trainNotes
			if trainNeed then
				do JavaScript "addRequisitionItem('Train')" in document 1
				delay 2
				
				
				-- Common to both singles and returns
				-- departure station
				set theElement to "j_id0:j_id1:TheForm:j_id104:j_id221:0:j_id241"
				do JavaScript "document.getElementById('" & theElement & "').value='" & depStation & "';" in document 1
				
				
				-- destination station
				set theElement to "j_id0:j_id1:TheForm:j_id104:j_id221:0:j_id250"
				do JavaScript "document.getElementById('" & theElement & "').value='" & destStation & "';" in document 1
				
				-- Train notes
				set theElement to "j_id0:j_id1:TheForm:j_id104:j_id221:0:j_id263"
				do JavaScript "document.getElementById('" & theElement & "').value='" & trainNotes & "';" in document 1
				
				-- Departure time
				set theElement to "j_id0:j_id1:TheForm:j_id104:j_id221:0:j_id245"
				set textToFind to depTrainTime
				do JavaScript "var objSelect = document.forms['j_id0:j_id1:TheForm']['" & theElement & "'];setSelectedValue(objSelect, '" & textToFind & "');
function setSelectedValue(selectObj, valueToSet) { for (var i = 0; i < selectObj.options.length; i++) { if (selectObj.options[i].text== valueToSet) { selectObj.options[i].selected = true; return; } }}" in document 1
				
				-- Only if single leg option
				if returnTrain is false then
					set radio1 to "j_id0:j_id1:TheForm:j_id104:j_id221:0:j_id235:0"
					do JavaScript "document.getElementById('" & radio1 & "').click();" in document 1
					
					-- Only execute if we need a return leg
				else
					
					
					-- set the return date to the toDate!!
					set theElement to "j_id0:j_id1:TheForm:j_id104:j_id221:0:j_id254"
					do JavaScript "DatePicker.insertDate('" & toDate & "', '" & theElement & "',false);" in document 1
					
					
					-- Return time
					set theElement to "j_id0:j_id1:TheForm:j_id104:j_id221:0:j_id258"
					set textToFind to retTrainTime
					do JavaScript "var objSelect = document.forms['j_id0:j_id1:TheForm']['" & theElement & "'];setSelectedValue(objSelect, '" & textToFind & "');
function setSelectedValue(selectObj, valueToSet) { for (var i = 0; i < selectObj.options.length; i++) { if (selectObj.options[i].text== valueToSet) { selectObj.options[i].selected = true; return; } }}" in document 1
					
					
				end if
				
			end if
			
			
			-- Taxi! TAXI! TAXI!!!   
			if taxiNeed then
				do JavaScript "addRequisitionItem('Taxi')" in document 1
				delay 2
			end if
			
			-- Car! CAR! CAR!!!   
			if carNeed then
				do JavaScript "addRequisitionItem('CarHire')" in document 1
				delay 2
			end if
			
			-- Other traveller!
			if OTNeed then
				do JavaScript "addTraveller()" in document 1
				delay 2
			end if
			
			
			
		on error
			display dialog "Ooops.  Couldn't activate the Safari window.  Make sure you have just one open. Close any Safari developer windows."
			
		end try
		
	end tell
	delay 3
end repeat

set csvText to null
set listOfRequests to null

return


on replaceText(this_text, search_string, replacement_string)
	set AppleScript's text item delimiters to the search_string
	set the item_list to every text item of this_text
	set AppleScript's text item delimiters to the replacement_string
	set this_text to the item_list as string
	set AppleScript's text item delimiters to ""
	return this_text
end replaceText

-- Following cvv to list-of-list script from http://macscripter.net/viewtopic.php?pid=125444#p125444

(* Assumes that the CSV text adheres to the convention:
   Records are delimited by LFs or CRLFs (but CRs are also allowed here).
   The last record in the text may or may not be followed by an LF or CRLF (or CR).
   Fields in the same record are separated by commas (unless specified differently by parameter).
   The last field in a record must not be followed by a comma.
   Trailing or leading spaces in unquoted fields are not ignored (unless so specified by parameter).
   Fields containing quoted text are quoted in their entirety, any space outside them being ignored.
   Fields enclosed in double-quotes are to be taken verbatim, except for any included double-quote pairs, which are to be translated as double-quote characters.
       
   No other variations are currently supported. *)

on csvToList(csvText, implementation)
	-- The 'implementation' parameter must be a record. Leave it empty ({}) for the default assumptions: ie. comma separator, leading and trailing spaces in unquoted fields not to be trimmed. Otherwise it can have a 'separator' property with a text value (eg. {separator:tab}) and/or a 'trimming' property with a boolean value ({trimming:true}).
	set {separator:separator, trimming:trimming} to (implementation & {separator:",", trimming:false})
	
	script o -- Lists for fast access.
		property qdti : getTextItems(csvText, "\"")
		property currentRecord : {}
		property possibleFields : missing value
		property recordList : {}
	end script
	
	-- o's qdti is a list of the CSV's text items, as delimited by double-quotes.
	-- Assuming the convention mentioned above, the number of items is always odd.
	-- Even-numbered items (if any) are quoted field values and don't need parsing.
	-- Odd-numbered items are everything else. Empty strings in odd-numbered slots
	-- (except at the beginning and end) indicate escaped quotes in quoted fields.
	
	set astid to AppleScript's text item delimiters
	set qdtiCount to (count o's qdti)
	set quoteInProgress to false
	considering case
		repeat with i from 1 to qdtiCount by 2 -- Parse odd-numbered items only.
			set thisBit to item i of o's qdti
			if ((count thisBit) > 0) or (i is qdtiCount) then
				-- This is either a non-empty string or the last item in the list, so it doesn't
				-- represent a quoted quote. Check if we've just been dealing with any.
				if (quoteInProgress) then
					-- All the parts of a quoted field containing quoted quotes have now been
					-- passed over. Coerce them together using a quote delimiter.
					set AppleScript's text item delimiters to "\""
					set thisField to (items a thru (i - 1) of o's qdti) as string
					-- Replace the reconstituted quoted quotes with literal quotes.
					set AppleScript's text item delimiters to "\"\""
					set thisField to thisField's text items
					set AppleScript's text item delimiters to "\""
					-- Store the field in the "current record" list and cancel the "quote in progress" flag.
					set end of o's currentRecord to thisField as string
					set quoteInProgress to false
				else if (i > 1) then
					-- The preceding, even-numbered item is a complete quoted field. Store it.
					set end of o's currentRecord to item (i - 1) of o's qdti
				end if
				
				-- Now parse this item's field-separator-delimited text items, which are either non-quoted fields or stumps from the removal of quoted fields. Any that contain line breaks must be further split to end one record and start another. These could include multiple single-field records without field separators.
				set o's possibleFields to getTextItems(thisBit, separator)
				set possibleFieldCount to (count o's possibleFields)
				repeat with j from 1 to possibleFieldCount
					set thisField to item j of o's possibleFields
					if ((count thisField each paragraph) > 1) then
						-- This "field" contains one or more line endings. Split it at those points.
						set theseFields to thisField's paragraphs
						-- With each of these end-of-record fields except the last, complete the field list for the current record and initialise another. Omit the first "field" if it's just the stub from a preceding quoted field.
						repeat with k from 1 to (count theseFields) - 1
							set thisField to item k of theseFields
							if ((k > 1) or (j > 1) or (i is 1) or ((count trim(thisField, true)) > 0)) then set end of o's currentRecord to trim(thisField, trimming)
							set end of o's recordList to o's currentRecord
							set o's currentRecord to {}
						end repeat
						-- With the last end-of-record "field", just complete the current field list if the field's not the stub from a following quoted field.
						set thisField to end of theseFields
						if ((j < possibleFieldCount) or ((count thisField) > 0)) then set end of o's currentRecord to trim(thisField, trimming)
					else
						-- This is a "field" not containing a line break. Insert it into the current field list if it's not just a stub from a preceding or following quoted field.
						if (((j > 1) and ((j < possibleFieldCount) or (i is qdtiCount))) or ((j is 1) and (i is 1)) or ((count trim(thisField, true)) > 0)) then set end of o's currentRecord to trim(thisField, trimming)
					end if
				end repeat
				
				-- Otherwise, this item IS an empty text representing a quoted quote.
			else if (quoteInProgress) then
				-- It's another quote in a field already identified as having one. Do nothing for now.
			else if (i > 1) then
				-- It's the first quoted quote in a quoted field. Note the index of the
				-- preceding even-numbered item (the first part of the field) and flag "quote in
				-- progress" so that the repeat idles past the remaining part(s) of the field.
				set a to i - 1
				set quoteInProgress to true
			end if
		end repeat
	end considering
	
	-- At the end of the repeat, store any remaining "current record".
	if (o's currentRecord is not {}) then set end of o's recordList to o's currentRecord
	set AppleScript's text item delimiters to astid
	
	return o's recordList
end csvToList

-- Get the possibly more than 4000 text items from a text.
on getTextItems(txt, delim)
	set astid to AppleScript's text item delimiters
	set AppleScript's text item delimiters to delim
	set tiCount to (count txt's text items)
	set textItems to {}
	repeat with i from 1 to tiCount by 4000
		set j to i + 3999
		if (j > tiCount) then set j to tiCount
		set textItems to textItems & text items i thru j of txt
	end repeat
	set AppleScript's text item delimiters to astid
	
	return textItems
end getTextItems

-- Trim any leading or trailing spaces from a string.
on trim(txt, trimming)
	if (trimming) then
		repeat with i from 1 to (count txt) - 1
			if (txt begins with space) then
				set txt to text 2 thru -1 of txt
			else
				exit repeat
			end if
		end repeat
		repeat with i from 1 to (count txt) - 1
			if (txt ends with space) then
				set txt to text 1 thru -2 of txt
			else
				exit repeat
			end if
		end repeat
		if (txt is space) then set txt to ""
	end if
	
	return txt
end trim
