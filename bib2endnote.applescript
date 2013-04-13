--- Incremental Endnote updates
--- Peadar Grant
--- Use at your own risk!

--- Set locations
set bibdeskFile to ((path to home folder as text) & "Dropbox:Bibliography:bibliography.bib")
set endnoteFile to ((path to home folder as text) & "Documents:Bibliography:EndnoteLibrary.enlp")

--- Template for exporting
property templateString : "<records>
<$publications><$endNoteString/></$publications>
</records>"

--- Get the list of publications
log "Getting publications list"
tell application "BibDesk"
	activate
	open bibdeskFile
	set theDoc to get first document
	set publicationList to (get publications of theDoc)
end tell

--- Open Endnote
log "Opening Endnote"
tell application "EndNote X6"
	activate
	set myDB to open endnoteFile
end tell

--- Loop through each, see if it needs to be transferred, and if so then do it.
log "Transferring new records to Endnote"
repeat with thisPublication in publicationList
	tell application "BibDesk"
		set searchString to (get cite key of thisPublication)
	end tell
	tell application "EndNote X6"
		set res to find ("XX:" & searchString & ":XX") in field "Database Provider"
	end tell
	if (count of res) = 0 then
		tell application "BibDesk"
			set exportedText to templated text theDoc using text templateString for thisPublication
		end tell
		tell application "EndNote X6"
			import exportedText into myDB
			set recordSet to (retrieve "shown" records in myDB)
			set thisRecord to item 1 of recordSet
			set field "Database Provider" of record thisRecord to ("XX:" & searchString & ":XX")
			field "Database Provider" of record thisRecord
		end tell
	end if
end repeat

--- Look for orphaned records, and then delete them
log "Deleting orphaned records from Endnote"
tell application "EndNote X6"
	set publicationList to find " "
end tell
repeat with thisPublication in publicationList
	tell application "EndNote X6"
		set searchString to field "Label" of record thisPublication
	end tell
	tell application "BibDesk"
		set occurences to (count of (search theDoc for searchString))
	end tell
	if occurences = 0 then
		tell application "EndNote X6"
			delete record thisPublication
		end tell
	end if
end repeat

log "Complete"
