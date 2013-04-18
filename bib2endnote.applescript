--- Incremental Endnote updates
--- Peadar Grant
--- Use at your own risk!

(*
    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.
*)

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
