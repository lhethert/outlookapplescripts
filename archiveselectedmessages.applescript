set myName to "Archive selected messages"
set mailName to "Microsoft Outlook"
set inboxName to "Inbox"
set archiveName to "Archive"

tell application "Microsoft Outlook"
	set frontWin to front window
	set winName to name of frontWin
	set currMsgs to current messages
	if currMsgs = {} then
		display alert ("No selected messages in window: " & winName) message "No messages selected"
		return 0
	end if
	set firstMsg to item 1 of currMsgs
	set onMyComputer to on my computer
	
	# Point to archive folders (Archive/Received and Archive/Sent)
	try
		set archiveFolder to folder archiveName of onMyComputer
		set destSentFolder to folder "Sent" of archiveFolder
		set destRecvFolder to folder "Received" of archiveFolder
	on error errorMessage number errorNumber
		display alert ("Archive folder not found") message (errorMessage & "Error number: " & errorNumber) as critical
		return 0
	end try
	
	# Count messages and notify user
	set msgCount to (count items in currMsgs)
	display notification with title myName subtitle ("Attempting to move " & msgCount & " messages to " & archiveName)
	
	# Iterate over selected messages and archive based on whether they're sent/received
	# and use the date to file messages away
	set defAccount to default account
	set defSenderEmail to email address of defAccount
	try
		repeat with theMessage in currMsgs
			set senderObj to sender of theMessage
			set senderEmail to address of senderObj
			if (senderEmail = defSenderEmail) then
				my archiveMessage(theMessage, destSentFolder)
			else
				my archiveMessage(theMessage, destRecvFolder)
			end if
		end repeat
	on error errorMessage number errorNumber
		display alert "Archiving failed" message (errorMessage & "
Error number: " & errorNumber) as critical
		return 0
	end try
	display notification with title myName subtitle ("Successfully archived " & msgCount & " messages")
end tell

# Subroutine to archive a message to a specific destination folder, creating a sub-folder
# based on the year of the message. Checks to see if the folder exists prior to moving it.
on archiveMessage(theMessage, destFolder)
	tell application "Microsoft Outlook"
		set (is read) of theMessage to true
		set dateRecv to time received of theMessage
		set yearRecv to year of dateRecv
		set yearRecvStr to ("" & yearRecv)
		
		if (exists folder yearRecvStr of destFolder) then
			set theFolder to folder yearRecvStr of destFolder
			move theMessage to theFolder
		else
			try
				make new mail folder at destFolder with properties {name:yearRecvStr}
			on error errorMessage number errorNumber
				display alert "Error" message ("Failed to create folder" & yearRecvStr) as critical
				return 0
			end try
			set newFolder to folder yearRecvStr of destFolder
			move theMessage to newFolder
		end if
	end tell
end archiveMessage

