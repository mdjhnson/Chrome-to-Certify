on run {fromEmail, parameters}
	set outputPath to ((path to home folder as string) & "Automation:Send to Certify")
	
	set theResponse to display dialog "What's the receipt for? (This will be the file name)" default answer "" with icon note buttons {"Cancel", "Continue"} default button "Continue"
	--> {button returned:"Continue", text returned:"Receipt Amazon Keyboard"}
	
	tell application "System Events"
		set userName to name of current user
	end tell
	
	tell application "Google Chrome"
		set myWindow to front window
		#set myTabs to tabs in myWindow
		set myTab to active tab in myWindow
		set cDate to do shell script "date +'%Y.%m.%d.%H.%M.%S'"
		
		set outputFolder to "/Users/" & userName & "/Automation/Send to Certify"
		set myCount to 0
		
		activate
		
		#repeat with myTab in myTabs
		set myCount to myCount + 1
		set fileName to text returned of theResponse & " " & cDate & ".pdf"
		log fileName
		
		--N.B.: the following opens the system print window, not Google Chrome’s
		tell myTab to print
		
		tell application "System Events"
			tell process "Google Chrome"
				repeat until window "Print" exists
					delay 0.1
				end repeat
				
				set printWindow to window "Print"
				
				tell printWindow
					set myButton to menu button "PDF"
					click myButton
					
					repeat until exists menu 1 of myButton
						delay 0.1
					end repeat
					
					set myMenu to menu 1 of myButton
					set myMenuItem to menu item "Save as PDF" of myMenu
					click myMenuItem
					
					repeat until exists sheet 1
						delay 0.1
					end repeat
					
					set saveSheet to sheet 1
					tell saveSheet
						set value of first text field to fileName
						keystroke "g" using {command down, shift down}
						
						repeat until exists sheet 1 of saveSheet
							delay 0.1
						end repeat
						
						set goDialogue to sheet 1 of saveSheet
						
						tell goDialogue
							set value of first text field to outputFolder
							#click button "Go"
							delay 0.1
							key code 36
						end tell
						
						
						set value of text field "Save As:" to fileName
						click button "Save"
					end tell
				end tell
				
				repeat while printWindow exists
					delay 0.05
				end repeat
			end tell
		end tell
		#end repeat
	end tell
	
	delay 2
	set theFile to alias (outputPath & ":" & fileName)
	
	tell application "Mail"
		#activate
		set theFrom to fromEmail
		set theTos to {"receipts@certify.com"}
		set theSubject to fileName
		set theContent to "See attached"
		set theAttachment to theFile
		set theDelay to 3
		
		set theMessage to make new outgoing message with properties {sender:theFrom, subject:theSubject, content:theContent & return & return, visible:true}
		tell theMessage
			set visibile to true
			set sender to theFrom
			repeat with theTo in theTos
				make new recipient at end of to recipients with properties {address:theTo}
			end repeat
			#end tell
			#tell content of theMessage
			try
				make new attachment with properties {file name:theAttachment} at after the last word of the last paragraph
				set message_attachment to 0
			on error errmess -- oops
				log errmess -- log the error
				set message_attachment to 1
			end try
			log "message_attachment = " & message_attachment
		end tell
		delay theDelay
		send theMessage
		delay 3
		quit
	end tell
	
	tell application "Finder"
		delete theFile
	end tell
	
	
	display dialog "Receipt [" & fileName & "] sent to certify."
end run
