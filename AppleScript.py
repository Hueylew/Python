import subprocess
import time

time.sleep(5)

def get_window_title():
    cmd = """
        use AppleScript version "2.4" -- Yosemite (10.10) or later
        use scripting additions

        tell application "Mail"
	        set selectedMessages to selection
	        if (count of selectedMessages) is equal to 0 then
		    display alert "No Messages Selected" message "Select the messages you want to collect before running this script."
	    end if
	    set theText to ""
	    set theDate to ""
	    set theSubject to ""
	    repeat with theMessage in selectedMessages
		    set theSubject to theSubject & (subject of theMessage) as string
		    set theText to theText & (content of theMessage) as string
		    set theDate to theDate & (date received of theMessage) as string
	    end repeat
        end tell

        tell application "Keyboard Maestro Engine"
	        setvariable "Subject" to theSubject
	        setvariable "Content" to theText
	        setvariable "Date" to theDate
        end tell
    """
    result = subprocess.run(['osascript', '-e', cmd], capture_output=True)
    return result.stdout

print(get_window_title())
