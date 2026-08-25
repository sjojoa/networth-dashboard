-- Net worth dashboard launcher (macOS).
-- Rebuilds NetWorth_Dashboard.html only if the workbook or template changed,
-- then opens it. Locates itself, so the folder can be moved or renamed.
--
-- Build with:  ./launcher/build_launcher.sh

set myPath to POSIX path of (path to me)
if myPath ends with "/" then set myPath to text 1 thru -2 of myPath
set dashDir to do shell script "/usr/bin/dirname " & quoted form of myPath

set builder to quoted form of (dashDir & "/build_dashboard.py")
set page to quoted form of (dashDir & "/NetWorth_Dashboard.html")

try
	do shell script "/usr/bin/python3 " & builder & " --if-needed"
on error errMsg
	display dialog "Could not rebuild the net worth dashboard." & return & return & ¬
		errMsg & return & return & "Opening the last version that was built." ¬
		buttons {"OK"} default button "OK" with icon caution
end try

do shell script "/usr/bin/open " & page
