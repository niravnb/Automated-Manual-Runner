Set WshShell = WScript.CreateObject("WScript.Shell")

WScript.Sleep 500

Dim n
n = Inputbox("How many test cases you want to execute?", "Automated Manual Runner")

WScript.Sleep 5000

For i = 1 to n

WshShell.SendKeys "^r"
WScript.Sleep 4000	
WshShell.SendKeys "{ENTER}"
WScript.Sleep 4000
WshShell.SendKeys "^a"
WScript.Sleep 4000
WshShell.SendKeys "^q"
WScript.Sleep 4000

Next

MsgBox ("Success! You Executed " & n & " Test cases.  This script is Created By Nirav Bhavsar.")