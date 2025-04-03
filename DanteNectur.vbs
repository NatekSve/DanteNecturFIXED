vb
Dim objShell
Set objShell = CreateObject("WScript.Shell")

' Display message box
MsgBox "dante boutta take a poop in the toilet", vbExclamation, "Warning"

' Spam errors
For i = 1 To 10
    MsgBox "Error: Something went wrong!", vbCritical, "Error"
Next

' Open annoying sites
For i = 1 To 30
    objShell.Run "https://theannoyingsite.com/"
Next

For i = 1 To 20
    objShell.Run "https://www.youtube.com/watch?v=QF0ACRyNDbM"
Next

' Change screen color (this part is not possible in VBScript, but we can simulate it)
' Instead, we will create a simple HTML page to simulate the effect
Dim fso, htmlFile
Set fso = CreateObject("Scripting.FileSystemObject")
Set htmlFile = fso.CreateTextFile("C:\temp\rainbow.html", True)

htmlFile.WriteLine "<html><body style='background-color: green; color: white; font-size: 50px;'>"
For i = 1 To 100
    htmlFile.WriteLine "DANTE LOVES POOPING<br>"
Next
htmlFile.WriteLine "</body></html>"
htmlFile.Close

objShell.Run "C:\temp\rainbow.html"

' Open Notepad with message
For i = 1 To 23
    Dim notepad
    Set notepad = CreateObject("WScript.Shell")
    notepad.Run "notepad.exe"
    WScript.Sleep 100
    notepad.AppActivate "Notepad"
    WScript.Sleep 100
    notepad.SendKeys "Dante loves to poop in ur face"
Next

' Simulate blue screen (not possible in VBScript, but we can create a message box)
MsgBox "Blue Screen of Death", vbCritical, "System Error"