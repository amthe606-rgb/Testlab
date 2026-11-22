' =====================================================================
'  Notepad automation demo
'  Double-click. No install, no console window.
'
'  Shows a script driving ANOTHER application's UI:
'  launch Notepad on a file, move the caret, type, press Ctrl+S,
'  then read the file back to prove the save happened.
' =====================================================================
Option Explicit

Dim sh, fso, p, f, e, pid, s

Set sh  = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

p = sh.ExpandEnvironmentStrings("%TEMP%") & "\notepad-demo.txt"

' -- 1. create the file directly. no app needed for this part ----------
Set f = fso.CreateTextFile(p, True)
f.WriteLine "Created by the payload demo."
f.WriteLine "Notepad will now edit and save this file."
f.Close

' -- 2. launch Notepad on it, keep its PID -----------------------------
Set e = sh.Exec("notepad.exe """ & p & """")
pid = e.ProcessID
WScript.Sleep 1200

' -- 3. focus it, jump to end, type, save ------------------------------
If Not sh.AppActivate(pid) Then
    MsgBox "Could not focus Notepad. Close other windows and retry.", 48, "notepad-demo"
    WScript.Quit
End If
WScript.Sleep 300
sh.SendKeys "^{END}"                 ' Ctrl+End : caret to end of file
WScript.Sleep 200
sh.SendKeys " line added by the payload at " & Time
WScript.Sleep 300
sh.SendKeys "^s"                     ' Ctrl+S : file exists, so no Save As dialog
WScript.Sleep 800

' -- 4. read it back: proof, not assumption ----------------------------
Set f = fso.OpenTextFile(p, 1)
s = f.ReadAll
f.Close

MsgBox "Saved. Current contents:" & vbCrLf & vbCrLf & s & vbCrLf & _
       "File: " & p, 64, "notepad-demo"
