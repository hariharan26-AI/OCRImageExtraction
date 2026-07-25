' Start App.vbs
' -----------------------------------------------------------------
' Double-click this file to start the app -- no terminal window,
' no typed commands. The browser opens by itself automatically
' (app.py already does that part).
'
' SETUP (do this once):
'   1. Put this file in the SAME folder as app.py.
'   2. pythonPath below already points to your Python install.
'      If you ever reinstall Python somewhere else, update it here.
' -----------------------------------------------------------------

Set objShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

scriptDir  = fso.GetParentFolderName(WScript.ScriptFullName)
pythonw    = "C:\Users\Kishore\AppData\Local\Programs\Python\Python311\pythonw.exe"
appPy      = scriptDir & "\app.py"

If Not fso.FileExists(pythonw) Then
    MsgBox "Could not find pythonw.exe at:" & vbCrLf & pythonw & vbCrLf & vbCrLf & _
           "Open 'Start App.vbs' with Notepad and fix the pythonw path.", _
           vbExclamation, "Setup needed"
    WScript.Quit
End If

If Not fso.FileExists(appPy) Then
    MsgBox "Could not find app.py in this folder:" & vbCrLf & scriptDir & vbCrLf & vbCrLf & _
           "Make sure Start App.vbs is in the SAME folder as app.py.", _
           vbExclamation, "Setup needed"
    WScript.Quit
End If

objShell.CurrentDirectory = scriptDir
objShell.Run """" & pythonw & """ """ & appPy & """", 0, False

