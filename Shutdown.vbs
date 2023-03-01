Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")

Set colOS = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")

For Each objOS in colOS
  response = MsgBox("Are you sure you want to shut down your computer?", vbYesNo + vbQuestion, "Shut Down Computer")

  If response = vbYes Then
    psCmd = "powershell.exe -WindowStyle Hidden -Command ""Stop-Computer"""
    
    Set obiShell = CreateObject("WScript.Shell")
    obiShell.Run psCmd, 0, True
  End If
Next