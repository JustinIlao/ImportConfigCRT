#$language = "VBScript"
#$interface = "1.0"


'Inspired by Josh Lowes Scripts(Exportconfig, and Cleardevice) https://github.com/prof-lowe/SecureCRT 
'This script is designed to minimize human error when copy and pasting over configs from documents.
'Future ability to shorten a normal run config into a run brief config, and also get rid of unneccessary run brief commands



Sub main
    crt.Screen.Synchronous = True
    

    'Script from Rob van der Woudes scripting to open file and get file path in vbs.
    Set wShell=CreateObject("WScript.Shell")
    Set oExec=wShell.Exec("mshta.exe ""about:<input type=file id=FILE><script>FILE.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>""")
    TxtFile = oExec.StdOut.ReadLine

    'Object file to read is created
    Set objFileToRead = CreateObject("Scripting.FileSystemObject")
    Set strFileText = objFileToRead.OpenTextFile(TxtFile,1)
    
    
    crt.Screen.Send chr(8) & chr(13)
    
    'Send conf t to enter configuration mode
    If Not (crt.Screen.WaitForString("(config)#", 1)) Then
        crt.Screen.Send "conf t" & vbCrLf
    End If

    'If conf t is sent in privilege level 0, then send enable command first
    If (crt.Screen.WaitForString ("% Invalid input detected at '^' marker.", 1)) Then 
        crt.Screen.Send "en" & vbCrLf
        crt.Screen.Send "conf t" & vbCrLf
    End If

    
    'loop through every line in TxtFile until the end and enter.
    Do While strFileText.AtEndOfStream <> True
        Dim configLine
        configLine = strFileText.ReadLine
        crt.Screen.Send configLine & vbCrLf
    Loop
    strFileText.Close



End Sub