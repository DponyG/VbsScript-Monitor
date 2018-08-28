
' ***************************************************************************************************
' Written By Samuel Grenon on 08/25/2018 
' HelpCallMiddle.vbs
' This vbs script simply monitors a change in the file HelpCall.vbs. Once it finds a modification
' it will read from HelpMessage.txt that HelpCall.vbs creates and print it out to an IE screen for
' 60 seconds. For this Script to Work the user must have read and write permissions in the file
' locations folder. Also note that this file will pull at constant time every 20 seconds.
' ***************************************************************************************************


' ***************************************************************************************************
' Declare in Memory
' ***************************************************************************************************

Dim objFso, objFile, objFileTxt, objIE, strFileText, strFilePath, strFileTextPath

' ***************************************************************************************************
' Set the filePath for the HelpCall script as well as the text file it creates. 
' We then Initalize the Scriping File System and get a reference to HellCall.vbs
' ***************************************************************************************************

strFilePath = ".\HelpCall.vbs"
strFileTextPath = ".\HelpMessage.txt"
Set objFso = CreateObject("Scripting.FileSystemObject")
Set objFile = objFso.GetFile(strFilePath)  


' ***************************************************************************************************
' The Do While loop checks for modifications on HelpCall.vbs. It checks to see if the file
' was modified every 20 seconds. If the Loop catches the change in the modification date
' then it will do some simple file error checking and Open up a message in Internet Explorer
' and free up our resources. If the loop finds a change in the modification date but can 
' not find the text file it is suppose to read from, it will simply print out an error and finish.
' ***************************************************************************************************

Do 

   If DateDiff("s", objFile.DateLastModified, Now) < 20 Then 
        If objFso.FileExists(strFileTextPath) Then
            Set objFileTxt = objFso.OpenTextFile(strFileTextPath)
            strFileText = objFileTxt.ReadAll()
            wscript.echo "FILE MODIFIED"
            objFileTxt.Close()
            outputMessage()
            Set objFileTxt = Nothing
            Set objFso = Nothing
            Set objFile = Nothing
            Exit Do
        Else
            wscript.echo ("HelpMessage.txt could not be found")
            Set objFileTxt = Nothing
            Set objFso = Nothing
            Set objFile = Nothing
            Exit Do
        End If 
    End If

    wscript.Sleep(20000)

Loop 


' ***************************************************************************************************
' outPutMessage reads from the test file HelpMessage.txt and simply prints it to an IE Screen. The 
' window will be open for 60 seconds before closing. After the window closes it will delete the 
' HelloMessage.txt file and free up resources.
' ***************************************************************************************************

Sub outputMessage() 
    Set objShell = CreateObject("Wscript.Shell")
    Set objIE = CreateObject("InternetExplorer.Application")
    objIE.Navigate("about:blank")
    objIE.FullScreen=True
    Set objDoc = objIE.Document.Body
    strHTML = "<H1>"&strFileText&"<H1>"
    objDoc.InnerHTML = strHTML
    objIE.Visible = True
    objIE.StatusBar = False
    objShell.AppActivate "Windows Internet Explorer"
    WScript.Sleep(60000)
    objFso.DeleteFile(strFileTextPath)
    objIE.Quit
    Set objIE = Nothing
End Sub














    

