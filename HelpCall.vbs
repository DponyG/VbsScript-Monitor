
' ***************************************************************************************************
' Declare in Memory
' ***************************************************************************************************

Dim wshNetwork, objFso, objFile, objShell, objFolder, objFolderItem, strComputername
Dim strUserName, strHelpString, strFileName, strFullPath, strFolderName

' ***************************************************************************************************
' We create a Network object and Initalize the computername and the username for the person 
' who opened the file. We then free up resources
' ***************************************************************************************************

Set wshNetwork = WScript.CreateObject("WScript.Network")
strComputername = wshNetwork.ComputerName
strUserName = wshNetwork.UserName
Set wshNetwork = Nothing

' ***************************************************************************************************
' We Initalize a string to hold the HelpMessage.txt file name. We create the Scripting File System
' Object and check to see if the HelpMessage.txt file already exists. If it does we delete it.
' ***************************************************************************************************

strFileName = ".\HelpMessage.txt"
Set objFso = CreateObject("Scripting.FileSystemObject")
If objFso.FileExists(strFileName) Then
    objFso.DeleteFile(strFileName)
End If

' ***************************************************************************************************
' We Get the Absolute Path of the directory this file is in. We search the folder for this file and we
' change its modification date so that HelpCallMiddle will be able detect it. 
' We then free up Resources
' ***************************************************************************************************

strFolderName = ".\"
strFullPath = objFso.GetAbsolutePathName(FolderName)
Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.NameSpace(strFullPath)
Set objFolderItem = objFolder.ParseName("HelpCall.vbs")
objFolderItem.ModifyDate = Now
Set objShell = Nothing
Set objFolder = Nothing
set objFolderItem = Nothing

' ***************************************************************************************************
' We Create a new TextFile with the appropriate username and help message. We then free up resources
' ***************************************************************************************************

Set objFile = objFso.CreateTextFile(strFileName)
objFile.WriteLine(strUserName & " NEEDS YOUR HELP NOW SENT AT : " & Now)
objFile.WriteLine() 
objFile.close()
Set objFile = Nothing
Set objFso = Nothing














