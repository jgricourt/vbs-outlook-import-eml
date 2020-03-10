'===================================================================
' Description: VBS script to import *.eml files.
'
' Future fix : no sleep time trick for waiting the mail to be openned in Outlook (background process)
' @See https://stackoverflow.com/questions/38969503/shellexecute-and-wait
'
' Third party software : http://patorjk.com/software/taag/#p=display&f=Big&t=Outlook%20-%20import%20*%20EML
'
' Check : 
'
' - Default email client (Outlook) should be selected in Windows configuration @see "Default apps"
' - Default application for opening *.eml files should be Outlook
'
' Author : Jean-Gaël Gricourt 
' Inspired by : Robert Sparnaaij - http://www.howto-outlook.com/howto/import-eml-files.htm
' Version: 1.02
' Website: none
' Github : https://github.com/jgricourt/vbs-outlook-import-eml.git
'
' Usage : cscript import-eml.vbs
'
' Potential issues (already adressed in version 1.01): 
' 
' Multiple email windows opened during execution / error after execution : "Erreur d'exécution Microsoft VBScript Desc: Le serveur distant n'existe pas ou n'est pas disponible"
' Cause : the main Outlook window is not fully openned after calling oOutlook.Session.PickFolder
' Solution(s) : - open Outlook prior executing the script
'               - increase WScript.Sleep time.
'===================================================================

Option Explicit
                                                                                                                                    
Wscript.Echo "  _____ _____ _____ _____ _____ _____ _____ _____ _____ _____ _____ _____ _____ _____ _____ _____ _____ _____ "
Wscript.Echo " |_____|_____|_____|_____|_____|_____|_____|_____|_____|_____|_____|_____|_____|_____|_____|_____|_____|_____|"
Wscript.Echo "    ___        _   _             _              _                            _            _____ __  __ _      "            
Wscript.Echo "   / _ \ _   _| |_| | ___   ___ | | __         (_)_ __ ___  _ __   ___  _ __| |_  __  __ | ____|  \/  | |     "             
Wscript.Echo "  | | | | | | | __| |/ _ \ / _ \| |/ /  _____  | | '_ ` _ \| '_ \ / _ \| '__| __| \ \/ / |  _| | |\/| | |     "             
Wscript.Echo "  | |_| | |_| | |_| | (_) | (_) |   <  |_____| | | | | | | | |_) | (_) | |  | |_   >  <  | |___| |  | | |___  "             
Wscript.Echo "   \___/ \__,_|\__|_|\___/ \___/|_|\_\         |_|_| |_| |_| .__/ \___/|_|   \__| /_/\_\ |_____|_|  |_|_____| "             
Wscript.Echo "                                                           |_|                                                "
Wscript.Echo "                                                                                                              "
Wscript.Echo "  Author : jgricourt@gmail.com                                                                                "
Wscript.Echo "  Relase date : 26/12/2019                                                                                    "
Wscript.Echo "  Version : 1.0                                                                                               "
Wscript.Echo "  _____ _____ _____ _____ _____ _____ _____ _____ _____ _____ _____ _____ _____ _____ _____ _____ _____ _____ "
Wscript.Echo " |_____|_____|_____|_____|_____|_____|_____|_____|_____|_____|_____|_____|_____|_____|_____|_____|_____|_____|"
Wscript.Echo "                                                                                                              "

'Select initial folder in filesystem
Dim oShell : Set oShell = CreateObject("Shell.Application")
Dim oInitialFolder : Set oInitialFolder = oShell.BrowseForFolder(0, "Select the folder containing eml-files", 0)

' Call SearchFolders / ImportEML
If Not (oInitialFolder Is Nothing) Then

  'Select Outlook destination folder  
  Dim oOutlook : Set oOutlook = CreateObject("Outlook.Application")    
  Dim oOutlookFolder : Set oOutlookFolder = oOutlook.Session.PickFolder
  
  'Check if Outlook is running
  On Error Resume Next
  Dim oOutlook2 : Set oOutlook2 = GetObject(, "Outlook.Application")
  If Err.Number = 0 Then
    Wscript.Echo "Outlook is running fine ..."
  Else    
    Wscript.Echo "Outlook is not running then force open ..."
    oOutlook.Session.GetDefaultFolder(6).Display 'Force the full Outlook window to open
    Err.Clear

    'Wscript.Echo "Outlook is not running ... import aborted !"    
    'Wscript.Quit
  End If

  'Select FileSystem source folder
  Dim oFileSystemObject : Set oFileSystemObject = CreateObject("Scripting.FileSystemObject")
  Dim oFolder : Set oFolder = oFileSystemObject.GetFolder(oInitialFolder.Self.path)
  
  If NOT oOutlookFolder Is Nothing Then
    'Delete emails from current folder : oOutlookFolder
    Dim i
    Dim total : total = oOutlookFolder.Items.Count
    Dim oMessage
    For i = 1 To total
      Set oMessage = oOutlookFolder.Items.Item(total - i + 1)
      oMessage.Delete
      Set oMessage = Nothing
    Next

    'Import current folder  
    ImportEML oOutlookFolder, oFolder 
     
    'Search sub folders
    SearchFolders oOutlookFolder, oFolder
  End If

End If

MsgBox "Outlook import completed.", 64, "Import EML"

'===================================================================
' Import EML files in the given folder
'===================================================================

Sub ImportEML(oOutlookFolder, oParentFolder)

Wscript.Echo oParentFolder

Dim oFile

'Read each file of the folder
For Each oFile In oParentFolder.Files

  Dim sExt : sExt = oFileSystemObject.GetExtensionName(oFile.Name)

  'Import eml file to Outlook
  If LCase(sExt) = "eml" Then

    'DEBUG
    'Wscript.Echo oFile

    'Warning : eml files must be associated to Outlook for opening
	  oShell.ShellExecute oFile.Path, "", "", "open", 1
	  WScript.Sleep 250 

    'To be tested : https://www.codeproject.com/Tips/507798/Differences-between-Run-and-Exec-VBScript

    'Move email to Outlook destination folder
    Dim MyInspector : Set MyInspector = oOutlook.ActiveInspector
	  Dim MyItem : Set MyItem = oOutlook.ActiveInspector.CurrentItem
	  MyItem.Move oOutlookFolder   
  End If
Next

End Sub

'===================================================================
' Search folders recursively
'===================================================================

Sub SearchFolders(oOutlookFolder, oParentFolder)

Dim oFileSystem : Set oFileSystem = CreateObject("Scripting.FileSystemObject")
Dim oFolder

With oFileSystem.GetFolder(oParentFolder)

  if .SubFolders.Count > 0 Then
    For each oFolder in .SubFolders
          
      'Create or recreate Outlook next destination folder : oFolder.Name                 
      On Error Resume Next
      Dim oOutlookFolderNext : Set oOutlookFolderNext = oOutlookFolder.Folders.Item(oFolder.Name)
      If Err.Number <> 0 Then 
        Err.clear
      Else               
        oOutlookFolderNext.Delete
      End If
      Set oOutlookFolderNext = oOutlookFolder.Folders.Add(oFolder.Name)  

      'DEBUG
      'Wscript.Echo oOutlookFolderNext & " < " & oFolder

      'Import current folder  
      ImportEML oOutlookFolderNext, oFolder

      'Search sub folders
      SearchFolders oOutlookFolderNext, oFolder

    Next
  End if
End With

End Sub

If Err.Number <> 0 Then 
  ShowError("It failed ...")
Else 
  Wscript.Echo "Success !!!"
End If

Sub ShowError(message)
    WScript.Echo "Error: " & message
    WScript.Echo "Src: " & Err.Source & " Desc: " &  Err.Description
    Err.Clear
End Sub
