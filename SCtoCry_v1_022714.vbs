'# SctoCry File Importer vesrion 1.0_022514
'# Language: VBScript
'# Author: D.Ring aka Wölfhelm
'# Date: 022714
'#
'# Update Notes - Changed SC detection code to use a diffrent registry section for installation detection
'#              - Changed SC detection code to check both 64bit(WOW6432) and 32bit registry locations
'#              - corrected a spelling error (determined) 
'#              - Changed z-Zip code to check for both 32bit and 64bit versions
'#              - Added check for \ in zPath
'#              - Changed the method of execution for 7-Zip to a shell command. This allows the window to display the processing it is doing.
'#              - Changing the shell excution also allows the script to shell commands at a higher security level (Commands run like a shortcut on Windows instead of a DOS command)
'# Released under:
'#		    GNU GENERAL PUBLIC LICENSE
'#		       Version 2, June 1991
'# 
'# This script provided is not supported under any standard support program or service. The script is provided AS IS 
'# without warranty of any kind. The author further disclaims all implied warranties including, without limitation, 
'# any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of 
'# the use or performance of the script and documentation remains with you. In no event shall the author or anyone else 
'# involved in the creation, production, or delivery of the script be liable for any damages whatsoever (including, 
'# without limitation, damages for loss of business profits, business interruption, loss of business information, or 
'# other pecuniary loss) arising out of the use of or inability to use the script or documentation, even if the Author 
'# has been advised of the possibility of such damages.

'# Now that the legal stuff is out of the way...

' # Overview:
'# Being new to the CryEngine SDK/Sandbox I was supprise by the diffrent issues I ran into as I attempted to
'# open Star Citizen files with the CryEngine Sandbox editor.  I decided to work towards a completely automated 
'# method of setting up the files needed to view Star Citizen assets and maintain a seperate from the default installation
'# of the CryEngine Sandbox that can be eazily updated as new versions of Star Citizen are released,  
  
'# What the script does
'# Insures Star Citizen is installed finds the install path
'# Insures 7-Zip is installed finds the install path
'# Insures CryEngine Sandbox SDK is installed finds the install path

'# Extracts all Star Citizen files from the StarCitizen\Data directory into a folder in the Cry Engine named GameStarCitizen.
'# Renames all .chrparams file extentions into .smaraprhc (it reverses the name of eaze of identifications and to reverse the process.
'# Creates a game.cfg
'# Creates the folder and file \GameStarCitizen\Materials\material_layers_default.mtl
'# Creates the folder and file \GameStarCitizen\Scripts\physics.lua
'# Logs all extracted files - This log is give a unique name every time it is created.
'# Logs all renaned file extentions
'# Tracks all criticatial functons with logging.
'# Provides user information for each step.

'ELEVATION
If Not WScript.Arguments.Named.Exists("elevate") Then
  CreateObject("Shell.Application").ShellExecute WScript.FullName _
    , WScript.ScriptFullName & " /elevate", "", "runas", 1
  WScript.Quit
End If

'START
Dim strDateTimeStamp, objFSO, strFileCount, strFolderCount, logfile, intFCount1, intFCount2, objLogFile, intTFC
dim strNoRepeats, strLastPath, intRunningFolderCount, intRunningFileCount
Dim zPath
Dim strPathValue
Dim bolExitScript
Dim strSetupPath
Dim strSendPath
Dim strZ7Path
Dim strRenamed

Const ForAppending = 2
strFolderCount = 0
strFileCount = 0

bolExitScript = "False"

'Log the Date
WriteToLog("Date:" & NOW)

'Log the OS version
WriteToLog("System OS: " & CheckForInstall("HKLM","SOFTWARE\Microsoft\Windows NT\CurrentVersion","ProductName"))

If CheckForService("7z.exe") = "running" then
     bolExitScript = "True"
     msgbox "Please close all instances of 7-Zip before running this script. Stopping Script..."
end if

CheckSCInstall
Check7ZipInstall

'Call functions to get the From and To paths

msgbox "Please make sure you run CryEngine Sandbox from it's current location and save a level." _
& vbCrLf & "The file save path is used to determine where CryEngine Sandbox is running from"

strSetupPath = CheckCryEngineInstall
strSendPath = SCPath
strZ7Path = Z7Path

'Unzip and create the files
GoThoughFiles

Set objFSO = CreateObject("Scripting.FileSystemObject")

strDateTimeStamp = replace(now,"/","-")
strDateTimeStamp = replace(strDateTimeStamp," ","_")
strDateTimeStamp = replace(strDateTimeStamp,":","")

LogFile = "C:\SCTrackFile_" & strDateTimeStamp & ".log"
Set objLogFile = objFSO.CreateTextFile(logfile, 2, True)

objStartFolder = CryEngPath & "\GameStarCitizen"
Set objFolder = objFSO.GetFolder(objStartFolder)

intTFC = objFolder.Subfolders.Count 
Set colFiles = objFolder.Files


If bolExitScript <> "True" then
   RenameAndMod
   msgbox "The track file import log will now be created. When the log is complete, a prompt will let you know." _
       & " You can see the log file size growing if you refresh and look at c:\SCTrackFile_ Todays Date and Time.log" 
   GetFilesFromMainFolderONLY(objStartFolder)
   GetFilesFromAllSubFolders objFSO.GetFolder(objStartFolder)
   'Make missing files 
   MakeGameCFG
   MakeMaterialLayers
   CreateFolder(CryEngPath & "\GameStarCitizen\Scripts")
   MakePhysicsLUA
   WriteToLog(strRenamed)
   msgbox "Importing and setup complete. CryEngine Sandbox ready for use." _
           & "Two logs where created the c:\SCtoCryImporter.txt log. This log tracks all importing tasks and data," _
           & "if there was an issue a copy of this log will be helpful." _
           & vbCrLf & "The second file is c:\SCTrackFileDateTime.log it contains trackig data for every file that was imported."
 
end if

sub RenameAndMod
            If bolExitScript = "True" then
               exit sub
            end if
              'wait for the files to be competely unpacked...
              Do while CheckForService("7z.exe") = "running"
                 WScript.Sleep 1000
              Loop
         msgbox ("Importing Process Complete" & vbCrLf _
                 & "In the CryEngine installation folder a directory named GameStarCitizen has been created. " _
                 & "The GameStarCitizen folder contains Star Citizen .pak files that have been extracted into " _
                 & "their own directories seperate from the default CryEngine installation." & vbCrLf _
                 &  vbCrLf _
                 & "Before using these file a few further steps are needed, the script will now work through these steps, assisting with technical issues and completing setup")
  
                 iAnswer = MsgBox("Work around for .chrparams files...." & vbCrLf _
                 & "Before using Star Citizen objects, character animation files need to be renamed or " _ 
                 & "CryEngine will crash. This script renames all .chrparams file extentions to .smaraprhc" & vbCrLf _
                 & "game.cfg...." & vbCrLf _
                 & "The file game.cfg will be created in the GameStarCitizen folder, this file is used by CryEngine" & vbCrLf _
                 & "system.cfg...." & vbCrLf _
                 & "In the CryEngineSDK directory, system.cfg needs to be edited to make GameStarCitizen the default folder " & vbCrLf _
                 & " THE SCRIPT WILL NOT MAKE THIS EDIT you need too:                                              " _
                 & "sys_game_folder=" & chr(34) & "GameStarCitizen" & chr(34) & ""_
                 & "                                  -- sys_game_folder=GameSDK" _
                 & vbCrLf _
                 & "Logging...." & vbCrLf _
                 & "A log entry will be created of every extracted file. " _ 
                 & "The file is C:\SCTrackFile_DateTime.log" & vbCrLf _
                 , vbOKCancel + vbQuestion, "Continue")
              if iAnswer = vbCancel Then 
                 'cancel button was pressed
                  bolExitScript = "True"
                  Exit Sub
              End if
end sub

sub MakeGameCFG
    dim strMake
    strMake = "sys_game_name=AA" & chr(34) & "Star Citizen" & chr(34) & vbCrLf _
            & "sys_localization_folder=Localization" & vbCrLf _
            & "sys_dll_game=" 
            '& "sys_dll_game=CryGameSDK.dll" 
    MakeFile strMake,CryEngPath & "\GameStarCitizen\game.cfg"
end sub

sub MakeMaterialLayers
    dim strMake
    strMake = "<Material MtlFlags=" & chr(34) & "524288" & chr(34) & " Shader=" & chr(34) & "Illum" & chr(34) & " GenMask=" & chr(34) & "2000000000001" & chr(34) & " StringGenMask=" & chr(34) & "%ALLOW_SILHOUETTE_POM%SUBSURFACE_SCATTERING" & chr(34) & " SurfaceType=" & chr(34) & chr(34) & " MatTemplate=" & chr(34) & chr(34) & " Diffuse=" & chr(34) & "0,0,0" & chr(34) & " Specular=" & chr(34) & "0,0,0" & chr(34) & " Emissive=" & chr(34) & "0,0,0" & chr(34) & " Shininess=" & chr(34) & "10" & chr(34) & " Opacity=" & chr(34) & "1" & chr(34) & " LayerAct=" & chr(34) & "1" & chr(34) & ">" & vbCrLf _
            & "<Textures />" & vbCrLf _
            & "<PublicParams GlossFromDiffuseContrast=" & chr(34) & "1" & chr(34) & " FresnelScale=" & chr(34) & "1" & chr(34) & " GlossFromDiffuseOffset=" & chr(34) & "0" & chr(34) & " FresnelBias=" & chr(34) & "1" & chr(34) & " GlossFromDiffuseAmount=" & chr(34) & "0" & chr(34) & " GlossFromDiffuseBrightness=" & chr(34) & "0.333" & chr(34) & " IndirectColor=" & chr(34) & "0.25,0.25,0.25" & chr(34) & "/>" & vbCrLf _
            & "</Material>"
    MakeFile strMake,CryEngPath & "\GameStarCitizen\Materials\material_layers_default.mtl"
end sub

sub GoThoughFiles
    If strSetupPath = "False" then
      If bolExitScript = "True" then
         exit sub
      end if
      msgbox "The path to CryEngine could not be determined. Closing script"
    else
      WriteToLog("Files will be extracted:" & vbCrLf & "From: " & strSendPath & "\CitizenClient\Data" & vbCrLf & "To: " & strSetupPath)
      'Extract the files...
      GetFileNames strZ7Path,strSendPath & "\CitizenClient\Data" ,strSetupPath
    end if
end sub

sub MakePhysicsLUA
    dim strMake
    strMake = "--------------------------------------" & vbCrLf _
               & "-- Dummy file to prevent errors of this file missing - edit as needed" & vbCrLf _
               & "--------------------------------------"
    MakeFile strMake,CryEngPath & "\GameStarCitizen\Scripts\physics.lua"
end sub

sub CheckSCInstall
    'Check if Star Citizen is installed
    'NOTE: The strValue name needs to be left blank to return the Default value of the key. 
    'Check on a 64bit OS...
    If CheckForInstall("HKLM","SOFTWARE\Wow6432Node\Cloud Imperium Games\StarCitizen Launcher.exe","") = "Not Found" then
        'Check on a 32bit OS
        If CheckForInstall("HKLM","SOFTWARE\Cloud Imperium Games\StarCitizen Launcher.exe","") = "Not Found" then
           WriteToLog("Star Citizen not installed")
           msgbox "The installation for Star Citizen was not found." & vbCrLF _
               & "Please install Star Citizen..."
           bolExitScript = "True"
        else
          WriteToLog("Checking for install - Star Citizen installed")
          WriteToLog("     Star Citizen path: " & SCPath)
          'msgbox ("Star Citizen located at: " & SCPath)
        end if
    else
       WriteToLog("Checking for install - Star Citizen installed")
       WriteToLog("     Star Citizen path: " & SCPath)
       'msgbox ("Star Citizen located at: " & SCPath)
    end if
end sub

sub Check7ZipInstall
    'Check if 7-Zip is installed
    'NOTE: The strValue name needs to be left blank to return the Default value of the key. 
    If CheckForInstall("HKLM","SOFTWARE\7-Zip","Path") = "Not Found" then
         '64bit version not found, look for 32bit version
          If CheckForInstall("HKLM","SOFTWARE\Wow6432Node\7-Zip","Path") = "Not Found" then
             WriteToLog("7-Zip not installed")
             msgbox "The installation for 7-Zip was not found." & vbCrLF _
              & "Please install 7-Zip..."
             bolExitScript = "True" 
          else
             WriteToLog("Checking for install - 7-Zip (32bit) installed")
             WriteToLog("     7-Zip path: " & Z7Path)
            'msgbox ("7-Zip located at: " & Z7Path)
    end if   
    else
       WriteToLog("Checking for install - 7-Zip (64bit) installed")
       WriteToLog("     7-Zip path: " & Z7Path)
       'msgbox ("7-Zip located at: " & Z7Path)
    end if
end sub

Function CheckCryEngineInstall
    'Check if CryEngine Sandbox is installed
    'NOTE: The strValue name needs to be left blank to return the Default value of the key. 
    If CheckForInstall("HKCU","Software\Crytek\Sandbox\Recent File List","File1") = "Not Found" then
       WriteToLog("CryEngine not installed")
       msgbox "The CryEngine Sandbox execution path was not found." & vbCrLF _
               & "Please insure CryEngine SDK is installed" & vbCrLF _
               & "AND" & vbCrLF _
               & "Run CryEngine Sandbox and save at least one file." 
           bolExitScript = "True"
           CheckCryEngineInstall = "False"
    else
          'msgbox "CryEngine is installed"
       If ReportFolderStatus(CryEngPath & "\GameStarCitizen") = "False" then
          CreateFolder(CryEngPath & "\GameStarCitizen")
          CheckCryEngineInstall = CryEngPath & "\GameStarCitizen"
       else
          msgbox "The folder " & CryEngPath & "\GameStarCitizen" & " already exist." _
                 & " Decide what you would like to do with the current directory. Then run the script again." _
                 & " The script will quit now."
                     bolExitScript = "True"
         WriteToLog("Script canceled automaticly because an exisiting GameStarCitizen folder was found in the CryEngine directory." _
                     & vbCrLf & "The folder can be renamed or deleted to address this issue.")
         CheckCryEngineInstall = "False"
       end if
    end if
end function

Function CreateFolder(strFolderName)
   Dim objFSO, objFolder, objShell, strDirectory
   Set objFSO = CreateObject("Scripting.FileSystemObject")

   If objFSO.FolderExists(strFolderName) Then
      Set objFolder = objFSO.GetFolder(strFolderName)
      msgbox strFolderName & " already exist." _
             & "Decide what you want to do with the current directory. The script will quit now."
                bolExitScript = "True"
   Else               
       Set objFolder = objFSO.CreateFolder(strFolderName)
       CreateFolder = "True"
   End If

End Function

Function CheckForInstall(Hkey,strKey,strPath) 
  Dim iAnswer

   Const HKEY_CURRENT_USER = &H80000001
   Const HKEY_LOCAL_MACHINE = &H80000002
   Const HKEY_USERS = &H80000003

   strComputer = "."
   Set objRegistry = GetObject("winmgmts:\\" & _ 
       strComputer & "\root\default:StdRegProv")

   strKeyPath = strKey
   strValueName = strPath
  
   If Hkey = "HKLM" then
      objRegistry.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue
   end if

   If Hkey = "HKUR" then
      objRegistry.GetStringValue HKEY_USERS,strKeyPath,strValueName,strValue
   end if

   If Hkey = "HKCU" then
      objRegistry.GetStringValue HKEY_CURRENT_USER,strKeyPath,strValueName,strValue
   end if

   If IsNull(strValue) Then
     Select Case strPath
        Case "ProductName" 
           CheckForInstall = "Not Found"
           iAnswer = MsgBox("The script did not determine what OS this system is using." & vbCrlf _
                 & "Please let the developers know about this issue so it can be " & vbCrlf _
                 &  "addressed. Cancel out of any prompts that follow to exit", vbOKCancel + vbQuestion, "Continue")
              if iAnswer = vbCancel Then 
                 'cancel button was pressed
                  bolExitScript = "True"
                  Exit function
              End If
        Case ""
            CheckForInstall = "Not Found"
                 Exit function
        Case "Path"
            CheckForInstall = "Not Found"
                 Exit function
        Case "File1"
            CheckForInstall = "Not Found"
                 Exit function
        ElseCase
          'nothing 
     end Select
   End If
 CheckForInstall=strValue
end function

Function SCPath
  Dim strPath
  Dim arrKeyWord
  Dim intPlaceCount
  Dim intCount
  Dim strPathVal  
  Dim iK
  Dim ArrSCPath

 intCount = -1
  
  strPath = ReadReg("HKLM\SOFTWARE\Wow6432Node\Cloud Imperium Games\StarCitizen Launcher.exe\")
  'Itemize the path
     ArrSCPath = Split(strPath,"\",-1,vbTextCompare)
	For iK = 0 To UBound(ArrSCPath)
		If ArrSCPath(iK) = "Launcher" then
                  intPlaceCount = ik
                end if
	Next
        Do while intPlaceCount -1 > intCount
           intCount = intCount +1
           If intCount = 0 then
               strPathVal = ArrSCPath(0) 
           else 
               strPathVal = strPathVal  & "\" & ArrSCPath(intCount)
           end if
        Loop
    SCPath = strPathVal
end function

Function Z7Path
  Dim strPath
  strPath = ReadReg("HKLM\SOFTWARE\7-Zip\Path")
  If strPath = "" then
    strPath = ReadReg("HKLM\SOFTWARE\Wow6432Node\7-Zip\Path")
  end if
  Z7Path = Left(strPath,36)
end function

Function CryEngPath
    Dim strPath
    Dim arrKeyWord
    Dim intPlaceCount
    Dim intCount
    Dim strPathVal  
    Dim iK

    intCount = -1

    strPath = ReadReg("HKEY_CURRENT_USER\Software\Crytek\Sandbox\Recent File List\File1") 
    ArrCryEng = Split(strPath,"\")
 
'Itemize the path
     ArrCryEng = Split(strPath,"\",-1,vbTextCompare)
	For iK = 0 To UBound(ArrCryEng)
		If ArrCryEng(iK) = "Levels" then
                  intPlaceCount = ik
                end if
	Next
        Do while intPlaceCount  -2 > intCount
           intCount = intCount +1
           If intCount = 0 then
               strPathVal = ArrCryEng(0) 
           else 
               strPathVal = strPathVal  & "\" & ArrCryEng(intCount)
           end if
        Loop
    CryEngPath = strPathVal
end function

Function ReportFolderStatus(fldr)
   Dim fso, msg
   Set fso = CreateObject("Scripting.FileSystemObject")
   WriteToLog("     Searching for CryEngine Sandbox Objects folder...")
   If (fso.FolderExists(fldr)) Then
      msg = "True"
      WriteToLog("     Objects folder path: " & fldr)
   Else
      msg = "False"
      WriteToLog("     Folder " & fldr & " does not exist.")
   End If
   ReportFolderStatus = msg
End Function

Function ReadReg(RegPath)
     Dim objRegistry, Key
     Dim r 
     Set objRegistry = CreateObject("Wscript.shell")
     On Error Resume Next
      Key = objRegistry.RegRead(RegPath)
     On Error Goto 0
     Err.clear
     If hex(Err.number) = "80070002" Then     
       ReadReg = "Not Installed"
       bolExitScript = "True"
     else   
        ReadReg = Key
     end if
End Function

Sub WinRunOrg(cmd)
     WriteToLog(cmd)
     Set objShell = CreateObject("WScript.Shell") 
     msgbox cmd
     Set objScriptExec = objShell.Exec(cmd) 
     strIpConfig = objScriptExec.StdOut.ReadAll 
     WriteToLog(strIpConfig)
End Sub

Sub WinRun(cmd,args)
    dim app
    Set app = CreateObject("Shell.Application")
    app.ShellExecute cmd, args, "", "runas", 1
    WriteToLog(cmd & " " & args)
End Sub

Sub MoveAFile (strFrom, strTo)
   Dim fso
   Set fso = CreateObject("Scripting.FileSystemObject")
   fso.MoveFile strFrom, strTo
End Sub

Function GetFileNames(zPath,strPathFrom,strPathTo)
  If bolExitScript = "True" then
     exit function
  end if
  Dim fsoFileTest, folder, files, sFolder, strFileList, strexcluded
  Dim iAnswer, fileFolderName, strWaitFlag

  strWaitFlag = "Wait"

  Set fsoFileTest = CreateObject("Scripting.FileSystemObject")
  sFolder = strPathFrom
  
  Set folder = fsoFileTest.GetFolder(sFolder)
  Set files = folder.Files
 
'NOTE: To exclude a file from import, add a Case entry for it.
  For each folderIdx In files
      Select Case folderIdx.Name
'*****************************************************************************
'EXCLUDE FILES - START 
         'Files in this list will not be imported
         'Add all entries below this line
         Case "game.cfg"
               strexcluded = "game.cfg" & VbCrLf & strexcluded
         Case "Videos.pak"
               strexcluded = "Videos.pak" & VbCrLf & strexcluded
         Case "UI.pak"
               strexcluded = "UI.pak" & VbCrLf & strexcluded
         Case "Scripts.pak"
               strexcluded = "Scripts.pak" & VbCrLf & strexcluded
'EXCLUDE FILES - END
'*****************************************************************************
            'Add all entries above this line
         Case Else
              strFileList = folderIdx.Name & vbCrLf & strFileList  
      End select
  Next

iAnswer = MsgBox("These files will be extracted and imported into CryEngine" & VbCrLf _
                & "Importing files from: " & strPathFrom & vbCrlf & vbCrlf & strFileList & vbCrlf & "These files will not be imported - Edit the script in section 'EXCLUDE FILES' to add/remove files from import..." & VbCrLf & VbCrLf & strexcluded & vbCrlf & "Press OK to continue or Cancel to quit", vbOKCancel + vbQuestion, "Continue")
if iAnswer = vbCancel Then 
    'cancel button was pressed
    bolExitScript = "True"
    WriteToLog("Setup was canceled by user...")
    Exit function
End If 
 Msgbox "This process takes a few minutes, do not close the 7z.exe windows, they will close automaticly after processing..."
  For each folderIdx In files
      Select Case folderIdx.Name
         'Add all entries below this line
         Case "game.cfg"
         Case "Videos.pak"
         Case "UI.pak"
         Case "Scripts.pak"
         'Add all entries above this line
         Case Else
             fileFolderName = Left(folderIdx.Name,Len(folderIdx.Name)-4)
               If Right(zPath,1) <> "\" then
                  zPath = zPath & "\"
               end if
              'Run 1 7-Zip at at time.......
              Do while CheckForService("7z.exe") = "running"
                 WScript.Sleep 500
              Loop
              WinRun zPath & "7z.exe"," x " & chr(34) & strPathFrom & chr(34) & "\" & folderIdx.Name & " -o" & chr(34) & strPathTo & "\GameStarCitizen" & chr(34) & " -r -y"
              'Prior command not used currently: WinRun(chr(34) & zPath & chr(34) & "7z x" & chr(34) & strPathFrom & "\" & chr(34) & folderIdx.Name & " -o" & chr(34) & strPathTo & chr(34) & " -r -y")
      End select
  Next
end function

function CheckForService(strServiceName)
   sComputerName = "."
   Set objWMIService = GetObject("winmgmts:\\" & sComputerName & "\root\cimv2")
   sQuery = "SELECT * FROM Win32_Process"
   Set objItems = objWMIService.ExecQuery(sQuery)
   'iterate all item(s)

   For Each objItem In objItems
       If objItem.Name = strServiceName then
         CheckForService = "running"
         exit For
       else 
         CheckForService = "notrunning"
       end if
   Next
end function

Function GetDetailedInfo (strFolder)
    Set objShell = CreateObject ("Shell.Application")
    Set objFolder = objShell.Namespace (strFolder)
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim arrHeaders(13)
    Dim strLogEntry
    For i = 0 to 13
       arrHeaders(i) = objFolder.GetDetailsOf (objFolder.Items, i)
    Next
    For Each strFileName in objFolder.Items
        For i = 0 to 13
            If i <> 9 then
                WriteToLog (arrHeaders(i) _
                    & ": " & objFolder.GetDetailsOf (strFileName, i))
            End If
        Next
        WriteToLog(strLogEntry)
    Next
End Function

sub WriteToLog(strInput)
   Const ForAppending = 8
   Const ForWriting = 2
   Const ForReading = 1
   Const QUOTE = """"

   Dim strFile

   set objFSO = CreateObject("Scripting.FileSystemObject")
   strFile ="C:\SCtoCryImporter.txt"
   If objFSO.FileExists(strFile) Then 
      set objFile = objFSO.OpenTextFile(strFile, ForAppending) 
   Else 
      set objFile = objFSO.CreateTextFile(strFile, False)
   End If

   objFile.WriteLine strInput
   objFile.Close
end sub

sub MakeFile(strInput,strPath)
   Const ForAppending = 8
   Const ForWriting = 2
   Const ForReading = 1
   Const QUOTE = """"

   Dim strFile

   set objFSO = CreateObject("Scripting.FileSystemObject")
   strFile = strPath
   If objFSO.FileExists(strFile) Then 
      set objFile = objFSO.OpenTextFile(strFile, ForAppending) 
   Else 
      set objFile = objFSO.CreateTextFile(strFile, False)
   End If

   objFile.WriteLine strInput
   objFile.Close
end sub

Function GetFilesFromMainFolderONLY(strPath)
  On Error Resume Next
  Dim fso, foldersub, folder, ofolder, files, NewsFile, sFolder, sFile
 
  Set fso = CreateObject("Scripting.FileSystemObject")  
  Set folder = fso.GetFolder(strPath)  
  Set files = folder.Files    
  strFolderCount = 1

  objLogFile.Write strFolderCount & " " & strPath
  objLogFile.Writeline
  For each folderIdx In files 
    strFileCount = strFileCount + 1   
    objLogFile.Writeline("      " & strFileCount & " " & strPath & folderIdx.Name) 
  Next
  intRunningFolderCount = strFolderCount
  intRunningFileCount = strFileCount
  strFileCount = 0
End Function

Sub GetFilesFromAllSubFolders(Folder)
    If strNoRepeats <> Folder then
       if strFolderCount <> 1 then
          objLogFile.Write strFolderCount & " " & Folder
          objLogFile.Writeline
       end if
    end if
    For Each Subfolder in Folder.SubFolders
        strFolderCount = strFolderCount + 1
        If intTFC = Folder.Subfolders.Count  then
            objLogFile.Write strFolderCount & " " & Folder & "\" & Subfolder.Name & vbCrLf
            'Rename files ending with .chrparams to prevent the 2 files loaded error in CryEngine
            FileRename(Folder & "\" & Subfolder.Name)
            strNoRepeats = Folder & "\" & Subfolder.Name
            objLogFile.Writeline
        end if
            
        Set objFolder = objFSO.GetFolder(Subfolder.Path) 
        Set colFiles = objFolder.Files
        For Each objFile in colFiles
            strFileCount = strFileCount + 1
            objLogFile.Write "      " & strFileCount & " " & Subfolder.Path & "\" & objFile.Name & vbCrLf
        Next
            'objLogFile.Write GetFileInfo(Subfolder.Path & "\" & objFile.Name)
            FileRename(Subfolder.Path)
            intRunningFileCount = intRunningFileCount + strFileCount
            intRunningFolderCount = strFolderCount
            strFileCount = 0
            GetFilesFromAllSubFolders Subfolder
    Next
End Sub

function GetFileInfo(strfilename)
    strComputer = "."
    dim strdatalog
    dim strToSQL

    Set objWMIService = GetObject("winmgmts:" _
        & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    strToSQL = Replace(strfilename,"\","\\")
    Set colFiles = objWMIService.ExecQuery _
        ("Select * from CIM_Datafile Where name = '" & strToSQL & "'")

    For Each objFile in colFiles
        strdatalog = strdatalog & VbCrLf
        strdatalog = strdatalog & "            File name: " & objFile.FileName & "." & objFile.Extension & vbCrLf
        strdatalog = strdatalog & "            Creation date: " & objFile.CreationDate & vbCrLf
        strdatalog = strdatalog & "            File size: " & objFile.FileSize & vbCrLf
       'strdatalog = strdatalog & "            Path: " & objFile.Name & vbCrLf
        GetFileInfo = strdatalog
    Next
end function

sub FileRename(strFolder)
    Dim objFSO, objFolder, objFile, strNewName, strOldName
    Dim strPath, strName

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(strFolder)
    For Each objFile In objFolder.Files
      If (Right(LCase(objFile.Name), 10) = ".chrparams") Then 
        ' Rename the file.
        strOldName = objFile.Path
        strPath = objFile.ParentFolder
        strName = objFile.Name
        strRenamed = "Renamed: " & strOldName & "\" & strName & VbCrLf & strRenamed
        ' Change name by changing extension to ".smaraprhc"
        strName = Left(strName, Len(strName) - 10) & ".smaraprhc"
        strNewName = strPath & "\" & strName
        ' Rename the file.
        objFSO.MoveFile strOldName, strNewName
      else
        'strRenamed = "Not Renamed: " & strFolder & "\" & objFile.Name & VbCrLf & strRenamed
      End If
     Next
end sub