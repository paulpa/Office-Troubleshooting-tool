'=======================================================================================================
' Name: WICollectFiles.vbs
' Author: Microsoft Customer Support Services
' Copyright (c) 2009, Microsoft Corporation
'
' Collects .msi and .msp files from the Windows Installer cache location "%windir%\Installer"
' and copies them to the specified folder.
' This allows to build up a resiliency location which can be used for the "MspFixUp.vbs"
' solution to recover cached files that have gone missing.
'=======================================================================================================

Option Explicit

Const SOLUTIONNAME                      = "WICollectFiles"
Const SCRIPTBUILD                       = "1.01"

Const MSIOPENDATABASEMODE_READONLY      = 0
Const PID_REVNUMBER                     = 9 'package code for .msi / GUID patch code for .msp

Dim oMsi,oFso,oWShell
Dim sTemp,sWinDir,sWICacheDir,sRestoreLocation,sMessage

Dim sErr_Syntax: sErr_Syntax = vbCrLf & vbTab & "Usage:  " & vbTab&SOLUTIONNAME & ".vbs [SRestoreLocation=<Folder>]" & vbCrLf

'=======================================================================================================

'Init
Set oMsi    = CreateObject("WindowsInstaller.Installer")
Set oFso    = CreateObject("Scripting.FileSystemObject")
Set oWShell = CreateObject("Wscript.Shell")

sTemp       = oWShell.ExpandEnvironmentStrings("%temp%") & "\"
sWinDir     = oWShell.ExpandEnvironmentStrings("%windir%") & "\"
sWICacheDir = sWinDir & "Installer\"

'Parse the command line
ParseCmdLine

'Ensure trailing '\' for sRestoreLocation
If NOT Right(sRestoreLocation, 1) = "\" Then sRestoreLocation = sRestoreLocation & "\"

'Show info dialog
sMessage = "Files are being copied to folder " & sRestoreLocation& vbCrLf & "A Windows Explorer window will open after the script has run."
oWShell.Popup sMessage, 10, "Collect Cached Windows Installer Files"

'Collect the files from WI cache
CollectFiles

'Open the folder with the collected files
oWShell.Run "explorer /e," & chr(34) & sRestoreLocation & chr(34)

'=======================================================================================================

'Command Line Parser
Sub ParseCmdLine

Dim sTmp, sArg
Dim iCnt, iArgCnt
Dim arrArg, arrRestoreLocations
Dim fArray, fIgnoreOnce

On Error Resume Next

sTmp = ""
iArgCnt = WScript.Arguments.Count

If Not iArgCnt > 0 Then 
    'Use defaults
    sRestoreLocation = sTemp&"WIResiliency\"
    If NOT oFso.FolderExists(sRestoreLocation) Then oFso.CreateFolder(sRestoreLocation)
    Exit Sub
End If

For iCnt = 0 To (iArgCnt-1)

    sArg = ""
    sArg = UCase(WScript.Arguments(iCnt))
    If InStr(sArg,"=") > 0 Then
        Set arrArg = Nothing
        arrArg=Split(sArg, "=", 2)
        sArg = arrArg(0)
        fArray = True
    End If
    
    Select Case sArg
    
    Case "/?"
        Wscript.Echo sErr_Syntax
        Wscript.Quit
   
    Case "/SRESTORELOCATION","-SRESTORELOCATION","SRESTORELOCATION" 'SRestoreLocation
        If fArray Then
            sRestoreLocation = arrArg(1)
            fArray = False
        Else
            fIgnoreOnce = True
            sRestoreLocation = WScript.Arguments(iCnt + 1)
        End If
        sRestoreLocation = Replace(sRestoreLocation, ",", ";")
        arrRestoreLocations = Split(sRestoreLocation, ";")
        sRestoreLocation = ""
        sRestoreLocation = arrRestoreLocations(0)
    
    Case Else
        If NOT fIgnoreOnce Then
            sTmp = sTmp & vbCrLf & "Warning: Invalid command line switch '" & WScript.Arguments(iCnt) & "' will be ignored."
            fIgnoreOnce = NOT fIgnoreOnce
        End If
    
    End Select
Next 'iCnt

If Not sTmp = "" Then wscript.echo sTmp

If NOT oFso.FolderExists(sRestoreLocation) Then
    'Use defaults
    sRestoreLocation = sTemp & "WIResiliency\"
    If NOT oFso.FolderExists(sRestoreLocation) Then oFso.CreateFolder(sRestoreLocation)
End If

End Sub 'ParseCmdLine
'=======================================================================================================

Sub CollectFiles

Dim File
Dim sFileNewName, sFileNewPath, sFolder
Dim i

On Error Resume Next

i = 0
For i = 0 To 1
	Select Case i
	Case 0
		sFolder = Replace(wscript.ScriptFullName, wscript.ScriptName, "")
	Case 1
		sFolder = sWICacheDir
	End Select
	
	For Each File in oFso.GetFolder(sFolder).Files
		sFileNewName = "" : sFileNewPath = ""
		Select Case LCase(Right(File.Name, 4))
		Case ".msp"
			sFileNewName = Left(oMsi.SummaryInformation(File.Path, MSIOPENDATABASEMODE_READONLY).Property(PID_REVNUMBER), 38) & ".msp"
		Case ".msi"
			sFileNewName = GetMsiProductCode(File.Path) & "_" & GetMsiPackageCode(File.Path) & ".msi"
		Case Else
		End Select
		If NOT sFileNewName = "" Then
			sFileNewPath = sRestoreLocation & sFileNewName
			If NOT oFso.FileExists(sFileNewPath) Then oFso.CopyFile File.Path, sFileNewPath
		End If
	Next 'File
Next 'i

End Sub 'CollectFiles
'=======================================================================================================
'Obtain the ProductCode (GUID) from a .msi package
'The function will open the .msi database and query the 'Property' table to retrieve the ProductCode

Function GetMsiProductCode(sMsiFile)
    On Error Resume Next
    
    Dim MsiDb, Record
    Dim qView
    
    GetMsiProductCode = ""
    
    Set MsiDb = oMsi.OpenDatabase(sMsiFile,MSIOPENDATABASEMODE_READONLY)
    Set qView = MsiDb.OpenView("SELECT `Value` FROM Property WHERE `Property` = 'ProductCode'")
    qView.Execute
    Set Record = qView.Fetch
    GetMsiProductCode = Record.StringData(1)
    qView.Close

End Function 'GetMsiProductCode
'=======================================================================================================


'Obtain the PackageCode (GUID) from a .msi package
'The function will the .msi'S SummaryInformation stream

Function GetMsiPackageCode(sMsiFile)
    On Error Resume Next
    
    GetMsiPackageCode = ""
    GetMsiPackageCode = oMsi.SummaryInformation(sMsiFile,MSIOPENDATABASEMODE_READONLY).Property(PID_REVNUMBER)

End Function 'GetMsiPackageCode
'=======================================================================================================
