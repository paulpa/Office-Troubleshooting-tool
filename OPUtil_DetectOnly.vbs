'*******************************************************************************
' Name: OPUtil.vbs - Office Patch Utility
' Author: Microsoft Customer Support Services
' Copyright (c) Microsoft Corporation
' 
' Utility for Office patch maintenance tasks 
' "view, log, repair, apply, remove, clean"
' Formerly published as MspFixUp.vbs
'*******************************************************************************


'-------------------------------------------------------------------------------
'[INI] Section for script behavior customizations

'Directory for Log output.
'Example: "\\<server>\<share>\"
'Default: sPathOutputFolder = vbNullString -> %temp% directory is used
Dim sPathOutputFolder
sPathOutputFolder = ""

'Quiet switch.
'Default: False -> A summary log opens automatically when done
Dim bQuiet
bQuiet = True

'Set fDetectOnly to 'True' if only a log should be generated
'None of the detected actions required will be executed!
'Default: False -> execute detected actions required
Dim fDetectOnly
fDetectOnly = True

'Optional location to provide .msp patch files that should be applied
'A list of full path references to folders with .msp files, separated by semicolons
'Default: sUpdateLocation = ""
Dim sUpdateLocation
sUpdateLocation = ""

'Optional location to restore .msi & .msp packages that have gone missing
'A list of fully qualified paths to folders with .msi and/or .msp files, separated by semicolons
'Default: sUpdateLocation = ""
Dim sRestoreLocation
sRestoreLocation = ""

'Option to explicitly exclude the Windows Installer cache when searching for applicable .msp files.
'This allows to enforce patches are only applied from the provided SUpdateLocation folders
'Note: For detection integrity it's still required to include the patches in the sequence logic!
'Default: fExcludeCache = False -> Scans %WinDir%\Installer folder for applicable patches
Dim fExcludeCache
fExcludeCache = False

'Option to include OCT patches from the Windows Installer cache.
'This option is a subset of the above 'fExcludeCache' option
'it allows to include that cached (installed) OCT patches get applied
'NOTE: It's recommended to keep this set to 'False' unless you have a specific requirement to enforce this
'Default: fIncludeOctCache = False -> Filters OCT patches from %WinDir%\Installer folder
Dim fIncludeOctCache
fIncludeOctCache = False

'Check the integrity of the local Windows Installer cache and try to repair missing .msi and .msp files if needed
'Default: fRepairCache = True -> Try to repair missing files
Dim fRepairCache
fRepairCache = True

'Unregister patches that have gone missing from the local Windows Installer cache to unblock maintenance transactions
'Default: fReconcileCache = False -> Don't unregister missing patches
Dim fReconcileCache
fReconcileCache = False

'Apply .msp patch files
'Patch files are applied from the optional SUpdateLocation folders and the local Windows Installer cache.
'Note: To fine tune the behavior see the options for SUpdateLocation, fExcludeCache, fIncludeOctCache
'Default: fApplyPatch = True -> Apply patches
Dim fApplyPatch
fApplyPatch = True

'Control the behavior of MsiRestartManager
'Default: fDisableRestartManager = True -> Disable Restart Manager 
Dim fDisableRestartManager
fDisableRestartManager = True

'Remove installed patches
'Allows to uninstall superseded patches or a specified list.
'The list allows passing in a KB number(s) or PatchCode(s)
'Default: fRemovePatch = False -> Do not attempt to uninstall patches
'         sMspRemoveFilter  = "Superseded" -> If enabled default to remove superseded patches
'         sMspProductFilter = ""           -> If enabled remove from all products
Dim fRemovePatch, sMspRemoveFilter, sMspProductFilter
fRemovePatch = False
  sMspRemoveFilter = "Superseded"
  sMspProductFilter= ""

'Delete cached .msp files that are no longer referenced by any product from the local Windows Installer cache
Dim fCleanCache
fCleanCache = False

'Suppress debug logging details from the log
'Default: fNoDebugLog = False -> Add debug logging information
Dim fNoDebugLog
fNoDebugLog = False


'DO NOT CUSTOMIZE BELOW THIS LINE!
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
' Global Declarations
'-------------------------------------------------------------------------------

Const SOLUTIONNAME                      = "OPUtil"
Const SCRIPTBUILD                       = "3.17"

Const FOR_READING                       = 1
Const FOR_WRITING                       = 2
Const TRISTATE_USEDEFAULT               = -2
Const LEN_GUID                          = 38
Const USERSID_NULL                      = ""
Const USERSID_EVERYONE                  = "s-1-1-0"

Const PID_TITLE                         = 2
Const PID_SUBJECT                       = 3 'Displayname
Const PID_TEMPLATE                      = 7 'compatible platform and language versions for .msi / PatchTargets for .msp
Const PID_LASTAUTHOR                    = 8 'Transform Substorages
Const PID_REVNUMBER                     = 9 'package code for .msi / GUID patch code for .msp

Const MSIINSTALLSTATE_LOCAL             = 3

Const MSIINSTALLCONTEXT_USERMANAGED     = 1
Const MSIINSTALLCONTEXT_USERUNMANAGED   = 2
Const MSIINSTALLCONTEXT_MACHINE         = 4
Const MSIINSTALLCONTEXT_ALL             = 7

Const MSIPATCHSTATE_APPLIED             = 1
Const MSIPATCHSTATE_SUPERSEDED          = 2
Const MSIPATCHSTATE_OBSOLETED           = 4
Const MSIPATCHSTATE_REGISTERED          = 8
Const MSIPATCHSTATE_ALL                 = 15

Const MSIOPENDATABASEMODE_READONLY      = 0
Const MSIOPENDATABASEMODE_PATCHFILE     = 32

Const MSIREADSTREAM_ANSI                = 2

Const HKCR                              = &H80000000
Const HKCU                              = &H80000001
Const HKLM                              = &H80000002
Const HKU                               = &H80000003

Const MSP_NOSEQ                         = 0
Const MSP_NOBASE                        = 1
Const MSP_MINOR                         = 2
Const MSP_SMALL                         = 3

Const COL_FILENAME                      = 0
Const COL_TARGETS                       = 1
Const COL_PATCHCODE                     = 2
Const COL_SUPERSEDES                    = 3
Const COL_KB                            = 4
Const COL_PACKAGE                       = 5
Const COL_RELEASE                       = 6
Const COL_SEQUENCE                      = 7
Const COL_FAMILY                        = 8
Const COL_PATCHXML                      = 9
Const COL_PATCHTABLES                   = 10
Const COL_REFCNT                        = 11
Const COL_APPLIEDCNT                    = 12
Const COL_SUPERSEDEDCNT                 = 13
Const COL_APPLICABLECNT                 = 14
Const COL_NOQALBASELINECNT              = 15
Const COL_PATCHBASELINES                = 16
Const COL_MAX                           = 16

Const REG_GLOBALCONFIG                  = "Software\Microsoft\Windows\CurrentVersion\Installer\UserData\"
Const REG_PRODUCT                       = "Software\Classes\Installer\"
Const REG_PRODUCTPERUSER                = "Software\Microsoft\Installer\"
Const REG_PRODUCTPERUSERMANAGED         = "Software\Microsoft\Windows\CurrentVersion\Installer\Managed\"

Const ERR_REBOOT                        = "A reboot is required to complete the update(s)!"

Const OFFICE_ALL                        = "78E1-11D2-B60F-006097C998E7}.0001-11D2-92F2-00104BC947F0}.6000-11D3-8CFE-0050048383C9}.6000-11D3-8CFE-0150048383C9}.7000-11D3-8CFE-0150048383C9}.BE5F-4ED1-A0F7-759D40C7622E}.BDCA-11D1-B7AE-00C04FB92F3D}.6D54-11D4-BEE3-00C04F990354}.CFDA-404E-8992-6AF153ED1719}."
'Office 2000 -> KB230848; Office XP -> KB302663; Office 2003 -> KB832672
Const OFFICE_2000                       = "78E1-11D2-B60F-006097C998E7}"
Const ORK_2000                          = "0001-11D2-92F2-00104BC947F0}"
Const PRJ_2000                          = "BDCA-11D1-B7AE-00C04FB92F3D}"
Const VIS_2002                          = "6D54-11D4-BEE3-00C04F990354}"
Const OFFICE_2002                       = "6000-11D3-8CFE-0050048383C9}"
Const OFFICE_2003                       = "6000-11D3-8CFE-0150048383C9}"
Const WSS_2                             = "7000-11D3-8CFE-0150048383C9}"
Const MOSS_2003                         = "BE5F-4ED1-A0F7-759D40C7622E}"
Const PPS_2007                          = "CFDA-404E-8992-6AF153ED1719}" 'Project Portfolio Server 2007
Const OFFICEID                          = "000-0000000FF1CE}" 'cover O12, O14 with 32 & 64 bit
Const OFFICEDBGID                       = "000-1000000FF1CE}" 'Office Debug O12, O14 with 32 & 64 bit

Const xlSrcRange    = 1
Const xlYes         = 1
Const xlNo          = 2
Const xlMaximized   = &HFFFFEFD7

Const HROW          = 1 'Header Row

Const SEQ_PATCHFAMILY = 1
Const SEQ_PRODUCTCODE = 2
Const SEQ_SEQUENCE = 3
Const SEQ_ATTRIBUTE = 4

Const MET_COMPANY = 1
Const MET_PROPERTY = 2
Const MET_VALUE = 3

Const F_FILE = 1
Const F_COMPONENT = 2
Const F_FILENAME = 3
Const F_FILESIZE = 4
Const F_VERSION = 5
Const F_HASH = 6
Const F_LANGUAGE  = 7
Const F_ATTRIBUTE = 8
Const F_SEQUENCE = 9
Const F_PREDICTED = 10
Const F_COMPSTATE = 11
Const F_CURSIZE = 12
Const F_CURVERSION = 13
Const F_CURHASH = 14
Const F_FILEPATH = 15

Const S_PROP = 1
Const S_VAL  = 2

Const COL_TPC   = 0 'TargetProductCode
Const COL_TPCV  = 1 ' Validate
Const COL_TV    = 2 'TargetVersion
Const COL_TVV   = 3 ' Validate
Const COL_TVCT  = 4 ' ComparisonType
Const COL_TVCF  = 5 ' ComparisonFilter
Const COL_UV    = 6 'UpdatedVersion
Const COL_TL    = 7 'TargetLanguage
Const COL_TLV   = 8 ' Validate
Const COL_UC    = 9 'UpgradeCode
Const COL_UCV   = 10 ' Validate
Const COL_MST   = 11 'Transform
Const COL_ROW   = 12 'ExcelRow

Dim sErr_Syntax
sErr_Syntax    = vbCrLf & _
              "Usage:  " & vbTab & SOLUTIONNAME & ".vbs [/Option] ..." & vbCrLf & vbCrLf & _
              " /RepairCache  " & vbTab & "Tries to restore missing items in the local WI cache" & vbCrLf & _
              "   /SRestoreLocation=" & vbTab & "A list of fully qualified paths to folders with .msp files, separated by semicolons" & vbCrLf & _
              "          " & vbTab & vbTab & "<Folder01>;<\\Server02\Share02>;..." & vbCrLf & vbCrLf & _
              " /ReconcileCache " & "Unregisters missing patches in the cache to unblock broken WI configurations" & vbCrLf & vbCrLf& _
              " /ApplyPatch  " & vbTab & "Apply patches from current folder and SUpdateLocation" & vbCrLf & _
              "   /SUpdateLocation=" & vbTab & "A list of fully qualified paths to folders with .msi and/or .msp files, separated by semicolons" & vbCrLf & _
              "          " & vbTab & vbTab & "<Folder01>;<\\Server02\Share02>;..." & vbCrLf & _
              "   /ExcludeCache" & vbTab & "Will not apply any patches from %windir%\installer folder" & vbCrLf & _
              "   /IncludeOctCache" & vbTab & "Includes OCT patches from %windir%\installer folder into patch detection" & vbCrLf & vbCrLf & _
              " /RemovePatch= " & vbTab & "Uninstall specified list of patches, separated by semicolons" & vbCrLf & _
              "        " & vbTab & "Accepts KBxxxxxx;{PatchCode};<FullPath>;SUPERSEDED" & vbCrLf & vbCrLf & _
              " /CleanCache  " & vbTab & "Removes unreferenced (orphaned) patch files from the local WI cache" & vbCrLf & vbCrLf & _
              " /CabExtract=<Patch>" & vbTab & "Extracts the patch embedded .CAB file to the %temp% folder" & vbCrLf & vbCrLf & _
              " /ViewPatch=<Patch>" & vbTab & "Display the patch contents in Excel" & vbCrLf & vbCrLf & _
              " /DetectOnly" & vbTab & "Create a log file but do not execute any actions" & vbCrLf & vbCrLf & _
              " /q     " & vbTab & "Suppresses the automatic display of the log file" & vbCrLf & vbCrLf & vbCrLf & _
              " /register  " & vbTab & "Registers OPUtil context menu extensions for .msp files" & vbCrLf & vbCrLf & vbCrLf & _
              " /unregister" & vbTab & "UnRegisters OPUtil context menu extensions for .msp files" & vbCrLf & vbCrLf & vbCrLf & _
              "By default 'RepairCache' and 'ApplyPatch' are enabled." & vbCrLf & _
              "To disable use /[Option]=False." & vbCrLf & vbCrLf & _
              "Examples" & vbCrLf & "========" & vbCrLf & _
              "Default 'RepairCache' & 'ApplyPatch' from current directory:" & vbCrLf & _
              " cscript.exe " & SOLUTIONNAME & ".vbs" & vbCrLf & vbCrLf & _
              "Repair and reconcile a broken Windows Installer Cache:" & vbCrLf & _
              " cscript.exe " & SOLUTIONNAME & ".vbs /ReconcileCache /RepairCache /SRestoreLocation=<\\Location1\ShareName>;<\\Location2\ShareName>" & vbCrLf & vbCrLf & _
              "Create a log for applicability of specific patches:" & vbCrLf & _
              " cscript.exe " & SOLUTIONNAME & ".vbs /ApplyPatch /SUpdateLocation=<\\Location1\ShareName>;<\\Location2\ShareName> /ExcludeCache /DetectOnly" & vbCrLf & vbCrLf & _
              "Install applicable patches (including local Windows Installer cache):" & vbCrLf & _
              " cscript.exe " & SOLUTIONNAME & ".vbs /ApplyPatch /SUpdateLocation=<\\Location1\ShareName>;<\\Location2\ShareName>" & vbCrLf & vbCrLf & _
              "UnInstall patch(es):" & vbCrLf & _
              " cscript.exe " & SOLUTIONNAME & ".vbs /RemovePatch=KB123456;KB654321" & vbCrLf & _
              " cscript.exe " & SOLUTIONNAME & ".vbs /RemovePatch={PatchCode}" & vbCrLf & _
              " cscript.exe " & SOLUTIONNAME & ".vbs /RemovePatch=Superseded" & vbCrLf 

Dim oMsi, oFso, oReg, oWShell, oShellApp, oWmiLocal, XmlDoc, oFile, XlApp
Dim LogStream, ReadStream, LogProd

Dim fCScript, fx64, fCleanAggressive, fRebootRequired, fSumInit, fUpdatesCollected, fViewPatch, fCabExtract
Dim fShowLog, fForceRemovePatch, fContextMenu, fDeepScan, fDynSUpdateDiscovered, fMsiProvidedAsFile
Dim fNeedGenericSql, fRemovePatchQnD, fXl

Dim sLogFile, sLogSummary, sAppData, sTmp, sTemp, sWinDir, sWICacheDir, sScriptDir, sTimeStamp, sLogNoRef
Dim Location, Key, vWI, vWIMajorMinor, sProductVersionNew, sProductVersionReal, sMspFile, sApplyPatch
Dim sOSinfo, sOSVersion, sComputerName, sExternalMsi, sViewScope

Dim iIndex, iVersionNt

Dim arrUpdateLocations, arrRestoreLocations, arrSUpdatesAll, arrTmpLog, arrSchema
Dim dicSUpdatesAll, dicFamily, dicSummary, dicRepair, dicMspNoSeq, dicMspNoBase, dicMspMinor, dicMspSmall
Dim dicMspObsoleted, dicMspSequence, dicDynCultFolders, dicSqlCreateTbl, dicProdMst, dicFeatureStates

'-------------------------------------------------------------------------------
' Main
'-------------------------------------------------------------------------------
On Error Resume Next

'Initialize objects and defaults
Initialize

'Parse the command line
ParseCmdLine

'Validate ShowLog setting
If (fApplyPatch AND NOT fViewPatch) OR fCleanCache OR fForceRemovePatch OR fReconcileCache OR fRemovePatch OR fRepairCache Then fShowLog = True

sTmp = ""
Log  "Current Settings:"
Log Space(30) & "/RepairCache " & fRepairCache & ", " & "/SRestoreLocation=" & sRestoreLocation
Log Space(30) & "/ReconcileCache " & fReconcileCache
Log Space(30) & "/ApplyPatch " & fApplyPatch & ", /ExcludeCache " & fExcludeCache & ", /IncludeOctCache " & fIncludeOctCache & ", /SUpdateLocation=" & sUpdateLocation
Log Space(30) & "/RemovePatch " & fRemovePatch & ", Patches=" & sMspRemoveFilter
Log Space(30) & "/CleanCache " & fCleanCache
Log Space(30) & "/CabExtract " & fCabExtract
Log Space(30) & "/ViewPatch " & fViewPatch
Log Space(30) & "/DetectOnly " & fDetectOnly
Log Space(30) & "/Q " & bQuiet & vbCrLf
Log Space(30) & "For more details on available commands run " & chr(34) & "cscript OPUtil.vbs /?" & chr(34) & vbCrLf

If Not Err = 0 Then
    Log "Error: Could not determine script parameters. Aborting"
    Log vbCrLf & "End of script: " & Now
    LogStream.Close
    wscript.Quit 1
End If

'Log if DetectOnly
If fDetectOnly Then
    sTmp = "DetectOnly mode. No changes will be done to the system!"
    Log String(Len(sTmp), "=") & vbCrLf & sTmp & vbCrLf & String(Len(sTmp), "=") & vbCrLf
    sTmp = ""
End If 'fDetectOnly

'Set the bookmark for the summary
Log "[SUMMARY]"

'Add marker to indicate the start of debug logging section
sTmp = "Debug Logging Section"
Log vbCrLf & String(Len(sTmp), "=") & vbCrLf & sTmp & vbCrLf & String(Len(sTmp), "=")

'Ensure correct value for SRestoreLocation
arrRestoreLocations = EnsureLocation(sRestoreLocation & ";" & sUpdateLocation)

'Ensure correct value for SUpdateLocation
If (NOT fApplyPatch) AND (NOT fRemovePatch) AND sUpdateLocation="" _
 Then arrUpdateLocations = Split(sScriptDir, ";") _
 Else arrUpdateLocations = EnsureLocation(sUpdateLocation)

'"RepairCache" .msi/.msp resiliency'
'-------------------------------
If fRepairCache Then
    'Build the 'Repair' references
    InitRepairDic
    'Check and try to restore missing files if needed
    RepairCache
End If 'fRepairCache

'"ReconcileCache" patch reconcile
'--------------------------------
If fReconcileCache Then MspReconcile

'"RemovePatch" Uninstall patches
'-------------------------------
If fRemovePatch Then 
    Dim arrMspRemove
    Dim Msp, Item
    
    sTmp = "Running MspRemove with filter: '" & sMspRemoveFilter & "'"
    Log vbCrLf & vbCrLf & sTmp & vbCrLf & String(Len(sTmp), "-")
    If fCscript Then wscript.echo "Checking for removable patches"
    
    arrMspRemove = Split(sMspRemoveFilter, ";")
    For Each Item In arrMspRemove
        Msp = Item
        'Check if it's a reference to a .msp file
        If Len(Msp)>4 Then
            If LCase(Right(Msp, 4))=".msp" Then
                If oFso.FileExists(Msp) Then
                    sMspRemoveFilter = sMspRemoveFilter & Left(oMsi.SummaryInformation(Msp).Property(PID_REVNUMBER), 38)
                End If
            End If
        End If
        If NOT Left(Msp, 1) = "{" Then
            If NOT fUpdatesCollected Then CollectSUpdates
            For iIndex = 0 To UBound(arrSUpdatesAll)
                If arrSUpdatesAll(iIndex, COL_KB) = Replace(Msp, "KB", "") Then sMspRemoveFilter=sMspRemoveFilter & ";" & arrSUpdatesAll(iIndex, COL_PATCHCODE)
                If UCase(arrSUpdatesAll(iIndex, COL_RELEASE)) = UCase(Msp) Then sMspRemoveFilter=sMspRemoveFilter & ";" & arrSUpdatesAll(iIndex, COL_PATCHCODE)
            Next 'iIndex
        End If
    Next 'Item
    If InStr(sMspRemoveFilter, "{") > 0 Then MspRemove sMspRemoveFilter, sMspProductFilter
    If InStr(UCase(sMspRemoveFilter), "SUPERSEDED") > 0 Then MspRemove "Superseded", sMspProductFilter
End If

'"ApplyPatch" apply patches
'--------------------------
If fApplyPatch AND NOT fViewPatch Then ApplyPatches

'"CleanCache" orphaned .msp cleanup
'----------------------------------
If fCleanCache Then WICleanOrphans

'"ViewPatch"
'-----------
    If fViewPatch Then ViewPatch sMspFile

'"CabExtract"
'-------------
    If fCabExtract Then sTmp = CabExtract(sMspFile)

'Check reboot requirement
    If fRebootRequired Then
        Log vbCrLf & "Note: " & ERR_REBOOT
        If fCScript Then wscript.echo ERR_REBOOT
    End If 'fRebootRequired

'Close the temp log
    sTmp = "End of script: " & Now
    Log vbCrLf & vbCrLf & String(Len(sTmp), "=") & vbCrLf & sTmp
    LogStream.Close

'Create the final log including the summary section

'Log Notes & Errors
For Each Key in dicSummary.Keys
    sTmp = ""
    If NOT Left(Key, 1)="{" Then sLogSummary = sLogSummary & vbCrLf & Key & dicSummary.Item(Key)
Next 'Key

'By Patch Summary 
fSumInit = False
If fApplyPatch AND (NOT fViewPatch OR fDeepScan) Then
    If IsArray(arrSUpdatesAll) Then
        For iIndex = 0 To UBound(arrSUpdatesAll)
            If NOT ((Left(arrSUpdatesAll(iIndex, COL_FILENAME), Len(sWICacheDir)) = sWICacheDir) AND fExcludeCache) AND _
               ((NOT Left(arrSUpdatesAll(iIndex, COL_FILENAME), Len(sWICacheDir)) = sWICacheDir) OR _
               (Len(arrSUpdatesAll(iIndex, COL_APPLICABLECNT)) > 0) OR _
               (Len(arrSUpdatesAll(iIndex, COL_NOQALBASELINECNT)) > 0)) Then
                'Found patch to be logged
                'Check if Heading needs to be added
                If NOT fSumInit Then
                    fSumInit = True
                    sLogSummary = sLogSummary & vbCrLf & vbCrLf & "Summary By Patch" & vbCrLf & "================"
                    sLogNoRef = ""
                End If
                'Add patch & product details
                If (arrSUpdatesAll(iIndex, COL_REFCNT)= 0) Then
                    sLogNoRef = sLogNoRef & vbCrLf & " - KB " & arrSUpdatesAll(iIndex, COL_KB) & _
                        ", " & arrSUpdatesAll(iIndex, COL_PATCHCODE) & ", " & arrSUpdatesAll(iIndex, COL_PACKAGE) & ", " & arrSUpdatesAll(iIndex, COL_FILENAME)
                Else
                    sLogSummary = sLogSummary & vbCrLf & vbCrLf & "KB " & arrSUpdatesAll(iIndex, COL_KB) & _
                        ", " & arrSUpdatesAll(iIndex, COL_PATCHCODE) & ", " & arrSUpdatesAll(iIndex, COL_PACKAGE) & ", " & arrSUpdatesAll(iIndex, COL_FILENAME)
                    'Applied details
                    sTmp = vbTab & "Applied: "
                    If InStr(arrSUpdatesAll(iIndex, COL_APPLIEDCNT), ";")>0 Then
                        ReDim arrTmpLog(-1)
                        If Right(arrSUpdatesAll(iIndex, COL_APPLIEDCNT), 1)=";" Then _
                            arrSUpdatesAll(iIndex, COL_APPLIEDCNT) = Left(arrSUpdatesAll(iIndex, COL_APPLIEDCNT), Len(arrSUpdatesAll(iIndex, COL_APPLIEDCNT))-1)
                        arrTmpLog = Split(arrSUpdatesAll(iIndex, COL_APPLIEDCNT), ";")
                        sTmp = sTmp & "Patch is installed to " & UBound(arrTmpLog) + 1 & " product(s)"
                        For Each LogProd in arrTmpLog
                            sTmp = sTmp & vbCrLf & vbTab & vbTab & LogProd & " - "
                            sTmp = sTmp & oMsi.ProductInfo(LogProd, "ProductName")
                        Next 'LogPatch
                    
                        sLogSummary = sLogSummary & vbCrLf & sTmp
                    Else
                        sTmp = sTmp & "No"
                    End If
                
                    'Superseded details
                    sTmp = vbTab & "Superseded: "
                    If InStr(arrSUpdatesAll(iIndex, COL_SUPERSEDEDCNT), ";")>0 Then
                        ReDim arrTmpLog(-1)
                        If Right(arrSUpdatesAll(iIndex, COL_SUPERSEDEDCNT), 1)=";" Then _
                            arrSUpdatesAll(iIndex, COL_SUPERSEDEDCNT) = Left(arrSUpdatesAll(iIndex, COL_SUPERSEDEDCNT), Len(arrSUpdatesAll(iIndex, COL_SUPERSEDEDCNT))-1)
                        arrTmpLog = Split(arrSUpdatesAll(iIndex, COL_SUPERSEDEDCNT), ";")
                        sTmp = sTmp & "Patch is superseded for " & UBound(arrTmpLog) + 1 & " product(s)"
                        For Each LogProd in arrTmpLog
                            sTmp = sTmp & vbCrLf & vbTab & vbTab & LogProd & " - "
                            sTmp = sTmp & oMsi.ProductInfo(LogProd, "ProductName")
                        Next 'LogPatch
                        sLogSummary = sLogSummary & vbCrLf & sTmp
                    Else
                        sTmp = sTmp & "No"
                    End If
                
                    'Applicable details
                    sTmp = vbTab & "Applicable: "
                    If InStr(arrSUpdatesAll(iIndex, COL_APPLICABLECNT), ";") > 0 Then
                        ReDim arrTmpLog(-1)
                        If Right(arrSUpdatesAll(iIndex, COL_APPLICABLECNT), 1)=";" Then _
                            arrSUpdatesAll(iIndex, COL_APPLICABLECNT) = Left(arrSUpdatesAll(iIndex, COL_APPLICABLECNT), Len(arrSUpdatesAll(iIndex, COL_APPLICABLECNT))-1)
                        arrTmpLog = Split(arrSUpdatesAll(iIndex, COL_APPLICABLECNT), ";")
                        sTmp = sTmp & "Patch is applicable to " & UBound(arrTmpLog) + 1 & " product(s)"
                        For Each LogProd in arrTmpLog
                            sTmp = sTmp & vbCrLf & vbTab & vbTab & LogProd & " - "
                            sTmp = sTmp & oMsi.ProductInfo(LogProd, "ProductName")
                        Next 'LogPatch
                        sLogSummary = sLogSummary & vbCrLf & sTmp
                    Else
                        sTmp = sTmp & "No"
                    End If
                
                    'Applicable but no valid baseline details
                    sTmp = vbTab & "Can't apply: "
                    If InStr(arrSUpdatesAll(iIndex, COL_NOQALBASELINECNT), ";")>0 Then
                        ReDim arrTmpLog(-1)
                        If Right(arrSUpdatesAll(iIndex, COL_NOQALBASELINECNT), 1)=";" Then _
                            arrSUpdatesAll(iIndex, COL_NOQALBASELINECNT) = Left(arrSUpdatesAll(iIndex, COL_NOQALBASELINECNT), Len(arrSUpdatesAll(iIndex, COL_NOQALBASELINECNT))-1)
                        arrTmpLog = Split(arrSUpdatesAll(iIndex, COL_NOQALBASELINECNT), ";")
                        sTmp = sTmp & "Patch is applicable to " & UBound(arrTmpLog) + 1 & " product(s) but the product(s) do(es) not meet the required SP level "
                        For Each LogProd in arrTmpLog
                            sTmp = sTmp & vbCrLf & vbTab & vbTab & LogProd & " - " & oMsi.ProductInfo(LogProd, "ProductName") & _
                                   vbCrLf & vbTab & vbTab & "Patch baseline(s): " & arrSUpdatesAll(iIndex, COL_PATCHBASELINES) & ". Installed baseline: " & oMsi.ProductInfo(LogProd, "VersionString") & vbCrLf
                        Next 'LogPatch
                        sLogSummary = sLogSummary & vbCrLf & sTmp
                    Else
                        sTmp = sTmp & "No"
                    End If
                
                End If 
            End If
        Next 'iIndex
    End If 'IsArray
    If NOT sLogNoRef="" Then sLogSummary = sLogSummary & vbCrLf & vbCrLf & "Patch(es) that don't target any installed applications:" & sLogNoRef
End If 'fApplyPatch

'By Product Summary
If fSumInit Then sLogSummary = sLogSummary & vbCrLf & vbCrLf & vbCrLf & "Summary By Product" & vbCrLf & "==================" & vbCrLf

For Each Key in dicSummary.Keys
    sTmp = ""
    If Left(Key, 1)="{" Then 
        sTmp = Key & " - " & oMsi.ProductInfo(Key, "ProductName") 
        sLogSummary = sLogSummary & vbCrLf & sTmp & dicSummary.Item(Key)
    End If
Next 'Key
    Err.Clear

If sLogSummary = "===============" & vbCrLf & "Summary Section" & vbCrLf & "===============" & vbCrLf Then sLogSummary = sLogSummary & vbCrLf & " All appears to be well." & vbCrLf & vbCrLf & "For detailed logging see the 'Debug' section below." & vbCrLf
Set ReadStream= oFso.OpenTextFile(sLogFile, FOR_READING, False, TRISTATE_USEDEFAULT)
Set LogStream = oFso.CreateTextFile(sTemp & SOLUTIONNAME & ".log", True, True)
Do While Not ReadStream.AtEndOfStream
    sTmp = ReadStream.ReadLine
    If NOT InStr(sTmp, "[SUMMARY]")>0 Then
        LogStream.WriteLine sTmp
    Else
        LogStream.Write sLogSummary & vbCrLf
        If fRebootRequired Then LogStream.Write vbCrLf & ERR_REBOOT & vbCrLf & vbCrLf
        If (NOT ReadStream.AtEndOfStream AND NOT fNoDebugLog) Then LogStream.Write ReadStream.ReadAll
        Exit Do
    End If
Loop
ReadStream.Close
LogStream.Close
oFso.DeleteFile sLogFile

'copy log if needed
If oFso.FolderExists(sPathOutputFolder) Then
    Dim sLocalLog, sCustomFolderLog
    If NOT Right(sPathOutputFolder, 1) = "\" Then sPathOutputFolder = sPathOutputFolder & "\"
    sLocalLog = sTemp & SOLUTIONNAME & ".log"
    sCustomFolderLog = sPathOutputFolder & oWShell.ExpandEnvironmentStrings("%COMPUTERNAME%") & "_" & SOLUTIONNAME & ".log"
    oFso.CopyFile sLocalLog, sCustomFolderLog, True
End If

'Show completion notice
If fCscript AND NOT fViewPatch Then wscript.echo "Script execution complete."

'Show the log
If (NOT bQuiet) AND fShowLog Then oWShell.Run chr(34) & sTemp & SOLUTIONNAME & ".log" & chr(34)

'END 
'====
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'   Initialize
'
'   Initialize Objects and defaults
'-------------------------------------------------------------------------------
Sub Initialize()
    Dim Item, ComputerItem, Process, Processes, DateTime
    Dim iInstanceCnt

    On Error Resume Next

    fCScript = False
    fx64 = False
    fViewPatch = False
    fCabExtract = False
    fShowLog = False
    sMspFile = ""
    sApplyPatch = ""
    fUpdatesCollected = False
    fCleanAggressive = False
    fForceRemovePatch = False
    fRemovePatchQnD = False
    fRebootRequired = False
    fContextMenu = False
    fDeepScan = False
    fDynSUpdateDiscovered = False
    fNeedGenericSql = True
    fMsiProvidedAsFile = False

    Set dicSqlCreateTbl = CreateObject("Scripting.Dictionary")
    Set dicProdMst = CreateObject("Scripting.Dictionary")
    Set dicFeatureStates = CreateObject("Scripting.Dictionary")

    sLogSummary = "===============" & vbCrLf & "Summary Section" & vbCrLf & "===============" & vbCrLf

    Set oMsi = CreateObject("WindowsInstaller.Installer")
    Set oFso = CreateObject("Scripting.FileSystemObject")
    Set oWShell = CreateObject("WScript.Shell")
    Set oShellApp = CreateObject("Shell.Application")
    Set oReg = GetObject("winmgmts:\\.\root\default:StdRegProv")
    Set oWmiLocal   = GetObject("winmgmts:\\.\root\cimv2")
    Set DateTime = CreateObject("WbemScripting.SWbemDateTime")
    Set XmlDoc = CreateObject("Microsoft.XMLDOM")

    'Ensure there's only a single instance running of this script
    iInstanceCnt = 0
    Set Processes = oWmiLocal.ExecQuery("Select * From Win32_Process")
    For Each Process in Processes
        If LCase(Mid(Process.Name, 2, 6))="script" Then 
            If InStr(LCase(Process.CommandLine), "oputil")>0 Then iInstanceCnt=iInstanceCnt+1
        End If
    Next 'Process
    If iInstanceCnt > 1 Then
        If fCScript Then wscript.echo "Error: Another instance of this script is already running."
        wscript.quit
    End If

    'Obtain the current timestamp
    DateTime.SetVarDate Now, True
    sTimeStamp = Left(DateTime.Value, 14)

    'Are we running on Cscript?
    fCScript = (LCase(Mid(Wscript.FullName, Len(Wscript.Path)+2, 1)) = "c")

    'Get environment path info
    sAppData            = oWShell.ExpandEnvironmentStrings("%appdata%") & "\"
    sTemp               = oWShell.ExpandEnvironmentStrings("%temp%") & "\"
    sWinDir             = oWShell.ExpandEnvironmentStrings("%windir%") & "\"
    sWICacheDir         = sWinDir & "Installer\"
    sScriptDir          = wscript.ScriptFullName
    sScriptDir          = Left(sScriptDir, InStrRev(sScriptDir, "\"))

    'Init default for the resiliency .msi & .msp location
    sRestoreLocation = sRestoreLocation & ";" & sScriptDir & ";" & sWICacheDir
    If Left(sRestoreLocation, 1)=";" Then sRestoreLocation = Mid(sRestoreLocation, 2)

    'Create the logfile with initial data
    sLogFile = sTemp & "~" & SOLUTIONNAME & ".log"
    Set LogStream = oFso.CreateTextFile(sLogFile, True, True)
    Log "Microsoft Customer Support Services - " & SOLUTIONNAME & " V "& SCRIPTBUILD & " - " & Now & vbCrLf
    vWI = oMsi.Version
    vWIMajorMinor = Left(vWi, 3)
    sTmp = "Windows Installer Version:"
    Log sTmp & Space(30-Len(sTmp)) & vWI
    sTmp = "ComputerName:"
    Log  sTmp & Space(30-Len(sTmp)) & oWShell.ExpandEnvironmentStrings("%COMPUTERNAME%")

    If NOT sPathOutputFolder = "" Then 
        CreateFolderStructure sPathOutputFolder
        If NOT Right(sPathOutputFolder, 1) = "\" Then sPathOutputFolder = sPathOutputFolder & "\"
    End If
    If sPathOutputFolder = "" OR NOT oFso.FolderExists(sPathOutputFolder) Then sPathOutputFolder = sTemp


    'Initialize the 'Summary' log dictionary
    Set dicSummary = CreateObject("Scripting.Dictionary")

    If vWIMajorMinor = "4.5" AND vWI < "4.5.6001.22392" Then _
     LogSummary "Important Note:", "KB 972397 contains important updates for the installed version of Windows Installer! http://support.microsoft.com/kb/972397/EN-US/"
     'More recent version of WI: KB 2388997. KB search term "Msi.dll" AND "hotfix information" 

    'Detect if we're running on a 64 bit OS
    Set ComputerItem = oWmiLocal.ExecQuery("Select * from Win32_ComputerSystem")
    For Each Item In ComputerItem
        fx64 = Instr(Left(Item.SystemType, 3), "64") > 0
        sTmp = "OS Architecture:"
        Log sTmp & Space(30-Len(sTmp)) & Item.SystemType
    Next

End Sub 'Initialize
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'   ParseCmdLine
'
'   Command Line Parser
'-------------------------------------------------------------------------------
Sub ParseCmdLine()

    Dim sTmp, sArg
    Dim iCnt, iArgCnt
    Dim fIgnoreOnce, fArray, fValidCmdFound
    Dim arrArg

    On Error Resume Next

    fIgnoreOnce = False
    fArray = False
    iArgCnt = WScript.Arguments.Count

    If Not iArgCnt > 0 Then 
        'Use defaults
        Exit Sub
    End If

    Dim iActionCnt
    Dim fRepairCacheOrg, fReconcileCacheOrg, fApplyPatchOrg, fRemovePatchOrg, fCleanCacheOrg

    'Found command line argument(s) -> default to disabled modules
    iActionCnt = 0
    fValidCmdFound = False
    fRepairCacheOrg = fRepairCache  : fRepairCache = False
    fReconcileCacheOrg = fReconcileCache : fReconcileCache = False
    fApplyPatchOrg = fApplyPatch    : fApplyPatch = False
    fRemovePatchOrg = fRemovePatch  : fRemovePatch = False
    fCleanCacheOrg = fCleanCache    : fCleanCache = False

    For iCnt = 0 To (iArgCnt-1)

        sArg = ""
        sArg = UCase(WScript.Arguments(iCnt))
        If InStr(sArg, "=")>0 Then
            Set arrArg = Nothing
            arrArg=Split(sArg, "=", 2)
            sArg = arrArg(0)
            fArray = True
        End If
    
        Select Case sArg
    
        Case "/?"
            Wscript.Echo sErr_Syntax
            Wscript.Quit
   
    
        Case "/APPLYPATCH", "-APPLYPATCH", "APPLYPATCH"
            fValidCmdFound = True
            fShowLog = True
            If fArray Then
                If arrArg(1) = "FALSE" Then
                    fApplyPatch = False
                    fApplyPatchOrg = False
                    fValidCmdFound = False
                End If
                fApplyPatch = True
                iActionCnt = iActionCnt + 1
                sApplyPatch = arrArg(1)
                If oFso.FileExists(sApplyPatch) Then
                    Set oFile = oFso.GetFile(sApplyPatch)
                    If sUpdateLocation = "" Then sUpdateLocation = oFile.ParentFolder.Path Else sUpdateLocation = sUpdateLocation & ";" & oFile.ParentFolder.Path
                End If
                fArray = False
            Else
                fApplyPatch = True
                iActionCnt = iActionCnt + 1
            End If
    
        Case "/CABEXTRACT", "-CABEXTRACT", "CABEXTRACT"
            fValidCmdFound = True
            iActionCnt = iActionCnt + 1
            fCabExtract = True
            If fArray Then
                sMspFile = arrArg(1)
                fArray = False
            Else
                fIgnoreOnce = True
                If (iArgCnt-1) > iCnt Then _
                  sMspFile = WScript.Arguments(iCnt+1) Else
                  sMspFile = ""
            End If
    
        Case "/CLEANAGGRESSIVE", "-CLEANAGGRESSIVE", "CLEANAGGRESSIVE"
            fCleanAggressive = True
    
        Case "/CLEANCACHE", "-CLEANCACHE", "CLEANCACHE"
            fValidCmdFound = True
            fShowLog = True
            If fArray Then
                If arrArg(1) = "FALSE" Then
                    fCleanCache = False
                    fCleanCacheOrg = False
                End If
                If arrArg(1) = "TRUE" Then
                    fCleanCache = True
                    iActionCnt = iActionCnt + 1
                End If
                fArray = False
            Else
                fCleanCache = True
                iActionCnt = iActionCnt + 1
            End If
    
        Case "/CONTEXTMENU"
            fContextMenu = True
    
    '    Case "/DEEPSCAN"
    '        iActionCnt = iActionCnt + 1
    '        fDeepScan = True
    '        fDetectOnly = True
    '        fApplyPatch = True

        Case "/DETECTONLY", "-DETECTONLY", "DETECTONLY"
            If fArray Then
                If arrArg(1)="TRUE" OR arrArg(1)=1 Then fDetectOnly=True Else fDetectOnly=False
                fArray=False
            Else
                fDetectOnly = True
            End If 'fArray
    
        Case "/DISABLEREPAIR", "-DISABLEREPAIR", "DISABLEREPAIR"
            fRepairCache = False
    
        Case "/DISABLERERESTARTMANAGER", "-DISABLERERESTARTMANAGER", "DISABLERERESTARTMANAGER"
            fDisableRestartManager = False
    
        Case "/EXCLUDECACHE", "-EXCLUDECACHE", "EXCLUDECACHE"
            fExcludeCache = True
    
        'Warning: This is an undocumented and unsupported feature!
        Case "/FORCEREMOVEPATCH", "-FORCEREMOVEPATCH", "FORCEREMOVEPATCH"
            fForceRemovePatch = True
            fValidCmdFound = True
            fShowLog = True
            If fArray Then
                If arrArg(1) = "FALSE" Then 
                    fForceRemovePatch = False
                Else
                    fRemovePatch = True
                    sMspRemoveFilter = Replace(arrArg(1), ",", ";")
                    iActionCnt = iActionCnt + 1
                End If
                fArray = False
            End If

        'Warning: This is an undocumented and unsupported feature!
        Case "/REMOVEPATCHQND", "-REMOVEPATCHQND", "REMOVEPATCHQND"
            fRemovePatchQnD = True
            fValidCmdFound = True
            fShowLog = True
            If fArray Then
                If arrArg(1) = "FALSE" Then 
                    fRemovePatchQnD = False
                Else
                    fRemovePatch = True
                    sMspRemoveFilter = Replace(arrArg(1), ",", ";")
                    iActionCnt = iActionCnt + 1
                End If
                fArray = False
            End If
    
        Case "/INCLUDEOCTCACHE", "-INCLUDEOCTCACHE", "INCLUDEOCTCACHE"
            fIncludeOctCache = True
    
        Case "/R", "-R", "R", "/REGISTER", "-REGISTER", "REGISTER"
            fValidCmdFound = True
            iActionCnt = iActionCnt + 1
            RegisterShellExt
    
        Case "/RECONCILECACHE", "-RECONCILECACHE", "RECONCILECACHE"
            fValidCmdFound = True
            fShowLog = True
            If fArray Then
                If arrArg(1) = "FALSE" Then 
                    fReconcileCache = False
                    fReconcileCacheOrg = False
                End If
                If arrArg(1) = "TRUE" Then
                    fReconcileCache = True
                    iActionCnt = iActionCnt + 1
                End If
                fArray = False
            Else
                fReconcileCache = True
                iActionCnt = iActionCnt + 1
            End If
    
        Case "/REMOVEPATCH", "-REMOVEPATCH", "REMOVEPATCH"
            fValidCmdFound = True
            fShowLog = True
            If fArray Then
                If arrArg(1) = "FALSE" Then 
                    fRemovePatch = False
                    fRemovePatchOrg = False
                Else
                    fRemovePatch = True
                    sMspRemoveFilter = Replace(arrArg(1), ",", ";")
                    iActionCnt = iActionCnt + 1
                End If
                fArray = False
            Else
                fRemovePatch = True
                iActionCnt = iActionCnt + 1
            End If
    
        Case "/REPAIRCACHE", "-REPAIRCACHE", "REPAIRCACHE"
            fValidCmdFound = True
            fShowLog = True
            If fArray Then
                If arrArg(1) = "FALSE" Then 
                    fRepairCache = False
                    fRepairCacheOrg = False
                End If
                If arrArg(1) = "TRUE" Then
                    fRepairCache = True
                    iActionCnt = iActionCnt + 1
                End If
                fArray = False
            Else
                fRepairCache = True
                iActionCnt = iActionCnt + 1
            End If
    
        Case "/SRESTORELOCATION", "-SRESTORELOCATION", "SRESTORELOCATION" 'SRestoreLocation
            If fArray Then
                sRestoreLocation = arrArg(1)
                fArray = False
            Else
                fIgnoreOnce = True
                sRestoreLocation = WScript.Arguments(iCnt+1)
            End If
    
        Case "/SUPDATELOCATION", "-SUPDATELOCATION", "SUPDATELOCATION", "/SUPDATESLOCATION", "-SUPDATESLOCATION", "SUPDATESLOCATION" 'SUpdateLocation
            If fArray Then
                sUpdateLocation = arrArg(1)
                fArray = False
            Else
                fIgnoreOnce = True
                sUpdateLocation = WScript.Arguments(iCnt+1)
            End If

        Case "/U", "-U", "U", "/UNREGISTER", "UNREGISTER", "UNREGISTER"
            fValidCmdFound = True
            iActionCnt = iActionCnt + 1
            UnRegisterShellExt
    
    
        Case "/VIEW", "/VIEWPATCH"
            iActionCnt = iActionCnt + 1
            fValidCmdFound = True
            fViewPatch = True
            fDetectOnly = True
            fApplyPatch = True
            'fDeepScan = True
            If fArray Then
                sMspFile = arrArg(1)
                If oFso.FileExists(sMspFile) Then
                    Set oFile = oFso.GetFile(sMspFile)
                    If sUpdateLocation = "" Then sUpdateLocation = oFile.ParentFolder.Path Else sUpdateLocation = sUpdateLocation & ";" & oFile.ParentFolder.Path
                End If
                fArray = False
            Else
                fIgnoreOnce = True
                If (iArgCnt-1) > iCnt Then _
                  sMspFile = WScript.Arguments(iCnt+1) Else _
                  sMspFile = ""
            End If
    
        Case "/SCOPE"
        ' only valid in combination with /ViewPatch
        ' specifies the scope for product specific PatchTables
            If fArray Then
                sViewScope = UCase(arrArg(1))
                sViewScope = Replace(sViewScope, ";", ",")
            End If
    
        Case "/Q"
            bQuiet = True
    
        Case Else
            If NOT fIgnoreOnce Then
                sTmp = ""
                sTmp = vbCrLf & "Warning: Invalid command line switch '" & WScript.Arguments(iCnt) & "' will be ignored." & vbCrLf
                If NOT bQuiet Then wscript.echo sTmp
                Log sTmp
                fIgnoreOnce = NOT fIgnoreOnce
            End If
    
        End Select

    Next 'iCnt

    'Ensure we had a valid Cmd
    If NOT fValidCmdFound OR iActionCnt = 0 Then
        'Restore defaults
        fShowLog        = True
        fRepairCache    = fRepairCacheOrg
        fReconcileCache = fReconcileCacheOrg
        fApplyPatch     = fApplyPatchOrg
        fRemovePatch    = fRemovePatchOrg
        fCleanCache     = fCleanCacheOrg
    End If

End Sub 'ParseCmdLine
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'   InitRepairDic
'
'   Build a dictionary array for a reference list of available .msi and .msp packages
'   Key for .msi is <ProductCode>_<PackageCode> (required to support AIP installs)
'   Key for .msp is <PatchCode>
'-------------------------------------------------------------------------------
Sub InitRepairDic()

    Dim File, Folder, MspDb, Record, SumInfo
    Dim sProductCode, sPackageCode, sKey
    Dim qView

    On Error Resume Next

    If fCscript Then wscript.echo "Collecting resiliency data"
    Set dicRepair = CreateObject("Scripting.Dictionary")

    For Each Folder in arrRestoreLocations
        If fCscript Then wscript.echo vbTab & "Collect files from " & Folder
        For Each File in oFso.GetFolder(Folder).Files
            Select Case LCase(Right(File.Name, 4))
            Case ".msp"
                sKey = ""
                sKey = oMsi.SummaryInformation(File.Path, MSIOPENDATABASEMODE_READONLY).Property(PID_REVNUMBER)
                If Len(sKey)>LEN_GUID Then sKey=Left(sKey, LEN_GUID)
                If Not dicRepair.Exists(sKey) Then dicRepair.Add sKey, File.Path
            Case ".msi"
                sKey = GetMsiProductCode(File.Path) & "_" & GetMsiPackageCode(File.Path)
                If Not dicRepair.Exists(sKey) Then dicRepair.Add sKey, File.Path
            Case Else
                'Do Nothing
            End Select
        Next 'File
    Next 'Folder

End Sub 'InitRepairDic
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'   RepairCache
'-------------------------------------------------------------------------------
Sub RepairCache()

    Dim File, Product, Prod, PatchList, Patch, Source, MsiSources, MspSources
    Dim sLocalMsi, sLocalMsp, sKey, sSourceKey, sRepair, sFile, sFileName, sPackage
    Dim sRegPackageCode, sMsiPackageCode, sGlobalPatchesKey, sClassesPatchesKey
    Dim fTrySource, fRepaired, fMsiOK, fMsiRename, fReLoop
    Dim dicMspError, dicMspChecked, dicMspUnreg, arrKeys
    Dim iReLoop

    On Error Resume Next

    Set dicMspError = CreateObject("Scripting.Dictionary")
    Set dicMspChecked = CreateObject("Scripting.Dictionary")
    Set dicMspUnreg = CreateObject("Scripting.Dictionary")

    sTmp = "Running RepairCache"
    Log vbCrLf & vbCrLf & sTmp & vbCrLf & String(Len(sTmp), "-")
    If fCscript Then wscript.echo "Scanning Windows Installer cache"

    For Each Product in oMsi.Products
        Log vbCrLf & "Product: " & Product & " - " & oMsi.ProductInfo(Product, "ProductName")
        If fCscript Then wscript.echo vbTab & "Scan " & Product & " - " & oMsi.ProductInfo(Product, "ProductName")
    
        'Check local .msi package
        sLocalMsi = "" : sRegPackageCode = "" : sFileName = "" : sRepair = ""
        fTrySource = False : fRepaired = False : fMsiOK = False : fMsiRename = False
        Err.Clear
        sLocalMsi = oMsi.ProductInfo(Product, "LocalPackage")
        sRegPackageCode = oMsi.ProductInfo(Product, "PackageCode")
        If Err = 0 Then
            If oFso.FileExists(sLocalMsi) Then
                sMsiPackageCode = GetMsiPackageCode(sLocalMsi)
                If sRegPackageCode = sMsiPackageCode Then 
                    fMsiOK = True
                    Log vbTab & "Success: Local .msi package " & sLocalMsi & " is available and valid."
                Else
                    'PackageCode mismatch! Windows Installer will not accept the local copy in the WI cache as valid.
                    fMsiRename = True
                    sTmp = vbTab & "Error: Local .msi package " & sLocalMsi & " is available but invalid. Registered PackageCode '" & sRegPackageCode & "' does not match cached files PackageCode '" & sMsiPackageCode & "'"
                    LogSummary Product, sTmp
                    If fCscript Then wscript.echo vbTab & vbTab & "Error: Local .msi package " & sLocalMsi & " is available but invalid."
                End If
            End If 'oFso.FileExists
        Else
            Err.Clear
        End If
    
        If NOT fMsiOK Then
            If sLocalMsi = "" Then 
                sTmp = vbTab & "Error: No local .msi package registered. Cannot restore."
                Log sTmp
            Else
                'Try to restore from available resources
                sRepair = "Error: Local .msi package missing. Attempt failed to restore "
                If fMsiRename Then sRepair = "Note: No matching .msi package available to replace the mismatched file "
                sKey = ""
                sKey = Product & "_" & sRegPackageCode
                If dicRepair.Exists(sKey) Then 
                    If fMsiRename Then
                        If NOT fDetectOnly Then
                            sFileName = ""
                            Set File = oFso.GetFile(sLocalMsi)
                            sFileName = File.Name
                            File.Name = "Renamed_" & File.Name
                            oFso.CopyFile dicRepair.Item(sKey), sLocalMsi
                        End If 'fDetectOnly
                    Else
                        If NOT fDetectOnly Then oFso.CopyFile dicRepair.Item(sKey), sLocalMsi
                    End If 'fMsiRename
                    If oFso.FileExists(sLocalMsi) _
                      Then sRepair = "Restored: Successfully connected to 'RestoreLocation' (" & dicRepair.Item(sKey) & ") to restore local .msi package " _
                      Else fTrySource = True
                
                    'Handle 'DetectOnly' exception
                    If fDetectOnly Then
                        sRepair = "Note: Restore is possible from 'RestoreLocation' (" & dicRepair.Item(sKey) & ") to restore local .msi package "
                        fTrySource = False
                    End If 'fDetectOnly
                Else
                    fTrySource = True
                End If
            
                'Try to restore from resgistered sources
                If fTrySource Then
                    'Obtain a productsex handle
                    Set Prod = oMsi.ProductsEx(Product, "", MSIINSTALLCONTEXT_ALL)(0)
                    'Get the sources
                    Set MsiSources = Prod.Sources(1)
                    sPackage = ""
                    sPackage = Prod.SourceListInfo("PackageName")
                    For Each Source in MsiSources
                        Log "Debug: Trying to connect to resiliency source " & Source
                        If fCscript Then wscript.echo vbTab & vbTab & "Trying to connect to resiliency source " & Source
                        If fRepaired Then Exit For
                        sFile = ""
                        sFile = Source & sPackage
                        If oFso.FileExists(sFile) Then
                            sSourceKey = ""
                            sSourceKey = GetMsiProductCode(sFile) & "_" & GetMsiPackageCode(sFile)
                            If sKey = sSourceKey Then
                                If Not dicRepair.Exists(sSourceKey) Then dicRepair.Add sSourceKey, sFile
                                If NOT fDetectOnly AND fMsiRename Then 
                                    sFileName = ""
                                    Set File = oFso.GetFile(sLocalMsi)
                                    sFileName = File.Name
                                    File.Name = "Renamed_" & File.Name
                                End If
                                If NOT fDetectOnly Then oFso.CopyFile sFile, sLocalMsi
                                fRepaired = oFso.FileExists(sLocalMsi)
                                If fDetectOnly Then sRepair = "Note: Restore is possible from 'registered InstallSource' (" & sFile & ") to restore local .msi package "
                            End If 'sKey = sSourceKey
                        End If
                    Next 'Source
                    If fRepaired Then sRepair = "Restored: Successfully connected to 'registered InstallSource' (" & sFile & ") to restore local .msi package "
                End If 'fTrySource

                If NOT oFso.FileExists(sLocalMsi) AND fMsiRename Then
                    'Undo rename
                    File.Name = sFileName
                    sRepair = "Error: Attmpt failed to replace the mismatched file. Original cached file has been restored "
                End If
                Set File = Nothing

                'Log the result
                sTmp = vbTab & sRepair & sLocalMsi & " (PackageCode: " & sRegPackageCode & ")"
                Log sTmp
            End If
            LogSummary Product, sTmp
            If fCscript Then wscript.echo vbTab & vbTab & sTmp
        End If
    
        'Check local .msp packages
        For iReLoop = 0 To 1
            fReLoop = False
            Set PatchList = oMsi.PatchesEx(Product, USERSID_NULL, MSIINSTALLCONTEXT_MACHINE, MSIPATCHSTATE_APPLIED + MSIPATCHSTATE_SUPERSEDED + MSIPATCHSTATE_OBSOLETED)
            If Err = 0 Then
                For Each Patch in PatchList
                    Err.Clear
                    sLocalMsp = "" : sRepair = "" : fTrySource = False : fRepaired = False
                    sLocalMsp = LCase(Patch.PatchProperty("LocalPackage"))
                    If Not dicMspChecked.Exists(Patch.PatchCode) Then dicMspChecked.Add Patch.PatchCode, sLocalMsp
                    If Not Err = 0 Then
                        Err.Clear
                        'This happens if a patch is registered but the global patch registration has gone missing.
                        'To work around this a correction entry is created
                        sTmp = vbTab & "Error: Failed to obtain local patch package data for patch '" & Patch.PatchCode & "'. Fixing patch registration."
                        If fDetectOnly Then sTmp = vbTab & "Error: Failed to obtain local patch package data for patch '" & Patch.PatchCode & "'. Patch registration would be fixed."
                        Log sTmp
                        LogSummary Product, sTmp
                        If fCscript Then wscript.echo vbTab & vbTab & sTmp
                        If NOT fDetectOnly Then
                            FixMspGlobalReg Patch.PatchCode
                            fReLoop = True
                            sLocalMsp = LCase(Patch.PatchProperty("LocalPackage"))
                        End If
                    End If
                    If NOT sLocalMsp = "" Then
                        If oFso.FileExists(sLocalMsp) Then 
                            Log vbTab & "Success: Confirmed local patch package as '" & sLocalMsp & "'" & vbTab & "for patch '" & Patch.PatchCode & "' - '" & Patch.PatchProperty("DisplayName") & "'."
                        Else
                            'Try to restore from available resources
                            sRepair = "Error: Local .msp package missing. Attempt failed to restore '"
                            sKey = ""
                            sKey = Patch.PatchCode
                            If dicRepair.Exists(sKey) Then
                                If NOT fDetectOnly Then oFso.CopyFile dicRepair.Item(sKey), sLocalMsp
                                If oFso.FileExists(sLocalMsp) Then sRepair = "Restored: Successfully connected to 'RestoreLocation' (" & dicRepair.Item(sKey) & ") to restore local .msp package '" Else fTrySource = True
                                'Handle 'DetectOnly' exception
                                If fDetectOnly Then
                                    sRepair = "Note: Restore is possible from 'RestoreLocation' (" & dicRepair.Item(sKey) & ") to restore local .msp package "
                                    fTrySource = False
                                    If NOT dicMspError.Exists(Patch.PatchCode) Then dicMspError.Add Patch.PatchCode, sLocalMsp
                                End If 'fDetectOnly
                            Else
                                fTrySource = True
                            End If
                    
                            'Try to restore from resgistered sources
                            If fTrySource Then
                                'Get the sources
                                sPackage = Patch.SourceListInfo("PackageName")
                                Set MspSources = Patch.Sources(1)
                                For Each Source in MspSources
                                    If fRepaired Then Exit For
                                    sFile = Source & sPackage
                                    If oFso.FileExists(sFile) Then
                                        sSourceKey = ""
                                        sSourceKey = oMsi.SummaryInformation(sFile, MSIOPENDATABASEMODE_READONLY).Property(PID_REVNUMBER)
                                        If sKey = sSourceKey Then
                                            If NOT dicRepair.Exists(sSourceKey) Then dicRepair.Add sSourceKey, sFile
                                            If NOT fDetectOnly Then oFso.CopyFile sFile, sLocalMsp
                                            fRepaired = oFso.FileExists(sLocalMsp)
                                            If fDetectOnly Then 
                                                sRepair = "Note: Restore is possible from 'registered InstallSource' (" & sFile & ") to restore local .msp package "
                                                If NOT dicMspError.Exists(Patch.PatchCode) Then dicMspError.Add Patch.PatchCode, sLocalMsp
                                            End If
                                        End If 'sKey = sSourceKey
                                    End If
                                Next 'Source
                                If fRepaired Then
                                    sRepair = "Restored: Successfully connected to 'registered InstallSource' (" & sFile & ") to restore local .msp package "
                                Else
                                    If NOT dicMspError.Exists(Patch.PatchCode) Then dicMspError.Add Patch.PatchCode, sLocalMsp
                                End If
                            End If 'fTrySource
                    
                            'Log the result
                            sTmp = vbTab & sRepair & sLocalMsp & "' - '" & Patch.PatchCode & "' - '" & Patch.PatchProperty("DisplayName")
                            Log sTmp
                            LogSummary Product, sTmp
                        End If 'NOT oFso.FileExists
                    End If 'Not sLocalMsp = ""
                    If NOT fReLoop Then EnsurePatchMetadata Patch, USERSID_NULL
                Next 'Patch
            Else
                sTmp = vbTab & "Error: PatchesEx API failed with error " & err.number & " - " & err.Description
                Log sTmp
                LogSummary Product, sTmp & " (Module RepairCache)"
                If fCscript Then wscript.echo vbTab & vbTab & sTmp
            End If 'Err = 0
            If NOT fReLoop Then Exit For
        Next 'iReLoop
    
    Next 'Prod

    'In case that a global patch entry exists which is no longer linked to any product this is not covered
    'in the logic above and requires this special handler
    sGlobalPatchesKey = REG_GLOBALCONFIG & "S-1-5-18\Patches\"
    sClassesPatchesKey = "Installer\Patches\"
    If RegEnumKey(HKLM, sGlobalPatchesKey, arrKeys) Then
        For Each sKey in arrKeys
            Patch = GetExpandedGuid(sKey)
            If NOT dicMspChecked.Exists(Patch) Then
                'Only care if it's impacting known patches from the repair dictionary
                If dicRepair.Exists(sKey) Then
                    'Flag to reconcile the registration to allow a clean transaction
                    If Not dicMspUnreg.Exists(sKey) Then dicMspUnreg.Add sKey, sKey
                End If 'dicRepair.Exists
            End If
        Next 'sKey
    End If 'RegEnumKey sGlobalPatchesKey
    If RegEnumKey(HKCR, sClassesPatchesKey, arrKeys) Then
        For Each sKey in arrKeys
            Patch = GetExpandedGuid(sKey)
            If NOT dicMspChecked.Exists(Patch) Then
                'Only care if it's impacting known patches from the repair dictionary
                If dicRepair.Exists(sKey) Then
                    'Flag to reconcile the registration to allow a clean transaction
                    If Not dicMspUnreg.Exists(sKey) Then dicMspUnreg.Add sKey, sKey
                End If 'dicRepair.Exists
            End If
        Next 'sKey
    End If 'RegEnumKey sClassesPatchesKey
    If dicMspUnreg.Count > 0 Then
        For Each sKey in dicMspUnreg.Keys
            RegDeleteKey HKLM, sGlobalPatchesKey & sKey & "\"
            RegDeleteKey HKCR, sClassesPatchesKey & sKey & "\"
        Next 'sKey
    End If 'dicMspUnreg > 0

    Err.Clear
    Set PatchList = oMsi.PatchesEx("", USERSID_NULL, MSIINSTALLCONTEXT_MACHINE, MSIPATCHSTATE_APPLIED + MSIPATCHSTATE_SUPERSEDED + MSIPATCHSTATE_OBSOLETED)
    If Err = 0 Then
        For Each Patch in PatchList
            'Only care if it's not a patch with known issues
            If NOT dicMspChecked.Exists(Patch.PatchCode) Then
                Err.Clear
                sLocalMsp = "" : sRepair = "" : fTrySource = False : fRepaired = False
                sLocalMsp = LCase(Patch.PatchProperty("LocalPackage"))
                If Not Err = 0 Then
                    Err.Clear
                    'This happens if a patch is registered but the global patch registration has gone missing.
                    'To work around this a correction entry is created
                    sTmp = vbTab & "Error: Failed to obtain local patch package data for patch '" & Patch.PatchCode & "'. Fixing patch registration."
                    If fDetectOnly Then sTmp = vbTab & "Error: Failed to obtain local patch package data for patch '" & Patch.PatchCode & "'. Patch registration would be fixed."
                    Log sTmp
                    LogSummary "", sTmp
                    If fCscript Then wscript.echo vbTab & vbTab & sTmp
                    If NOT fDetectOnly Then
                        FixMspGlobalReg Patch.PatchCode
                        sLocalMsp = LCase(Patch.PatchProperty("LocalPackage"))
                    End If
                End If
                If NOT sLocalMsp = "" Then
                    If oFso.FileExists(sLocalMsp) Then 
                        Log vbTab & "Success: Confirmed local patch package as '" & sLocalMsp & "'" & vbTab & "for patch '" & Patch.PatchCode & "' - '" & Patch.PatchProperty("DisplayName") & "'."
                    Else
                        'Try to restore from available resources
                        sRepair = "Error: Local .msp package missing. Attempt failed to restore '"
                        sKey = ""
                        sKey = Patch.PatchCode
                        If dicRepair.Exists(sKey) Then
                            If NOT fDetectOnly Then oFso.CopyFile dicRepair.Item(sKey), sLocalMsp
                            If oFso.FileExists(sLocalMsp) Then sRepair = "Restored: Successfully connected to 'RestoreLocation' (" & dicRepair.Item(sKey) & ") to restore local .msp package '" Else fTrySource = True
                            'Handle 'DetectOnly' exception
                            If fDetectOnly Then
                                sRepair = "Note: Restore is possible from 'RestoreLocation' (" & dicRepair.Item(sKey) & ") to restore local .msp package "
                                fTrySource = False
                                If NOT dicMspError.Exists(Patch.PatchCode) Then dicMspError.Add Patch.PatchCode, sLocalMsp
                            End If 'fDetectOnly
                        Else
                            fTrySource = True
                        End If
                    
                        'Try to restore from resgistered sources
                        If fTrySource Then
                            'Get the sources
                            sPackage = Patch.SourceListInfo("PackageName")
                            Set MspSources = Patch.Sources(1)
                            For Each Source in MspSources
                                If fRepaired Then Exit For
                                sFile = Source & sPackage
                                If oFso.FileExists(sFile) Then
                                    sSourceKey = ""
                                    sSourceKey = oMsi.SummaryInformation(sFile, MSIOPENDATABASEMODE_READONLY).Property(PID_REVNUMBER)
                                    If sKey = sSourceKey Then
                                        If NOT dicRepair.Exists(sSourceKey) Then dicRepair.Add sSourceKey, sFile
                                        If NOT fDetectOnly Then oFso.CopyFile sFile, sLocalMsp
                                        fRepaired = oFso.FileExists(sLocalMsp)
                                        If fDetectOnly Then 
                                            sRepair = "Note: Restore is possible from 'registered InstallSource' (" & sFile & ") to restore local .msp package "
                                            If NOT dicMspError.Exists(Patch.PatchCode) Then dicMspError.Add Patch.PatchCode, sLocalMsp
                                        End If
                                    End If 'sKey = sSourceKey
                                End If
                            Next 'Source
                            If fRepaired Then
                                sRepair = "Restored: Successfully connected to 'registered InstallSource' (" & sFile & ") to restore local .msp package "
                            Else
                                If NOT dicMspError.Exists(Patch.PatchCode) Then dicMspError.Add Patch.PatchCode, sLocalMsp
                                'MspReconcile logic is not designed to handle this special case. 
                                'Unregister is called straight away if MspReconcile is scheduled to run
                                If NOT fDetectOnly AND fReconcileCache Then UnregisterPatch Patch
                            End If
                        End If 'fTrySource
                    
                        'Log the result
                        sTmp = vbTab & sRepair & sLocalMsp & "' - '" & Patch.PatchCode & "' - '" & Patch.PatchProperty("DisplayName")
                        Log sTmp
                        LogSummary "", sTmp
                    End If 'NOT oFso.FileExists
                End If 'Not sLocalMsp = ""
            End If 'NOT dicMspError.Exists
        Next 'Patch
    Else
        sTmp = vbTab & "Error: PatchesEx API failed with error " & err.number & " - " & err.Description
        Log sTmp
        LogSummary "", sTmp &" (Module RepairCache)"
        If fCscript Then wscript.echo vbTab & vbTab & sTmp
    End If 'Err = 0

End Sub 'RepairCache
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'   MspReconcile
'
'   Unregister .msp files that have gone missing from the 
'   '%windir%\installer' folder
'-------------------------------------------------------------------------------
Sub MspReconcile()
    Const MAX_ATTEMPT = 100

    Dim Prod, Product, Patch, PatchList, ProductsList
    Dim sLocalMsp, sLocalMsi
    Dim iCnt
    Dim fResume, fMspOk
    Dim dicMspUnregister

    On Error Resume Next

    Set dicMspUnregister = CreateObject("Scripting.Dictionary")

    sTmp = "Running Module - Msp Reconcile"
    Log vbCrLf & vbCrLf & sTmp & vbCrLf & String(Len(sTmp), "-")
    If fCscript Then wscript.echo "Scanning for broken patches"

    iCnt = 0

    'Main detection loop
    Set ProductsList = oMsi.ProductsEx("", "", MSIINSTALLCONTEXT_MACHINE)
    For Each Prod in ProductsList
        Product = Prod.ProductCode
        If IsOfficeProduct (Product) Then
            Log vbCrLf & "Product: " & Product & " - " & oMsi.ProductInfo(Product, "ProductName")
            If fCscript Then wscript.echo vbTab & "Scan " & Product & " - " & oMsi.ProductInfo(Product, "ProductName")
        
            If NOT fRepairCache Then
                'Check local .msi package
                sLocalMsi = ""
                sLocalMsi = oMsi.ProductInfo(Product, "LocalPackage")
                If oFso.FileExists(sLocalMsi) Then
                    Log vbTab & "Success: Local .msi package " & sLocalMsi & " is available."
                Else
                    If sLocalMsi = "" Then 
                        sTmp = vbTab & "Error: No local .msi package registered."
                        Log sTmp
                    Else
                        sTmp = vbTab & "Error: Local .msi package " & sLocalMsi & " is missing."
                        Log sTmp
                    End If
                    LogSummary Product, sTmp
                    If fCscript Then wscript.echo vbTab & vbTab & sTmp
                End If
            End If 'NOT fRepairCache

            'Get the list of patches for the product
            fResume = True
            fMspOk = True
            Do While fResume 
                Err.Clear
                fResume = False
                Set PatchList = oMsi.PatchesEx(Product, USERSID_NULL, MSIINSTALLCONTEXT_MACHINE, MSIPATCHSTATE_ALL)
                If Err = 0 Then
                    For Each Patch in PatchList
                        Err.Clear
                        sLocalMsp = "" : sLocalMsp = LCase(Patch.PatchProperty("LocalPackage"))
                        If Not Err = 0 Then
                            fMspOk = False
                            Err.Clear
                            If NOT dicMspUnregister.Exists(Patch.PatchCode) Then
                                sTmp = vbTab & "Error: Failed to obtain local patch package data for patch '" & Patch.PatchCode & "'"
                                Log sTmp
                                LogSummary Product, sTmp
                                If fCscript Then wscript.echo vbTab & vbTab & sTmp
                            End If
                        End If 'Err = 0
                        If NOT oFso.FileExists(sLocalMsp) Then 
                            sTmp = vbTab & "Error: Local patch package '" & sLocalMsp & "' missing for patch '" & Patch.PatchCode & "' - '" & Patch.PatchProperty("DisplayName") & "'. Unregistering patch ..."
                            If fDetectOnly Then sTmp = vbTab & "Error: Local patch package '" & sLocalMsp & "' missing for patch '" & Patch.PatchCode & "' - '" & Patch.PatchProperty("DisplayName") & "'. This patch would need to be unregistered!"
                            Log sTmp
                            LogSummary Product, sTmp
                            If fCscript Then wscript.echo vbTab & vbTab & sTmp
                            'Call patch unregister routine
                            If NOT dicMspUnregister.Exists(Patch.PatchCode) Then dicMspUnregister.Add Patch.PatchCode, Patch.PatchCode
                            iCnt = iCnt + 1
                            If iCnt < MAX_ATTEMPT Then fResume = True
                            If NOT fDetectOnly Then UnregisterPatch Patch
                            'Refresh PatchesEx object and resume
                            If NOT fDetectOnly Then Exit For
                            'Reset flag for detect only case
                            fResume = False
                            fMspOk = False
                        Else
                            If NOT fRepairCache Then Log vbTab & "Success: Confirmed local patch package as '" & sLocalMsp & "'" & vbTab & "for patch '" & Patch.PatchCode & "' - '" & Patch.PatchProperty("DisplayName") & "'."
                        End If
                    Next 'Patch
                    If fMspOk Then Log vbTab & "Success: No locally cached .msp packages are missing."
                Else
                    sTmp = vbTab & "Error: PatchesEx API failed with error " & err.number & " - " & err.Description
                    Log sTmp
                    LogSummary Product, sTmp & " (Module MspReconcile)"
                    If fCscript Then wscript.echo vbTab & vbTab & sTmp
                End If 'Err = 0
            Loop
        End If 'IsOfficeProduct
    Next 'Product

End Sub 'MspReconcile
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'   ApplyPatches
'
'   Run patch detection to apply missing & applicable patches
'   Default is to search for patches in
'   A) provided SUpdateLocation folders
'   B) current directory from which script is called
'   C) %windir%\installer\
'-------------------------------------------------------------------------------
Sub ApplyPatches()
    Dim Product, Patch, Key
    Dim sPatchFile, sReturn, sAbsent
    Dim iIndex
    Dim fFeatureControl

    On Error Resume Next
    If fApplyPatch Then sTmp = "Running ApplyPatch" Else sTmp = "Running ApplyPatch for SUpdateLocation folder "
    If fViewPatch Then sTmp = "Running applicable patch detection"
    Log vbCrLf & vbCrLf & sTmp & vbCrLf & String(Len(sTmp), "-")
    If fCscript Then wscript.echo "Running applicable patch detection"
    sPatchFile = ""
    
    ' init the dictionary objects
    Set dicMspNoSeq = CreateObject("Scripting.Dictionary")
    Set dicMspNoBase = CreateObject("Scripting.Dictionary")
    Set dicMspMinor = CreateObject("Scripting.Dictionary")
    Set dicMspSmall = CreateObject("Scripting.Dictionary")
    Set dicMspObsoleted = CreateObject("Scripting.Dictionary")
    Set dicMspSequence = CreateObject("Scripting.Dictionary")

    ' get the patch references from all locations
    ' this calls into the patch details routine as well
    CollectSUpdates
    If IsArray(arrSUpdatesAll) Then
        Log "Debug:  Found " & UBound (arrSUpdatesAll) + 1 & " unique patch(es) in total." & vbCrLf
        If fCscript Then wscript.echo vbTab & "Found " & UBound(arrSUpdatesAll)+1 & " unique patch(es) in total."
    End If
    
    ' loop all products (filter on Office products) and call the pre-sequencer
    For Each Product in oMsi.Products
        If IsOfficeProduct (Product) Then
            Log vbCrLf & "Product: " & Product & " - " & oMsi.ProductInfo (Product, "ProductName") & ", Build: " & oMsi.ProductInfo(Product, "VersionString")
            If fCscript Then wscript.echo vbCrLf & vbTab & "Scan " & Product & " - " & oMsi.ProductInfo(Product, "ProductName") & ", Build: " & oMsi.ProductInfo(Product, "VersionString")
            ' ensure empty dics
            dicMspNoSeq.RemoveAll
            dicMspNoBase.RemoveAll
            dicMspMinor.RemoveAll
            dicMspSmall.RemoveAll
            dicMspObsoleted.RemoveAll
            dicMspSequence.RemoveAll
            dicFeatureStates.RemoveAll
            ' ensure empty value(s)
            sProductVersionReal = ""
            
            ' fill the dictionary objects with a raw list of applicable patches
            GetRawBuckets Product
            
            ' sequence the MinorUpdate bucket first to ensure we get the correct new build number
            SequenceMspMinor Product
            ' sequence the 2.x NoSequence bucket
            SequenceMspNoSeq Product
            ' sequence the BaselineLess bucket
            SequenceMspNoBase Product
            ' sequence the SmallUpdate bucket
            SequenceMspSmall Product
            
            ' log results
            Log vbTab & "Debug:  Baselineless patches bucket contains " & dicMspNoBase.Count & " patch(es) after sequencing."
            Log vbTab & "Debug:  Small patches bucket contains " & dicMspSmall.Count & " patch(es) after sequencing."
            Log vbTab & "Debug:  2.x style patches bucket contains " & dicMspNoSeq.Count & " patch(es) after sequencing."
            Log vbTab & "Debug:  Minor update (service pack) bucket contains " & dicMspMinor.Count & " patch(es) after sequencing."
        
            ' Apply the patches
            ' - invoke msiexec to apply the list of identified patches
            ' - apply each patch in a single/dedicated transaction
            '   this avoids conflicts with custom actions that may configure the REINSTALL list
            ' - ensure the key feature states have not changed
            ' Note: some parts of this happens in  Function ApplyPatch
            If NOT fViewPatch Then
                ' obtain the current feature states
                GetFeatureStates Product
                ' start with baselineless bucket
                sPatchFile = ""
                For Each Key in dicMspNoBase.Keys
                    iIndex = dicMspNoBase.Item (Key)
                    sPatchFile = arrSUpdatesAll (iIndex, COL_FILENAME)
                    fFeatureControl = GetFeatureControl (Product, sPatchFile)
                    If Len (sPatchFile) > 0 Then sReturn = ApplyPatch (Product, sPatchFile, fFeatureControl)
                Next 'Key
                ' execute baseline bucket (minor update aka service pack)
                sPatchFile = ""
                For Each Key in dicMspMinor.Keys
                    iIndex = dicMspMinor.Item (Key)
                    sPatchFile = arrSUpdatesAll (iIndex, COL_FILENAME)
                    fFeatureControl = GetFeatureControl (Product, sPatchFile)
                    If Len (sPatchFile) > 0 Then sReturn = ApplyPatch (Product, sPatchFile, fFeatureControl)
                Next 'Key
                ' execute the 2.x patches bucket
                sPatchFile = ""
                For Each Key in dicMspNoSeq.Keys
                    iIndex = dicMspNoSeq.Item (Key)
                    sPatchFile = arrSUpdatesAll (iIndex, COL_FILENAME)
                    fFeatureControl = GetFeatureControl (Product, sPatchFile)
                    If Len (sPatchFile) > 0 Then sReturn = ApplyPatch (Product, sPatchFile, fFeatureControl)
                Next 'Key
                ' execute the small patches bucket
                For Each Key in dicMspSmall.Keys
                    iIndex = dicMspSmall.Item (Key)
                    sPatchFile = arrSUpdatesAll (iIndex, COL_FILENAME)
                    fFeatureControl = GetFeatureControl (Product, sPatchFile)
                    If Len (sPatchFile) > 0 Then sReturn = ApplyPatch (Product, sPatchFile, fFeatureControl)
                Next 'Key
            End If 'fViewPatch
        End If 'IsOfficeProduct
    Next 'Product

End Sub 'ApplyPatches
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    CollectSUpdates
'
'    The arrUpdateLocations array is sorted and validated by now
'    It cannot be empty since it does at least contain the current directory.
'    As local folders are already sorted to the start of the array this will
'    ensure that local .msp files are favored over network patches.
'    Purpose if this routine is to have a reference and all metadata of
'    available .msp files.
'    This will be used as base to pre-sequence the applicable patches.
'-------------------------------------------------------------------------------
Sub CollectSUpdates()

    Dim File, Folder, SumInfo
    Dim sKey, sPatchTargets, sFilter
    Dim i, iCnt

    On Error Resume Next
    Set dicSUpdatesAll = Nothing
    Set dicSUpdatesAll = CreateObject("Scripting.Dictionary")
    CheckPatchExtract

    iCnt = 0
    'Collect a reference list of all patches
    For Each Folder in arrUpdateLocations
        If fCscript Then wscript.echo vbTab & "Collect files from " & Folder
        For Each File in oFso.GetFolder(Folder).Files
            If LCase(Right(File.Name, 4)) = ".msp" Then
                If LCase(File.Path)=LCase(sApplyPatch) OR sApplyPatch = "" OR (LCase(oFso.GetFolder(Folder).Path) & "\" = LCase(sWICacheDir)) Then
                    Set SumInfo = Nothing
                    Set SumInfo = oMsi.SummaryInformation(File.Path, MSIOPENDATABASEMODE_READONLY)
                    sKey = "" : sPatchTargets = ""
                    sKey = SumInfo.Property(PID_REVNUMBER)
                    sPatchTargets = SumInfo.Property(PID_TEMPLATE)
                    If Not dicSUpdatesAll.Exists(sKey) Then
                        'Found new patch
                        If IsOfficePatch(sPatchTargets) Then
                            dicSUpdatesAll.Add sKey, File.Path
                        Else
                            If NOT LCase(oFso.GetFolder(Folder).Path) & "\" = LCase(sWICacheDir) Then
                                sTmp = "Not an Office patch. Excluding patch " & File.Path & " from detection sequence."
                                Log vbTab & "Debug:  " & sTmp
                            End If
                        End If
                    Else
                        If NOT Left(File.Path, Len(sWICacheDir)) = sWICacheDir Then
                            sTmp = "Excluding patch " & File.Path & " from detection sequence as duplicate of " & dicSUpdatesAll.Item(sKey)
                            Log vbTab & "Debug:  " & sTmp
                            LogSummary "Note:", vbTab & sTmp
                        End If
                    End If 'dicSUpdatesAll.Exists
                End If
            End If '.msp
        Next 'File
        Log "Debug:  Found " & dicSUpdatesAll.Count - iCnt & " unique patch(es) in folder " & Folder
        If fCscript Then wscript.echo vbTab & "Found " & dicSUpdatesAll.Count - iCnt & " unique patch(es) in folder " & Folder
        iCnt = dicSUpdatesAll.Count
    Next 'Folder

    If dicSUpdatesAll.Count = 0 Then Exit Sub

    'Initialize the patch details array
    ReDim arrSUpdatesAll(dicSupdatesAll.Count-1, COL_MAX)

    'Collect all patch details
    i = 0
    If fCscript Then wscript.echo vbTab & "Obtaining patch details for identified patches"
    For Each key in dicSUpdatesAll.Keys
        AddPatchDetails dicSUpdatesAll.Item (key), i
        i = i + 1
    Next 'key

    fUpdatesCollected = True

End Sub 'CollectSUpdates
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    AddPatchDetails
'-------------------------------------------------------------------------------
Sub AddPatchDetails(sMspPath, iIndex)
    Dim SumInfo, Msp, Record
    Dim sSiTmp, sChar, sTitle
    Dim i, iSiCnt
    Dim qView
    Dim arrTitle, arrSi

    On Error Resume Next

    'Defaults
    '--------
    Set Record = Nothing
    arrSUpdatesAll(iIndex, COL_APPLIEDCNT)       = ""
    arrSUpdatesAll(iIndex, COL_SUPERSEDEDCNT)    = ""
    arrSUpdatesAll(iIndex, COL_APPLICABLECNT)    = ""
    arrSUpdatesAll(iIndex, COL_NOQALBASELINECNT) = ""

    'SummaryInformation
    '------------------
    Set SumInfo = oMsi.SummaryInformation(sMspPath, 0)
    arrSUpdatesAll(iIndex, COL_FILENAME)  = sMspPath                        'Msp FileName
    arrSUpdatesAll(iIndex, COL_TARGETS)   = SumInfo.Property(PID_TEMPLATE)  'PatchTargets
    arrSUpdatesAll(iIndex, COL_PATCHCODE) = SumInfo.Property(PID_REVNUMBER) 'PatchCode
    If Len(arrSUpdatesAll(iIndex, COL_PATCHCODE))>LEN_GUID Then
        arrSUpdatesAll(iIndex, COL_SUPERSEDES)=Mid(arrSUpdatesAll(iIndex, COL_PATCHCODE), LEN_GUID+1)
        arrSUpdatesAll(iIndex, COL_PATCHCODE)=Left(arrSUpdatesAll(iIndex, COL_PATCHCODE), LEN_GUID)
    End If
    
    'PatchXml
    '--------
    arrSUpdatesAll(iIndex, COL_PATCHXML) = oMsi.ExtractPatchXMLData(arrSUpdatesAll(iIndex, COL_FILENAME))
    
    'Other
    '-----
    arrSUpdatesAll(iIndex, COL_REFCNT) = 0
    
    'Patch tables
    '------------
    Set Msp = oMsi.OpenDatabase(sMspPath, MSIOPENDATABASEMODE_PATCHFILE)

    If Not Err = 0 Then
        'An error at this points indicates a severe issue
        sTmp = "Failed to read data from .msp package " & sMspPath
        Log vbTab & "Debug:  " & sTmp
        LogSummary "Error:", vbTab & sTmp
        Exit Sub
    End If
    arrSUpdatesAll(iIndex, COL_PATCHTABLES) = GetDatabaseTables(Msp)
    
    If InStr(arrSUpdatesAll(iIndex, COL_PATCHTABLES), "MsiPatchMetadata") > 0 Then
        'KB
        Set qView = Msp.OpenView("SELECT `Property`, `Value` FROM MsiPatchMetadata WHERE `Property`='KBArticle Number'")
        qView.Execute : Set Record = qView.Fetch()
        If Not Record Is Nothing Then
            arrSUpdatesAll(iIndex, COL_KB) = UCase(Record.StringData(2))
            arrSUpdatesAll(iIndex, COL_KB) = Replace(arrSUpdatesAll(iIndex, COL_KB), "KB", "")
        Else
            arrSUpdatesAll(iIndex, COL_KB) = ""
        End If
        qView.Close
        
        'StdPackageName
        Set qView = Msp.OpenView("SELECT `Property`, `Value` FROM MsiPatchMetadata WHERE `Property`='StdPackageName'")
        qView.Execute : Set Record = qView.Fetch()
        If Not Record Is Nothing Then
            arrSUpdatesAll(iIndex, COL_PACKAGE) = Record.StringData(2)
        Else
            arrSUpdatesAll(iIndex, COL_PACKAGE) = ""
        End If
        qView.Close
        
        'Release (required for SP uninstall)
        Set qView = Msp.OpenView("SELECT `Property`, `Value` FROM MsiPatchMetadata WHERE `Property`='Release'")
        qView.Execute : Set Record = qView.Fetch()
        If Not Record Is Nothing Then
            arrSUpdatesAll(iIndex, COL_RELEASE) = Record.StringData(2)
        Else
            arrSUpdatesAll(iIndex, COL_RELEASE) = ""
        End If
        qView.Close
    Else
        arrSUpdatesAll(iIndex, COL_KB) = ""
        arrSUpdatesAll(iIndex, COL_PACKAGE) = ""
        arrSUpdatesAll(iIndex, COL_RELEASE) = ""
    End If
    
    If arrSUpdatesAll(iIndex, COL_KB) = "" Then
        'Scan the SummaryInformation data for the KB
        For iSiCnt = 1 To 2
            Select Case iSiCnt
            Case 1
                arrSi = Split(SumInfo.Property(PID_SUBJECT), ";")
            Case 2
                arrSi = Split(SumInfo.Property(PID_TITLE), ";")
            End Select
            
            If IsArray(arrSi) Then
                For Each sTitle in arrSi
                    sSiTmp = ""
                    sSiTmp = Replace(UCase(sTitle), " ", "")
                    If InStr(sSiTmp, "KB")>0 Then
                        'Strip the KB
                        sSiTmp = Mid(sSiTmp, InStr(sSiTmp, "KB")+2)
                        For i = 1 To Len(sSiTmp)
                            sChar = ""
                            sChar = Mid(sSiTmp, i, 1)
                            If (Asc(sChar) >= 48 AND Asc(sChar) <= 57) Then arrSUpdatesAll(iIndex, COL_KB)=arrSUpdatesAll(iIndex, COL_KB) & sChar
                        Next 'i
                        'Ensure a valid length
                        If Len(arrSUpdatesAll(iIndex, COL_KB))<5 Then arrSUpdatesAll(iIndex, COL_KB)="" Else Exit For
                    End If
                Next
                If Len(arrSUpdatesAll(iIndex, COL_KB))>4 Then Exit For
            End If 'IsArray(arrSi)
            Next 'iSiCnt
    End If
    
    'PatchSequence & PatchFamily
    If InStr(arrSUpdatesAll(iIndex, COL_PATCHTABLES), "MsiPatchSequence") > 0 Then
        Set qView = Msp.OpenView("SELECT `PatchFamily`, `Sequence` FROM MsiPatchSequence")
        qView.Execute : Set Record = qView.Fetch()
        If Not Record Is Nothing Then
            Do Until Record Is Nothing
                arrSUpdatesAll(iIndex, COL_FAMILY) = arrSUpdatesAll(iIndex, COL_FAMILY) & ";" & Record.StringData(1)
                arrSUpdatesAll(iIndex, COL_SEQUENCE) = arrSUpdatesAll(iIndex, COL_SEQUENCE) & ";" & Record.StringData(2)
                Set Record = qView.Fetch()
            Loop
            arrSUpdatesAll(iIndex, COL_FAMILY) = Mid(arrSUpdatesAll(iIndex, COL_FAMILY), 2)
            arrSUpdatesAll(iIndex, COL_SEQUENCE) = Mid(arrSUpdatesAll(iIndex, COL_SEQUENCE), 2)
        Else
            arrSUpdatesAll(iIndex, COL_FAMILY) = ""
            arrSUpdatesAll(iIndex, COL_SEQUENCE) = "0"
        End If
        qView.Close
    Else
        arrSUpdatesAll(iIndex, COL_FAMILY) = ""
        arrSUpdatesAll(iIndex, COL_SEQUENCE) = "0"
    End If

    arrTitle = Split(SumInfo.Property(PID_TITLE), ";")
    If UBound(arrTitle)>0 Then
        If arrSUpdatesAll(iIndex, COL_FAMILY)="" Then arrSUpdatesAll(iIndex, COL_FAMILY) = arrTitle(1)
        If arrSUpdatesAll(iIndex, COL_PACKAGE)= "" Then arrSUpdatesAll(iIndex, COL_PACKAGE) = arrTitle(1)
    End If
    
    'Exception handler for OCT patches
    If InStr( LCase(arrSUpdatesAll(iIndex, COL_FAMILY)), "setupcustomizationfile") > 0 Then
        arrSUpdatesAll(iIndex, COL_KB) = "n/a (SetupCustomizationFile)"
        arrSUpdatesAll(iIndex, COL_PACKAGE) = "OCT"
        If IsBaselineRequired ("", arrSUpdatesAll(iIndex, COL_PATCHXML)) Then _
        LogSummary "Important Note:", sMspPath & " is a customization patch based on the original release of the OCT. A more recent OCT version is available from  http://www.microsoft.com/en-us/download/details.aspx?id=3795"
        'LogSummary "Important Note:", sMspPath & " is a customization patch based on the original release of the OCT. A more recent OCT version is available from  http://www.microsoft.com/downloads/details.aspx?displaylang=en&FamilyID=73d955c0-da87-4bc2-bbf6-260e700519a8"
    End If
    
End Sub 'AddPatchDetails
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    GetRawBuckets
'
'    This function parses all available patches and 
'    fills the unsequenced buckets for the given Product(Code)
'    Already Installed patches are eliminated here
'-------------------------------------------------------------------------------
Sub GetRawBuckets(sProductCode)
	Dim PatchList, Patch
	Dim sAppliedPatches, sApplicablePatches, sProductVersion, sDetectedVersion
	Dim iBucket, iIndex
	Dim fSkipRealDetection, fIsWICached

	On Error Resume Next

    ' get installed patches
    Set PatchList = oMsi.PatchesEx(sProductCode, USERSID_NULL, MSIINSTALLCONTEXT_MACHINE, MSIPATCHSTATE_ALL)
    sAppliedPatches = ""
    For Each Patch in PatchList
        sAppliedPatches = sAppliedPatches & ";" & Patch.PatchCode
    Next 'Patch
    fSkipRealDetection = False
    If Not Err=0 Then fSkipRealDetection = True

    ' sort the patches to their bucket
    If IsArray(arrSUpdatesAll) Then
        For iIndex = 0 To UBound(arrSUpdatesAll)
	        ' honor the fExcludeCache and fIncludeOctCache flag
            fIsWICached = False
            If Len(arrSUpdatesAll(iIndex, COL_FILENAME)) > Len(sWICacheDir) Then
                fIsWICached = (LCase(Left(arrSUpdatesAll(iIndex, COL_FILENAME), Len(sWICacheDir))) = LCase(sWICacheDir))
            End If
            If ((NOT fExcludeCache) OR (NOT fIsWICached)) AND _
                 NOT ((arrSUpdatesAll(iIndex, COL_PACKAGE) = "OCT") AND (fIsWICached) AND (NOT fIncludeOctCache)) Then
			    ' exclude patches that do
			    ' - not target the product
			    ' - are already applied
                If (InStr(arrSUpdatesAll(iIndex, COL_TARGETS), sProductCode)>0) Then
                    'Update reference counter
                    arrSUpdatesAll(iIndex, COL_REFCNT) = arrSUpdatesAll(iIndex, COL_REFCNT)+1
                    If (NOT InStr(sAppliedPatches, arrSUpdatesAll(iIndex, COL_PATCHCODE))>0) Then
			    	    ' patch targets the current product and is not applied
                        Select Case GetMspBucket(sProductCode, iIndex)
                        Case MSP_NOSEQ
                            dicMspNoSeq.Add arrSUpdatesAll(iIndex, COL_PATCHCODE), iIndex
                        Case MSP_NOBASE
						    dicMspNoBase.Add arrSUpdatesAll(iIndex, COL_PATCHCODE), iIndex
					    Case MSP_MINOR
                            dicMspMinor.Add arrSUpdatesAll(iIndex, COL_PATCHCODE), iIndex
                        Case MSP_SMALL
                            dicMspSmall.Add arrSUpdatesAll(iIndex, COL_PATCHCODE), iIndex
                        Case Else
                        End Select
                    Else
                        If NOT fIsWICached Then
				    	    ' update reference counter
                            arrSUpdatesAll(iIndex, COL_APPLIEDCNT) = arrSUpdatesAll(iIndex, COL_APPLIEDCNT) & sProductCode & ";"
                            sTmp = "Patch KB " & arrSUpdatesAll(iIndex, COL_KB) & " is already installed for this product." & _
                                   " Patch details: " & arrSUpdatesAll(iIndex, COL_PATCHCODE) & ", " & arrSUpdatesAll(iIndex, COL_PACKAGE) & ", " & arrSUpdatesAll(iIndex, COL_FILENAME)
                            Log vbTab & "Debug:  " & sTmp
                            LogSummary sProductCode, vbTab & sTmp
                        End If
                    End If
                Else
                    If NOT fIsWICached Then
                        sTmp = "Patch KB " & arrSUpdatesAll(iIndex, COL_KB) & " (" & arrSUpdatesAll(iIndex, COL_FILENAME) & ") does not target this product."
                        Log vbTab & "Debug:  " & sTmp
                    End If
                End If
            Else
                If fExcludeCache Then sTmp="ExcludeCache=True" Else sTmp="IncludeOctCache=False"
                If (InStr(arrSUpdatesAll(iIndex, COL_TARGETS), sProductCode)>0) Then _
                Log vbTab & "Debug:  Excluding patch per '" & sTmp & "' filter. " & arrSUpdatesAll(iIndex, COL_PATCHCODE) & ", " & arrSUpdatesAll(iIndex, COL_KB) & ", " & arrSUpdatesAll(iIndex, COL_PACKAGE) & ", " & arrSUpdatesAll(iIndex, COL_FILENAME)
            End If 'fExcludeCache
        Next 'iIndex
    End If 'IsArray
    
    'Validate integrity of the registered ProductVersion (build) as the sequencing
    'logic relies on the correctness of this value.
    If NOT fSkipRealDetection Then
        sProductVersion  = oMsi.ProductInfo(sProductCode, "VersionString")
        sDetectedVersion = GetRealBuildVersion(sAppliedPatches, sProductCode)
        If NOT sProductVersion = sDetectedVersion Then
            sTmp = vbTab & "Error: Registered build version does not match detected build version. Registered build: " & sProductVersion & ". Detected build: " & sDetectedVersion
            If fDetectOnly Then
                sTmp = sTmp & ". Registered build would be corrected to " & sDetectedVersion
            Else
                sTmp = sTmp & ". Updated registered build to new value " & sDetectedVersion
                UpdateProductVersion sProductCode, sDetectedVersion
            End If
            Log sTmp
            LogSummary sProductCode, sTmp
        End If
    End If
    
End Sub 'GetRawBuckets
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    GetMspBucket
'-------------------------------------------------------------------------------
Function GetMspBucket(sProductCode, iIndex)

    On Error Resume Next
    GetMspBucket = MSP_NOSEQ 'Default to 2.x NoSequence bucket
    ' check if it's a 3.x type patch which has sequence information
    If arrSUpdatesAll(iIndex, COL_SEQUENCE) = "0" Then
        ' this is a 2.x patch. Only continue if it's a service pack
        If NOT IsMinorUpdate(sProductCode, arrSUpdatesAll(iIndex, COL_PATCHXML)) Then Exit Function
    End If
    
    ' check if it's a BaselineLess patch
    If NOT IsBaselineRequired (sProductCode, arrSUpdatesAll(iIndex, COL_PATCHXML)) Then
        GetMspBucket = MSP_NOBASE
        Exit Function
    End If
    
    ' check if it's a "Minor Upgrade" aka "Service Pack" vs. as "Small Update"
    If IsMinorUpdate(sProductCode, arrSUpdatesAll(iIndex, COL_PATCHXML)) Then GetMspBucket = MSP_MINOR Else GetMspBucket = MSP_SMALL
End Function 'GetMspBucket
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    SequenceMspMinor
'
'    Sequence the Minor Update (service pack) bucket
'    The logic relies on the Office specific assumption that a service pack:
'       - is cumulative
'        - always uses "Equals" as baseline verification
'-------------------------------------------------------------------------------
Sub SequenceMspMinor(sProductCode)
    Dim Key, Patch, Sequences, Seq
    Dim sProductVersion, sProductVersionMax, sMspApplicable, sFamily, sSeq, sErr
    Dim iCntBld, iCntMsp, iIndex, iPos
    Dim fSeqFound, fHihgerBaselineExists
    Dim arrMspUpdatedVersions, arrMspSuperseded, dicMspUpdatedVersion, arrErr

    On Error Resume Next
    fSeqFound = False
    sMspApplicable = ""
    Set dicMspUpdatedVersion = CreateObject("Scripting.Dictionary")

    'Get the current product build
    sProductVersion = sProductVersionReal
    sProductVersionNew = sProductVersion

    'Get the updated build versions. Sorted descending
    'This call will already filter out superseded patches
    arrMspUpdatedVersions = GetMspUpdatedVersion(sProductCode, dicMspUpdatedVersion)

    'Check if there's an updated version available
    fHihgerBaselineExists = (UBound(arrMspUpdatedVersions)> -1)

    'Iterate the patches if we have a higher baseline than the current
    If fHihgerBaselineExists Then
        sProductVersionMax =arrMspUpdatedVersions(0)
        'Find applicable patch sequence
        For iCntBld = 0 To UBound(arrMspUpdatedVersions)
            For Each Key in dicMspMinor.Keys
                iIndex = dicMspMinor.Item(Key)
                If IsValidVersion (sProductCode, arrSUpdatesAll (iIndex, COL_PATCHXML), sProductVersionNew, sErr, iPos) Then
                    'Found new baseline. Add patch as applicable
                    'Remember new baseline
                    sProductVersionNew = dicMspUpdatedVersion.Item(iIndex)
                    fHihgerBaselineExists = (sProductVersionMax>sProductVersionNew)
                    arrErr = Split(sErr, ";", 2)
                    'Update reference counter
                    arrSUpdatesAll(iIndex, COL_APPLICABLECNT) = arrSUpdatesAll(iIndex, COL_APPLICABLECNT) & sProductCode & ";"
                    sTmp = "Found applicable service pack patch to update build from " & arrErr(1) & " to build " & sProductVersionNew & " : KB " & arrSUpdatesAll(iIndex, COL_KB) & _
                        ", " & arrSUpdatesAll(iIndex, COL_PATCHCODE) & ", " & arrSUpdatesAll(iIndex, COL_PACKAGE) & ", " & arrSUpdatesAll(iIndex, COL_FILENAME)
                    sMspApplicable = sMspApplicable & ";" & Key
                    Log vbTab & "Debug:  " & sTmp
                    LogSummary sProductCode, vbTab & sTmp
                    Exit For
                End If
            Next 'Key
            If NOT fHihgerBaselineExists Then Exit For
        Next 'iCntBld
    End If 'fHihgerBaselineExists

    For Each Key in dicMspMinor.Keys
        If NOT InStr(sMspApplicable, Key)>0 Then
            'patch excluded because higher baseline has been found
            iIndex = dicMspMinor.Item(Key)
            'Update reference counter
            arrSUpdatesAll(iIndex, COL_SUPERSEDEDCNT) = arrSUpdatesAll(iIndex, COL_SUPERSEDEDCNT) & sProductCode & ";"
            sTmp = "Excluding patch KB " & arrSUpdatesAll(iIndex, COL_KB) & " because it's superseded by a scheduled service pack installation." & _
                   " Patch details: " & arrSUpdatesAll(iIndex, COL_PATCHCODE) & ", " & arrSUpdatesAll(iIndex, COL_PACKAGE) & ", " & arrSUpdatesAll(iIndex, COL_FILENAME)
            If dicMspMinor.Exists(Key) Then dicMspMinor.Remove Key
            Log vbTab & "Debug:  " & sTmp
            If (NOT Left(arrSUpdatesAll(iIndex, COL_FILENAME), Len(sWICacheDir)) = sWICacheDir) Then LogSummary sProductCode, vbTab & sTmp
        End If
    Next 'Key

    'Add patch family sequence data if available
    For Each Key in dicMspMinor.Keys
        iIndex = dicMspMinor.Item(Key)
        AddSequenceData arrSUpdatesAll(iIndex, COL_PATCHXML)
    Next 'Key

    'Add obsoletion data
    For Each Key in dicMspMinor.Keys
        Set arrMspSuperseded = Nothing
        iIndex = dicMspMinor.Item(Key)
        arrMspSuperseded = Split(arrSUpdatesAll(iIndex, COL_SUPERSEDES), ";")
        For Each Patch in arrMspSuperseded
            If NOT dicMspObsoleted.Exists(Patch) Then dicMspObsoleted.Add Patch, Patch
        Next 'Patch
    Next 'Key

End Sub 'SequenceMspMinor
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    SequenceMspNoSeq
'
'    Sequence the bucket with 2.x NoSequence patches
'    Superseded (obsoleted) patches are filtered out 
'-------------------------------------------------------------------------------
Sub SequenceMspNoSeq(sProductCode)
    Dim Key, Patch
    Dim sErr
    Dim iIndex, iPos
    Dim arrMspSuperseded, arrErr

    On Error Resume Next
    
    ' build list of obsoleted patches
    For Each Key in dicMspNoSeq.Keys
        Set arrMspSuperseded = Nothing
        iIndex = dicMspNoSeq.Item(Key)
        arrMspSuperseded = Split(arrSUpdatesAll(iIndex, COL_SUPERSEDES), ";")
        For Each Patch in arrMspSuperseded
            If NOT dicMspObsoleted.Exists(Patch) Then dicMspObsoleted.Add Patch, Patch
        Next 'Patch
    Next 'Key

    For Each Key in dicMspNoSeq.Keys
        iIndex = dicMspNoSeq.Item(Key)
        ' remove patch if obsolete
        If dicMspObsoleted.Exists(Key) Then
            ' update reference counter
            arrSUpdatesAll(iIndex, COL_SUPERSEDEDCNT) = arrSUpdatesAll(iIndex, COL_SUPERSEDEDCNT) & sProductCode & ";"
            If NOT Left(arrSUpdatesAll(iIndex, COL_FILENAME), Len(sWICacheDir)) = sWICacheDir Then
                sTmp = "Patch KB " & arrSUpdatesAll(iIndex, COL_KB) & " is obsoleted by an already installed patch." & _
                       " Patch details: " & arrSUpdatesAll(iIndex, COL_PATCHCODE) & ". " & arrSUpdatesAll(iIndex, COL_PACKAGE) & ", " & arrSUpdatesAll(iIndex, COL_FILENAME)
                Log vbTab & "Debug:  " & sTmp
                LogSummary sProductCode, vbTab & sTmp
            End If
            If dicMspNoSeq.Exists(Key) Then dicMspNoSeq.Remove Key
        End If
        
        ' check if patch is applicable
        If IsValidVersion (sProductCode, arrSUpdatesAll (iIndex, COL_PATCHXML), sProductVersionNew, sErr, iPos) Then
            arrErr = Split(sErr, ";", 2)
            
            ' update reference counter
            arrSUpdatesAll(iIndex, COL_APPLICABLECNT) = arrSUpdatesAll(iIndex, COL_APPLICABLECNT) & sProductCode & ";"
            sTmp = "Found applicable 2.x style patch: KB " & arrSUpdatesAll(iIndex, COL_KB) & ", " & arrSUpdatesAll(iIndex, COL_PATCHCODE) & ", " & arrSUpdatesAll(iIndex, COL_PACKAGE) & ", " & arrSUpdatesAll(iIndex, COL_FILENAME) & _
                   vbCrLf & vbTab & vbTab & "Applicable baseline: " & arrErr(1)
        Else
            arrErr = Split(sErr, ";", 2)
            ' update reference counter
            arrSUpdatesAll(iIndex, COL_NOQALBASELINECNT) = arrSUpdatesAll(iIndex, COL_NOQALBASELINECNT) & sProductCode & ";"
            
            ' cache valid baselines
            arrSUpdatesAll(iIndex, COL_PATCHBASELINES) = arrErr(1)
            sTmp = "No valid baseline available for this 2.x style patch KB " & arrSUpdatesAll(iIndex, COL_KB) & "." & _
                   " Patch details: " & arrSUpdatesAll(iIndex, COL_PATCHCODE) & ", " & arrSUpdatesAll(iIndex, COL_PACKAGE) & ", " & arrSUpdatesAll(iIndex, COL_FILENAME) & _
                   vbCrLf & vbTab & vbTab & "Patch baseline(s): " & arrErr(1) & ". Installed baseline: " & sProductVersionNew
            If dicMspNoSeq.Exists(Key) Then dicMspNoSeq.Remove Key
        End If 'IsValidVersion
        Log vbTab & "Debug:  " & sTmp
        LogSummary sProductCode, vbTab & sTmp
    Next 'Key

End Sub 'SequenceMspNoSeq
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    SequenceMspNoBase
'
'    Sequence patches that do not validate the baseline
'-------------------------------------------------------------------------------
Sub SequenceMspNoBase (sProductCode)
	Dim Key, Element, Elements
	Dim sMspApplicable, sFamily, sSeq, sErr, sAttr
	Dim iIndex
	Dim fApplicable

    ' determine current patch family sequence
    For Each Key in dicMspNoBase.Keys
		iIndex = dicMspNoBase.Item(Key)
        AddSequenceData(arrSUpdatesAll(iIndex, COL_PATCHXML))
	Next 'Key

    ' determine applicable patches
	For Each Key in dicMspNoBase.Keys
		fApplicable = False
		iIndex = dicMspNoBase.Item(Key)
		XmlDoc.LoadXml(arrSUpdatesAll(iIndex, COL_PATCHXML))
		Set Elements = XmlDoc.GetElementsByTagName("SequenceData")
		For Each Element in Elements
			sFamily = "" : sSeq = "" : sAttr = ""
			sFamily = Element.selectSingleNode("PatchFamily").text
			sSeq = Element.selectSingleNode("Sequence").text
            sAttr = Element.selectSingleNode("Attributes").text
			sTmp = "Found applicable baselineless patch with sequence version " & sSeq & " : KB " & arrSUpdatesAll(iIndex, COL_KB) & ", " & arrSUpdatesAll(iIndex, COL_PATCHCODE) & ", " & arrSUpdatesAll(iIndex, COL_PACKAGE) & ", " & arrSUpdatesAll(iIndex, COL_FILENAME)
			If dicMspSequence.Exists(sFamily) Then
				If sSeq = dicMspSequence.Item(sFamily) Then
					fApplicable = True
				End If
			Else
				fApplicable = True
			End If
            ' ensure to flag as applicable if the patch does not supersede previous versions
            If sAttr = "0" Then fApplicable = True
		Next 'Element
		
		If NOT fApplicable Then
			'patch excluded because higher family patch available
			'Update reference counter
			arrSUpdatesAll(iIndex, COL_SUPERSEDEDCNT) = arrSUpdatesAll(iIndex, COL_SUPERSEDEDCNT) & sProductCode & ";"
			sTmp = "Patch KB " & arrSUpdatesAll(iIndex, COL_KB) & " is superseded by a later patch of the same family." & _
				   " Patch details: " & arrSUpdatesAll(iIndex, COL_PATCHCODE) & ", " & arrSUpdatesAll(iIndex, COL_PACKAGE) & ", " & arrSUpdatesAll(iIndex, COL_FILENAME) & _
				   vbCrLf & vbTab & vbTab & "Patch build: " & sSeq & ". Installed build: " & dicMspSequence.Item(sFamily)
			If dicMspNoBase.Exists(Key) Then dicMspNoBase.Remove Key
			Log vbTab & "Debug:  " & sTmp
			If (NOT Left(arrSUpdatesAll(iIndex, COL_FILENAME), Len(sWICacheDir)) = sWICacheDir) Then LogSummary sProductCode, vbTab & sTmp
		Else
			'Update reference counter
			arrSUpdatesAll(iIndex, COL_APPLICABLECNT) = arrSUpdatesAll(iIndex, COL_APPLICABLECNT) & sProductCode & ";"
			Log vbTab & "Debug:  " & sTmp
			LogSummary sProductCode, vbTab & sTmp
		End If
	Next 'Key
End Sub 'SequenceMspNoBase
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    SequenceMspSmall
'
'    Sequence small patches
'-------------------------------------------------------------------------------
Sub SequenceMspSmall(sProductCode)
	Dim Key, Element, Elements
	Dim sMspApplicable, sFamily, sSeq, sErr, sAttr
	Dim iIndex, iPos
	Dim fApplicable
	Dim arrErr

	On Error Resume Next
	sErr = ""
    
    ' determine current patch family sequence
	For Each Key in dicMspSmall.Keys
		iIndex = dicMspSmall.Item(Key)
    
	    ' load baselines the patch can be applied to
	    ' exclude patches that do not target the current baseline
		If IsValidVersion (sProductCode, arrSUpdatesAll (iIndex, COL_PATCHXML), sProductVersionNew, sErr, iPos) Then
			AddSequenceData(arrSUpdatesAll(iIndex, COL_PATCHXML))
		Else
			arrErr = Split(sErr, ";", 2)
			If arrErr(0) = "1" Then
				'Patch is superseded by the installed baseline (service pack)
				'Update reference counter
				arrSUpdatesAll(iIndex, COL_SUPERSEDEDCNT) = arrSUpdatesAll(iIndex, COL_SUPERSEDEDCNT) & sProductCode & ";"
				sTmp = "Patch KB " & arrSUpdatesAll(iIndex, COL_KB) & " is superseded by an already installed service pack." & _
					   " Patch details: " & arrSUpdatesAll(iIndex, COL_PATCHCODE) & ", " & arrSUpdatesAll(iIndex, COL_PACKAGE) & ", " & arrSUpdatesAll(iIndex, COL_FILENAME) & _
					   vbCrLf & vbTab & vbTab & "Patch baseline(s): " & arrErr(1) & ". Installed baseline: " & sProductVersionNew
			Else
				'Patch excluded because it does not apply to the available baseline
				'Cache valid baselines
				arrSUpdatesAll(iIndex, COL_PATCHBASELINES) = arrErr(1)
				
                'Update reference counter
				arrSUpdatesAll(iIndex, COL_NOQALBASELINECNT) = arrSUpdatesAll(iIndex, COL_NOQALBASELINECNT) & sProductCode & ";"
				sTmp = "No valid baseline available for patch KB " & arrSUpdatesAll(iIndex, COL_KB) & "." & _
					   " Patch details: " & arrSUpdatesAll(iIndex, COL_PATCHCODE) & ", " & arrSUpdatesAll(iIndex, COL_PACKAGE) & ", " & arrSUpdatesAll(iIndex, COL_FILENAME) & _
					   vbCrLf & vbTab & vbTab & "Patch baseline(s): " & arrErr(1) & ". Installed baseline: " & sProductVersionNew
			End If
			Log vbTab & "Debug:  " & sTmp
			If (NOT Left(arrSUpdatesAll(iIndex, COL_FILENAME), Len(sWICacheDir)) = sWICacheDir) Then LogSummary sProductCode, vbTab & sTmp
			If dicMspSmall.Exists(Key) Then dicMspSmall.Remove Key
		End If
	Next 'Key

    ' determine applicable patches
	For Each Key in dicMspSmall.Keys
		fApplicable = False
		iIndex = dicMspSmall.Item(Key)
		XmlDoc.LoadXml(arrSUpdatesAll(iIndex, COL_PATCHXML))
		Set Elements = XmlDoc.GetElementsByTagName("SequenceData")
		For Each Element in Elements
			sFamily = "" : sSeq = "" : sAttr = ""
			sFamily = Element.selectSingleNode("PatchFamily").text
			sSeq = Element.selectSingleNode("Sequence").text
            sAttr = Element.selectSingleNode("Attributes").text
			sTmp = "Found applicable patch with sequence version " & sSeq & " : KB " & arrSUpdatesAll(iIndex, COL_KB) & ", " & arrSUpdatesAll(iIndex, COL_PATCHCODE) & ", " & arrSUpdatesAll(iIndex, COL_PACKAGE) & ", " & arrSUpdatesAll(iIndex, COL_FILENAME)
			If dicMspSequence.Exists(sFamily) Then
				If sSeq = dicMspSequence.Item(sFamily) Then
					fApplicable = True
				End If
			Else
				fApplicable = True
			End If
            
            ' ensure to flag as applicable if the patch does not supersede previous versions
            If sAttr = "0" Then fApplicable = True
		Next 'Element
		
		If NOT fApplicable Then
			'patch excluded because higher family patch available
			'Update reference counter
			arrSUpdatesAll(iIndex, COL_SUPERSEDEDCNT) = arrSUpdatesAll(iIndex, COL_SUPERSEDEDCNT) & sProductCode & ";"
			sTmp = "Patch KB " & arrSUpdatesAll(iIndex, COL_KB) & " is superseded by a later patch of the same family." & _
				   " Patch details: " & arrSUpdatesAll(iIndex, COL_PATCHCODE) & ", " & arrSUpdatesAll(iIndex, COL_PACKAGE) & ", " & arrSUpdatesAll(iIndex, COL_FILENAME) & _
				   vbCrLf & vbTab & vbTab & "Patch build: " & sSeq & ". Installed build: " & dicMspSequence.Item(sFamily)
			If dicMspSmall.Exists(Key) Then dicMspSmall.Remove Key
			Log vbTab & "Debug:  " & sTmp
			If (NOT Left(arrSUpdatesAll(iIndex, COL_FILENAME), Len(sWICacheDir)) = sWICacheDir) Then LogSummary sProductCode, vbTab & sTmp
		Else
			'Update reference counter
			arrSUpdatesAll(iIndex, COL_APPLICABLECNT) = arrSUpdatesAll(iIndex, COL_APPLICABLECNT) & sProductCode & ";"
			Log vbTab & "Debug:  " & sTmp
			LogSummary sProductCode, vbTab & sTmp
		End If
	Next 'Key

End Sub 'SequenceMspSmall
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    GetMspUpdatedVersion
'
'    Return the updated version (build) of a minor update (service pack)
'    sorted from highest to lowest
'-------------------------------------------------------------------------------
Function GetMspUpdatedVersion(sProductCode, dicMspUpdatedVersion)
    Dim Key, Element, Elements
    Dim sVersions, sProductVersion, sProductVersionMsi, sErr
    Dim iIndex, iArrCnt, iPos
    Dim arrVersions, arrErr, dicTmp

    On Error Resume Next
    Set dicTmp = CreateObject("Scripting.Dictionary")
    
    ' get the current product build
    sProductVersion = sProductVersionReal
    sProductVersionMsi = ""
    sProductVersionMsi = GetMsiProductVersion(oMsi.ProductInfo(sProductCode, "LocalPackage"))
    If sProductVersionMsi = "" Then sProductVersionMsi = sProductVersionReal
    
    ' identify the available updated build (UpdatedVersion)
    sVersions = "" : sErr = ""
    For Each Key in dicMspMinor.Keys
        iIndex = dicMspMinor.Item(Key)
        'Don't assume we have a valid RTM build. Beta products may break the logic!
        XmlDoc.LoadXml(arrSUpdatesAll(iIndex, COL_PATCHXML))
        Set Elements = XmlDoc.GetElementsByTagName("TargetProduct")
        For Each Element in Elements
            If Element.selectSingleNode("TargetProductCode").text = sProductCode Then
                If IsValidVersion (sProductCode, arrSUpdatesAll (iIndex, COL_PATCHXML), sProductVersionMsi, sErr, iPos) Then
                    If Element.selectSingleNode("UpdatedVersion").text > sProductVersion Then
                        If NOT dicTmp.Exists(iIndex) Then 
                            dicTmp.Add iIndex, Element.selectSingleNode("UpdatedVersion").text
                            sVersions = sVersions & ";" & Element.selectSingleNode("UpdatedVersion").text
                        End If
                    Else
                        'patch excluded since not a higher baseline
                        'Update reference counter
                        arrSUpdatesAll(iIndex, COL_SUPERSEDEDCNT) = arrSUpdatesAll(iIndex, COL_SUPERSEDEDCNT) & sProductCode & ";"
                        sTmp = "Service pack patch KB " & arrSUpdatesAll(iIndex, COL_KB) & " is superseded by an already installed service pack." & _
                               " Patch details: " & arrSUpdatesAll(iIndex, COL_PATCHCODE) & ", " & arrSUpdatesAll(iIndex, COL_PACKAGE) & ", " & arrSUpdatesAll(iIndex, COL_FILENAME) & _
                               vbCrLf & vbTab & vbTab & "Patch build: " & Element.selectSingleNode("UpdatedVersion").text & ", Installed build: " & sProductVersion
                        Log vbTab & "Debug:  " & sTmp
                        LogSummary sProductCode, vbTab & sTmp
                        If dicMspMinor.Exists(Key) Then dicMspMinor.Remove Key
                    End If 'UpdatedVersion
                Else
                    'Not a RTM product
                    arrErr = Split(sErr, ";", 2)
                    sTmp = "No valid baseline available for service pack patch KB " & arrSUpdatesAll(iIndex, COL_KB) & ". This may indicate a BETA ProductVersion." & _
                            " Patch details: " & arrSUpdatesAll(iIndex, COL_PATCHCODE) & ", " & arrSUpdatesAll(iIndex, COL_PACKAGE) & ", " & arrSUpdatesAll(iIndex, COL_FILENAME) & _
                            bCrLf & vbTab & vbTab & "Patch baseline(s): " & arrErr(1) & ". Installed baseline: " & sProductVersion
                    Log vbTab & "Debug:  " & sTmp
                    LogSummary sProductCode, vbTab & sTmp
                End If 'IsValidVersion
            End If 'TargetProductCode
        Next 'Element
    Next 'Key

    If sVersions = "" Then
        Redim arrVersions(-1)
        GetMspUpdatedVersion = arrVersions
        Exit Function
    End If

    'Sort descending
    arrVersions = BubbleSort(Split(Mid(sVersions, 2), ";"))

    'Build the dictionary
    For iArrCnt = 0 To UBound(arrVersions)
        For Each Key in dicTmp.Keys
            If dicTmp.Item(Key)=arrVersions(iArrCnt) AND NOT dicMspUpdatedVersion.Exists(Key) Then dicMspUpdatedVersion.Add Key, dicTmp.Item(Key)
        Next 'Key
    Next 'iArrCnt

    'Return the sorted versions
    GetMspUpdatedVersion = arrVersions

End Function 'GetMspUpdatedVersion
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    IsValidVersion
'
'    Determine if a patch is applicable to the provided baseline
'-------------------------------------------------------------------------------
Function IsValidVersion(sProductCode, sXml, sProductVersion, sErr, iPos)
    Dim Element, Elements, Node
    Dim sCompType, sCompFlt, sVersion, sDelimiter, sTargets
    Dim iCnt, iLoop, iRet
    Dim fSuccess, fValidate
    Dim arrLeftNum, arrRightNum

    On Error Resume Next
    sErr = "" : sTargets = "" 
    iPos = 0
    fValidate = True
    sDelimiter = Delimiter(sProductVersion)
    arrLeftNum = Split(sProductVersion, sDelimiter)

    XmlDoc.LoadXml(sXml)
    Set Elements = XmlDoc.GetElementsByTagName("TargetProduct")
    For Each Element in Elements
        fSuccess = False
        If Element.selectSingleNode("TargetProductCode").text = sProductCode Then
            'Collect the compare details
            sCompType = "" : sCompFlt = "" : sVersion = ""
            Set Node = Element.selectSingleNode("TargetVersion")
            sCompType = Node.getAttribute("ComparisonType")
            sCompFlt = Node.getAttribute("ComparisonFilter")
            sVersion = Node.text
            sTargets = sTargets & ";" & sVersion
            fValidate = CBool(Node.getAttribute("Validate"))
            Set arrRightNum = Nothing
            arrRightNum = Split(sVersion, sDelimiter)
        
            'Set the filter setting
            Select Case sCompFlt
            Case "None"
                iLoop = -1
            Case "Major"
                iLoop = 0
            Case "MajorMinor"
                iLoop = 1
            Case "MajorMinorUpdate"
                iLoop = 2
            Case Else
            End Select
        
            'Compare the version strings based on the filter
            iRet = -2
            For iCnt = 0 To iLoop
                iRet = StrComp(arrLeftNum(iCnt), arrRightNum(iCnt))
                If NOT iRet = 0 Then Exit For
            Next 'iCnt

            'Evaluate the compare result
            Select Case sCompType
            Case "LessThan"
                fSuccess = (iRet = -1)
            Case "LessThanOrEqual"
                fSuccess = ((iRet = -1) OR (iRet = 0))
            Case "Equal"
                fSuccess = (iRet = 0)
            Case "GreaterThanOrEqual"
                fSuccess = ((iRet = 1) OR (iRet = 0))
            Case "GreaterThan"
                fSuccess = (iRet = 1)
            Case "None"
                fSuccess = True
            Case Else
            End Select
        
            If NOT fValidate Then fSuccess = True
            If fSuccess Then Exit For
        End If
        iPos = iPos + 1
    Next

    If fSuccess Then sErr = iRet & ";" & sVersion Else sErr = iRet & ";" & Join(RemoveDuplicates(Split(Mid(sTargets, 2), ";")), ";")
    IsValidVersion = fSuccess

End Function 'IsValidVersion
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    GetDatabaseTables
'
'    Return the tables of a given .msp file
'-------------------------------------------------------------------------------
Function GetDatabaseTables(MsiDb)
	Dim ViewTables, Table
	Dim sTables

    On Error Resume Next
    sTables = ""
    Set Table = Nothing
    Set ViewTables = MsiDb.OpenView("SELECT `Name` FROM `_Tables` ORDER BY `Name`")
    ViewTables.Execute
    Do
        Set Table = ViewTables.Fetch
        If Table Is Nothing then Exit Do
        sTables = sTables & "," & Table.StringData(1)
    Loop
    ViewTables.Close
    If Len(sTables) > 2 Then GetDatabaseTables = Mid(sTables, 2)

End Function 'GetDatabaseTables
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    GetDatabaseStreams
'
'    Return Streams of a given installer database
'-------------------------------------------------------------------------------
Function GetDatabaseStreams(MsiDb)
    Dim DbStreams, Record
    Dim sStreams

    On Error Resume Next
    sStreams = ""
    Set Record = Nothing
    Set DbStreams = MsiDb.OpenView("SELECT * FROM _Streams") : DbStreams.Execute
    Do
        Set Record = DbStreams.Fetch
        If Record Is Nothing Then Exit Do
        sStreams = sStreams & "," & Record.StringData(1)
    Loop
    DbStreams.Close
    If Len(sStreams)>2  Then GetDatabaseStreams = Mid(sStreams, 2) Else GetDatabaseStreams = ""

End Function 'GetDatabaseStreams
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    MspBldTargets
'
'    Returns the possible build versions the patch will be compared against
'-------------------------------------------------------------------------------
Function MspBldTargets(sProductCode, sXml)
    Dim Element, Elements, Node
    Dim sBaselines

    On Error Resume Next
    sBaselines = ""
    
    'Check baselines from XML
    XmlDoc.LoadXml(sXml)
    Set Elements = XmlDoc.GetElementsByTagName("TargetProduct")
    For Each Element in Elements
        If Element.selectSingleNode("TargetProductCode").text = sProductCode Then
            Set Node = Element.selectSingleNode("TargetVersion")
            sBaselines = sBaselines & ";" & Node.text
        End If
    Next

    If Len(sBaselines)>1 Then sBaselines = Mid(sBaselines, 2)
    MspBldTargets = sBaselines

End Function
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    GetMspBldTargets
'
'    Returns all possible build versions the patch will be compared against
'-------------------------------------------------------------------------------
Function GetMspBldTargets(sXml)
	Dim Element, Elements, Node
	Dim sBaselines

	On Error Resume Next
	sBaselines = ""
	
    'Check baselines from XML
	XmlDoc.LoadXml(sXml)
	Set Elements = XmlDoc.GetElementsByTagName("TargetProduct")
	For Each Element in Elements
		Set Node = Element.selectSingleNode("TargetVersion")
		If NOT InStr(sBaselines, Node.text)>0 Then 
			sBaselines = sBaselines & ";" & Node.text & GetSpLevel(Node.text)
		End If
	Next

	If Len(sBaselines)>1 Then sBaselines = Mid(sBaselines, 2)
	If NOT IsBaselineRequired("", sXml) Then sBaselines = "Baselineless Patch"
	GetMspBldTargets = Replace(sBaselines, ";", vbCrLf)
End Function
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    GetComparisonType
'
'    Returns the comparison term for the build version validation
'-------------------------------------------------------------------------------
Function GetComparisonType(sProductCode, sXml)
	Dim Element, Elements, Node

	On Error Resume Next
	
    XmlDoc.LoadXml(sXml)
	Set Elements = XmlDoc.GetElementsByTagName("TargetProduct")
	For Each Element in Elements
		If Element.selectSingleNode("TargetProductCode").text = sProductCode Then
			Set Node = Element.selectSingleNode("TargetVersion")
			GetComparisonType = Node.getAttribute("ComparisonType")
			Exit For
		End If
	Next

End Function 'GetComparisonType
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    IsBaselineRequired
'
'    Determines from patch xml if patch requires baseline validation
'-------------------------------------------------------------------------------
Function IsBaselineRequired(sProductCode, sXml)
	Dim Element, Elements, Node

	On Error Resume Next
	
    XmlDoc.LoadXml(sXml)
	Set Elements = XmlDoc.GetElementsByTagName("TargetProduct")
	For Each Element in Elements
		If sProductCode = "" Then sProductCode = Element.selectSingleNode("TargetProductCode").text
		If Element.selectSingleNode("TargetProductCode").text = sProductCode Then
			Set Node = Element.selectSingleNode("TargetVersion")
			IsBaselineRequired = CBool(Node.getAttribute("Validate"))
			Exit Function
		End If
	Next
    
    ' no match for given ProductCode found try with empty ProductCode
	If NOT sProductCode = "" Then IsBaselineRequired = IsBaselineRequired("", sXml)

End Function 'IsBaselineRequired
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    IsMinorUpdate

'    Determine if the patch is a minor update (service pack)
'-------------------------------------------------------------------------------
Function IsMinorUpdate(sProductCode, sXml)
	Dim Element, Elements, Node, ChildNodes

	On Error Resume Next
	
    XmlDoc.LoadXml(sXml)
	Set Elements = XmlDoc.GetElementsByTagName("TargetProduct")
	For Each Element in Elements
		If sProductCode = "" Then sProductCode = Element.selectSingleNode("TargetProductCode").text
		If Element.selectSingleNode("TargetProductCode").text = sProductCode Then
			For Each Node in Element.ChildNodes
				If Node.NodeName = "UpdatedVersion" Then
					IsMinorUpdate = True
					Exit Function
				End If
			Next 'Node
		End If
	Next 'Element

End Function 'IsMinorUpdate
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    AddSequenceData
'
'    Add patch family sequence information to the dictionary
'-------------------------------------------------------------------------------
Sub AddSequenceData(sXml)
	Dim Element, Elements
	Dim sFamily, sSeq

	On Error Resume Next
	
    XmlDoc.LoadXml(sXml)
	Set Elements = XmlDoc.GetElementsByTagName("SequenceData")
	For Each Element in Elements
		sFamily = "" : sSeq = ""
		sFamily = Element.selectSingleNode("PatchFamily").text
		sSeq = Element.selectSingleNode("Sequence").text
		'Only add to the family sequence number if it's marked to supersede earlier
		If Element.selectSingleNode("Attributes").text = "1" Then
			If dicMspSequence.Exists(sFamily) Then
				If sSeq > dicMspSequence.Item(sFamily) Then dicMspSequence.Item(sFamily) = sSeq
			Else
				dicMspSequence.Add sFamily, sSeq
			End If
		End If 'Attributes = 1
	Next 'Element

End Sub 'AddSequenceData
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    WICleanOrphans
'
'    Detect and remove unreferenced .msp files from
'    %windir%\installer folder
'-------------------------------------------------------------------------------
Sub WICleanOrphans()
    Dim File, Patch, AllPatches, Product, AllProducts, oRefFilesDic
    Dim sLocalFile, sTargetFolder, sFileErr
    Dim iFoo
    Dim fFoundOrphan, fMspOK, fMsiOK

    On Error Resume Next
    
    sTmp = "Running CleanCache to remove unreferenced .msi, .msp files from " & sWICacheDir
    Log vbCrLf & vbCrLf & sTmp & vbCrLf & String(Len(sTmp), "-")
    If fCscript Then wscript.echo "Checking for unreferenced files in folder " & sWICacheDir

    fFoundOrphan = False
    fMspOK = True
    fMsiOK = True
    sTargetFolder = sTemp & "MovedCacheFiles\"

    Err.Clear
    Set oRefFilesDic = CreateObject("Scripting.Dictionary")
    If Not Err = 0 Then Exit Sub
    
    ' collect referenced .msp files
    If fCscript Then wscript.echo vbTab & "Scanning .msp files"
    For iFoo = 1 To 1
        Set AllPatches = oMsi.PatchesEx("", USERSID_EVERYONE, MSIINSTALLCONTEXT_ALL, MSIPATCHSTATE_ALL)
        If Not Err = 0 Then
            sTmp = "Error: Failed to get a complete list of all .msp files. Aborting unreferenced .msp detection."
            Log vbTab & sTmp
            LogSummary "CleanCache", vbTab & sTmp
            Log vbTab & "       Source: " & Err.Source & "; Err# (Hex): " & Hex( Err ) & _
                       "; Err# (Dec): " & Err & "; Description : " & Err.Description
            If fCscript Then wscript.echo vbTab & sTmp
            Err.Clear
            fMspOK = False
            Exit For
        End If
        
        For Each Patch in AllPatches
            sLocalFile =  ""
            sLocalFile = LCase(Patch.Patchproperty("LocalPackage"))
            If NOT oRefFilesDic.Exists(sLocalFile) Then oRefFilesDic.Add sLocalFile, sLocalFile
        Next 'Patch
        
        If Not Err = 0 Then
            sTmp = "Error: Unhandled error. Aborting unreferenced .msp detection."
            Log vbTab & sTmp
            LogSummary "CleanCache", vbTab & sTmp
            Log vbTab & "       Source: " & Err.Source & "; Err# (Hex): " & Hex( Err ) & _
                       "; Err# (Dec): " & Err & "; Description : " & Err.Description
            If fCscript Then wscript.echo vbTab & sTmp
            Err.Clear
            fMspOK = False
            Exit For
        End If
    Next 'iFoo
    
    ' collect referenced .msi files
    If fCscript Then wscript.echo vbTab & "Scanning .msi files"
    For iFoo = 1 To 1
        Set AllProducts = oMsi.ProductsEx("", USERSID_EVERYONE, MSIINSTALLCONTEXT_ALL)
        If Not Err = 0 Then
            sTmp = "Error: Failed to get a complete list of all .msi files. Aborting unreferenced .msi detection."
            Log vbTab & sTmp
            LogSummary "CleanCache", vbTab & sTmp
            Log vbTab & "       Source: " & Err.Source & "; Err# (Hex): " & Hex( Err ) & _
                       "; Err# (Dec): " & Err & "; Description : " & Err.Description
            If fCscript Then wscript.echo vbTab & sTmp
            Err.Clear
            fMsiOK = False
            Exit For
        End If
        For Each Product in AllProducts
            sLocalFile = ""
            sLocalFile = LCase(Product.InstallProperty("LocalPackage"))
            If NOT oRefFilesDic.Exists(sLocalFile) Then oRefFilesDic.Add sLocalFile, sLocalFile
        Next 'Patch
        If Not Err = 0 Then
            sTmp = "Error: Unhandled error. Aborting unreferenced .msi detection."
            Log vbTab & sTmp
            LogSummary "CleanCache", vbTab & sTmp
            Log vbTab & "       Source: " & Err.Source & "; Err# (Hex): " & Hex( Err ) & _
                       "; Err# (Dec): " & Err & "; Description : " & Err.Description
            If fCscript Then wscript.echo vbTab & sTmp
            Err.Clear
            fMsiOK = False
            Exit For
        End If
    Next 'iFoo
    
    ' move unreferenced files
    For Each File in oFso.GetFolder(sWICacheDir).Files
        If Not Err = 0 Then
            Log vbTab & "       Source: " & Err.Source & "; Err# (Hex): " & Hex( Err ) & _
                       "; Err# (Dec): " & Err & "; Description : " & Err.Description
            Select Case Err
            Case 70
                'Permission denied. Skip this file
                sTmp = "Note: File move operation failed. Skipping file " & sFileErr
                Log vbTab & sTmp
                LogSummary "CleanCache", vbTab & sTmp
                Err.Clear
            Case Else
                sTmp = "Error: Unhandled error. Aborting unreferenced file detection."
                Log vbTab & sTmp
                LogSummary "CleanCache", vbTab & sTmp
                Err.Clear
                Exit Sub
            End Select 'Err
        End If
        sFileErr = File.Name
        Select Case LCase(Right(File.Name, 4))
        Case ".msp"
            If fMspOK Then
                If NOT oRefFilesDic.Exists(LCase(File.Path)) Then
                    fFoundOrphan = True
                    If Not oFso.FolderExists(sTargetFolder) Then oFso.CreateFolder sTargetFolder
                    sTmp = "Moving unreferenced file " & File.Path & vbTab & " -> " & sTargetFolder
                    If fDetectOnly Then sTmp = "Identified unreferenced file '" & File.Path & "'. This would be moved to folder " & sTargetFolder
                    Log vbTab & "Note: " & sTmp
                    LogSummary "CleanCache", vbTab & sTmp
                    If fCscript Then wscript.echo vbTab & vbTab & sTmp
                    If NOT fDetectOnly Then
                        If oFso.FileExists(sTargetFolder & File.Name) _ 
                        Then oFso.MoveFile File.Path, sTargetFolder & sTimeStamp & "_" & File.Name _
                        Else oFso.MoveFile File.Path, sTargetFolder & File.Name
                    End If 'fDetectOnly 
                End If 'NOT oRefFilesDic.Exists
            End If 'fMspOK
        Case ".msi"
            If fMsiOK Then
                If NOT oRefFilesDic.Exists(LCase(File.Path)) Then
                    fFoundOrphan = True
                    If Not oFso.FolderExists(sTargetFolder) Then oFso.CreateFolder sTargetFolder
                    sTmp = "Moving unreferenced file " & File.Path & vbTab & " -> " & sTargetFolder
                    If fDetectOnly Then sTmp = "Identified unreferenced file '" & File.Path & "'. This would be moved to folder " & sTargetFolder 
                    Log vbTab & "Note: " & sTmp
                    LogSummary "CleanCache", vbTab & sTmp
                    If fCscript Then wscript.echo vbTab & vbTab & sTmp
                    If NOT fDetectOnly Then
                        If oFso.FileExists(sTargetFolder & File.Name) _ 
                        Then oFso.MoveFile File.Path, sTargetFolder & sTimeStamp & "_" & File.Name _
                        Else oFso.MoveFile File.Path, sTargetFolder & File.Name
                    End If 'fDetectOnly 
                End If 'NOT oRefFilesDic.Exists
            End If 'fMsiOK
        Case Else
        End Select
    Next 'File

    If NOT fFoundOrphan Then Log vbTab & "Success: No unreferenced .msp files found."
    
    ' delete the moved files in aggresive mode
    sTargetFolder = sTemp & "MovedCacheFiles"
    If fCleanAggressive Then
        sTmp = "Deleting moved files folder " & sTargetFolder
        If fDetectOnly Then sTmp = "Moved files folder would be deleted: " & sTargetFolder
        Log vbTab & "Note: " & sTmp
        LogSummary "CleanCache", vbTab & sTmp
        If NOT fDetectOnly Then oFso.DeleteFolder sTargetFolder, True
    End If 'fCleanAggressive

End Sub 'WICleanOrphans
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    MspRemove

'    Uninstall a removable patch
'-------------------------------------------------------------------------------
Sub MspRemove(sPatchCodes, sProductCodes)
    Dim Product, Patch, oPatches, Update
    Dim sPatches, sCmd, sReturn, sStateFilter, sLogFilter
    Dim iFoo, iActivityCnt
    Dim arrLogFilter
    Dim fPatchLoop, fSupersededMode, fMatchFound
    Dim sPatchCodeCompressed, sProductCodeCompressed, sUserSid, sGlobalConfigKey, sMspFile
    Dim fForceReconcile
    Dim MspDb, Record
    Dim qView

    On Error Resume Next

    iActivityCnt = 0
    If UCase(sPatchCodes) = "SUPERSEDED" Then
        sStateFilter = MSIPATCHSTATE_SUPERSEDED
        fSupersededMode = True
        sLogFilter = "superseded patches"
    Else
        sStateFilter = MSIPATCHSTATE_ALL
        fSupersededMode = False
        arrLogFilter = Split(sPatchCodes, ";")
        sLogFilter = ""
        For Each Patch in arrLogFilter
            If InStr(Patch, "{") > 0 Then sLogFilter = sLogFilter & ";" & Patch
        Next 'Patch
        Set Patch = Nothing
        If NOT sLogFilter = "" Then sLogFilter="patch(es) " & Mid(sLogFilter, 2)
    End If

    'Loop all products
    For Each Product In oMsi.Products
        If IsOfficeProduct (Product) Then
            If InStr(sProductCodes, Product)>0 OR sProductCodes = "" Then
                For iFoo = 1 To 1
                    Do 
                        fPatchLoop = False
                        fMatchFound = False
                        Log "Scanning " & Product & " - " & oMsi.ProductInfo(Product, "ProductName") & " - for " & sLogFilter
                        Set oPatches = oMsi.PatchesEx(Product, USERSID_NULL, MSIINSTALLCONTEXT_MACHINE, sStateFilter)
                        If Not Err = 0 Then
                            Err.Clear
                            Log vbCrLf & " Failed to retrieve list of patches"
                            If fCscript Then wscript.echo vbTab & " Failed to retrieve list of patches"
                            Exit For 'iFoo
                        End If
                        For Each Patch in oPatches
                            If fRemovePatchQnD Then
                                UnregisterPatch Patch
                            Else
                                If InStr(sPatchCodes, Patch.PatchCode) > 0 OR fSupersededMode Then
                                    fForceReconcile = False
                                    sMspFile = Patch.PatchProperty("LocalPackage")
                                    iActivityCnt = iActivityCnt + 1
                                    sCmd = "msiexec.exe /i " & Product & _
                                              " MSIPATCHREMOVE=" & Patch.PatchCode & _
                                              " MSIRESTARTMANAGERCONTROL=Disable" & _
                                              " REBOOT=ReallySuppress" & _
                                              " /qb-" & _
                                              " /l*v " & chr(34) & sPathOutputFolder & sComputerName & "_" & Product & "_" & Patch.PatchCode & "_MspRemove.log" & chr(34)
                                    If Patch.PatchProperty("Uninstallable") = "1" Then
                                        fMatchFound = True
                                        sTmp = "Uninstalling patch " & Patch.PatchCode & " - " & Patch.PatchProperty("DisplayName")
                                        If fDetectOnly Then
                                            sTmp = "Uninstall attempt possible to remove patch " & Patch.PatchCode & " - " & Patch.PatchProperty("DisplayName")
                                        End If 'fDetectOnly
                                        Log vbTab & "Note: " & sTmp
                                        LogSummary Product, vbTab & sTmp
                                        If fCscript Then wscript.echo vbTab & vbTab & sTmp
                                        If NOT fDetectOnly Then 
                                            Log vbTab & "Debug: calling msiexec with '" & sCmd & "'"
                                            'Execute the patch uninstall
                                            sReturn = CStr(oWShell.Run(sCmd, 0, True))
                                            fRebootRequired = fRebootRequired OR (sReturn = "3010")
                                            sTmp = "Msiexec patch removal returned: " & sReturn & " " & MsiexecRetval(sReturn)
                                            Log vbTab & "Debug:  " & sTmp
                                            LogSummary Product, vbTab & sTmp
                                            If fCscript Then wscript.echo vbTab & vbTab & sTmp
                                            If NOT (sReturn="0" OR sReturn="3010") AND oFso.FileExists(sMspFile) AND fForceRemovePatch Then fForceReconcile = True
                                        End If 'NOT fDetectOnly
                                        sPatches = sPatches & Patch.PatchCode & ";"
                                    Else
                                        If fForceRemovePatch AND NOT fDetectOnly Then
                                            fMatchFound = True
                                            sTmp = "Attmpting forced uninstall of patch " & Patch.PatchCode & " - " & Patch.PatchProperty("DisplayName")
                                            Log vbTab & "Note: " & sTmp
                                            LogSummary Product, vbTab & sTmp
                                            'A) Tweak registry flag
                                            'Fill variables
                                            sPatchCodeCompressed = GetCompressedGuid(Patch.PatchCode)
                                            sProductCodeCompressed = GetCompressedGuid(Patch.ProductCode)
                                            sUserSid = Patch.UserSid : If sUserSid = "" Then sUserSid = "S-1-5-18\" Else sUserSid = sUserSid & "\"
                                            sGlobalConfigKey = REG_GLOBALCONFIG & sUserSid & "Products\" & sProductCodeCompressed & "\Patches\" 
                                            oFso.GetFile(sMspFile).Attributes = 0
                                            If RegValExists(HKLM, sGlobalConfigKey & sPatchCodeCompressed, "Uninstallable") Then
                                                oReg.SetDWordValue HKLM, sGlobalConfigKey & sPatchCodeCompressed, "Uninstallable", 1
                                                'B) Tweak cached .msp
                                                TweakDatabase(sMspFile)
                                            'Call msiexec to uninstall patch
                                                Log vbTab & "Debug: calling msiexec with '" & sCmd & "'"
                                                'Execute the patch uninstall
                                                sReturn = CStr(oWShell.Run(sCmd, 0, True))
                                                fRebootRequired = fRebootRequired OR (sReturn = "3010")
                                                sTmp = "Msiexec patch removal returned: " & sReturn & " " & MsiexecRetval(sReturn)
                                                Log vbTab & "Debug:  " & sTmp
                                                LogSummary Product, vbTab & sTmp
                                                If fCscript Then wscript.echo vbTab & vbTab & sTmp
                                                If NOT (sReturn="0" OR sReturn="3010") Then fForceReconcile = True
                                            Else
                                                fForceReconcile = True
                                            End If
                                        Else
                                            sTmp = "Patch " & Patch.PatchCode & " - " & Patch.PatchProperty("DisplayName") & " - is not uninstallable"
                                            Log vbTab & "Note: " & sTmp
                                            LogSummary Product, vbTab & sTmp
                                            If fCscript Then wscript.echo vbTab & vbTab & sTmp
                                        End If 'fForceRemovePatch
                                    End If 'Patch Uninstallable
                            
                                    If fForceReconcile Then 
                                        sTmp = "Msiexec based uninstall not possible. Unregistering patch "
                                        sTmp = sTmp & Patch.PatchCode & " - " & Patch.PatchProperty("DisplayName")
                                        Log vbTab & "Note: " & sTmp
                                        LogSummary Product, vbTab & sTmp
                                        If oFso.FileExists(sMspFile) Then oFso.MoveFile sMspFile, sTemp & oFso.GetFileName(sMspFile)
                                        UnregisterPatch Patch
                                    End If
                                End If 'InStr
                            End If 'fRemovePatchQnD
                        Next 'Patch
                        If NOT fMatchFound Then Log vbTab & "No match found for specified patch filter"
                    Loop While fPatchLoop
                Next 'iFoo
            End If 'InStr sProductCodes
        End If 'IsOfficeProduct
    Next 'Product

    If iActivityCnt = 0 Then LogSummary "RemovePatch", vbTab & "Nothing to remove for specified patch filter" & vbCrLf

End Sub 'MspRemove
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    TweakDatabase
'
'    Create and call an external script task to tweak the cached .msp
'    This is required to work around the missing option in VBScript  to release the database handle
'-------------------------------------------------------------------------------
Sub TweakDatabase(sMspFile)
    Const MSIOPENDATABASEMODE_TRANSACT      = 1

    Dim TweakDb
    Dim sTweakCmd

    Set TweakDb = oFso.CreateTextFile(sTemp & "TweakDb.vbs", True, True)
    TweakDb.WriteLine "On Error Resume Next"
    TweakDb.WriteLine "Set Msi = CreateObject(" & chr(34) & "WindowsInstaller.Installer" & chr(34) & ")"
    TweakDb.WriteLine "Set MspDb = Msi.OpenDatabase(" & chr(34) & sMspFile & chr(34) & ", " & MSIOPENDATABASEMODE_TRANSACT + MSIOPENDATABASEMODE_PATCHFILE & ")"
    TweakDb.WriteLine "Set ModifyView = MspDb.OpenView(" & chr(34) & "UPDATE `MsiPatchMetadata` SET `MsiPatchMetadata`.`Value`='1' WHERE `MsiPatchMetadata`.`Property`='AllowRemoval'" & chr(34) & ")"
    TweakDb.WriteLine "ModifyView.Execute"
    TweakDb.WriteLine "ModifyView.Close"
    TweakDb.WriteLine "MspDb.Commit"
    TweakDb.Close

    sTweakCmd = "cscript " & chr(34) & sTemp & "TweakDb.vbs" & chr(34)
    oWShell.Run sTweakCmd, 0, True
    oFso.DeleteFile sTemp & "TweakDb.vbs", True

End Sub 'TweakDatabase
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    CabExtract
'
'    Extracts the patch embedded .cab file to the %temp% folder
'    and returns a string list of extracted cab files
'-------------------------------------------------------------------------------
Function CabExtract(sMspFile)
    Dim MspDb, Record, File, CabFile, DataSize
    Dim qView
    Dim sCabList, sCabFile

    sCabList = ""
    If NOT oFso.FileExists(sMspFile) Then
        wscript.echo "File '" & sMspFile & "' does not exist."
        Exit Function
    End If
    If NOT LCase(Right(sMspFile, 4))=".msp" Then
        wscript.echo "'" & sMspFile & "' is not a valid .msp file."
        Exit Function
    End If

    Set File = oFso.GetFile(sMspFile)
    Set MspDb = oMsi.OpenDatabase(sMspFile, MSIOPENDATABASEMODE_PATCHFILE)
    Set qView = MspDb.OpenView("SELECT * FROM _Streams") : qView.Execute
    Do
        Set Record = qView.Fetch
        If Record Is Nothing Then Exit Do
        If InStr(UCase(Record.StringData(1)), "_CAB")>0 OR InStr(UCase(Record.StringData(1)), "CUST.CAB")>0 Then
            sCabFile = "" : sCabFile = Replace(File.Name, ".msp", "") & "_" & Record.StringData(1) & ".cab"
            Set CabFile = oFso.CreateTextFile(sTemp & sCabFile)
            CabFile.Write Record.ReadStream(2, Record.DataSize(2), MSIREADSTREAM_ANSI)
            CabFile.Close
            sCabList = ";" & sTemp & sCabFile
            oWShell.Run chr(34) & sTemp & sCabFile & chr(34)
        End If
    Loop
    qView.Close

    If Len(sCabList)>0 Then sCabList = Mid(sCabList, 2)
    CabExtract = sCabList

End Function 'CabExtract
'-------------------------------------------------------------------------------
'
'    Module CollectUpdates
'
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    CollectUpdates
'-------------------------------------------------------------------------------
Sub CollectUpdates(sFilter)
    Dim sUpdatesFolder, sMspPackageName
    Dim Msp, PatchEx, PatchesEx

    Const PRODUCTCODE_EMPTY = ""

    sUpdatesFolder = oWShell.ExpandEnvironmentStrings("%TEMP%") & "\Updates"
    If Not oFso.FolderExists(sTargetFolder) Then oFso.CreateFolder sUpdatesFolder

    'Get all applied patches
    Set PatchesEx = oMsi.PatchesEx(PRODUCTCODE_EMPTY, USERSID_NULL, MSIINSTALLCONTEXT_MACHINE, MSIPATCHSTATE_APPLIED)

    On Error Resume Next
    'Enum the patches
    For Each PatchEx in PatchesEx
        If Not Err = 0 Then Err.Clear
        'Connect to the patch file
        Set Msp = oMsi.OpenDatabase(PatchEx.PatchProperty("LocalPackage"), MSIOPENDATABASEMODE_PATCHFILE)
        Set SumInfo = msp.SummaryInformation
        If Err = 0 Then
            sMspPackageName = PatchEx.PatchProperty("LocalPackage")
            If InStr(SumInfo.Property(PID_TEMPLATES), OFFICEID)>0 OR InStr(SumInfo.Property(PID_TEMPLATES), OFFICEDBGID)>0 Then
                'Get the original patch name
                Set qView = msp.OpenView("SELECT `Property`, `Value` FROM MsiPatchMetadata WHERE `Property`='StdPackageName'")
                qView.Execute : Set record = qView.Fetch()
                'Copy and rename the patch to the original filename
                oFso.CopyFile patch.PatchProperty("LocalPackage"), sTargetFolder & "\" & record.StringData(2), True
            End If
        End If 'Err = 0
    Next 'patch
    oWShell.Run "explorer /e," & chr(34) & sTargetFolder & chr(34)

End Sub 'CollectUpdates

'-------------------------------------------------------------------------------
'
'                                        Module ViewPatch
'
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    ViewPatch
'
'    Show contents of .MSP files in an Excel workbook
'    This requires XL to be installed
'-------------------------------------------------------------------------------
Sub ViewPatch (sMspFile)
    Dim XlWkbk, XlSheet, MspDb, MspFile, FileStream
    Dim Element, Elements, Node, Record, qView, CompItem
    Dim MspTarget, sPatchCode, sPatchTables, sPatchStreams, sXml, sKB, sPacklet, sSequence, ScopeItem
    Dim iXlCalculation, iXlCol, iXlRowStart, iXlRowCur, iCnt
    Dim arrSeqTable, arrMetaTable, arrMspTargets, arrMspTransforms, arrMstDetail, arrObsoleted, arrViewScope
    Dim dicTransformRow, dicInstalledPT, dicSum, dicColIndex
    Dim fOfficePatch, fOctOld, fNeedObsoleted, fNeedMstSubStorage, fSupersedesPrevious, fRet

    Const IMPNOTE = "IMPORTANT NOTE"
    Const OCTOLD = "This is a customization patch based on the original release of the OCT. A more recent OCT version is available from  http://www.microsoft.com/downloads/details.aspx?displaylang=en & FamilyID=73d955c0-da87-4bc2-bbf6-260e700519a8"

    On Error Resume Next
    
    ' ensure .msp file
    If NOT oFso.FileExists(sMspFile) Then
        wscript.echo "File '" & sMspFile & "' does not exist."
        Exit Sub
    End If
    If NOT LCase(Right(sMspFile, 4))=".msp" Then
        wscript.echo "'" & sMspFile & "' is not a valid .msp file."
        Exit Sub
    End If
    
    ' get database handle
    Set MspDb = oMsi.OpenDatabase(sMspFile, MSIOPENDATABASEMODE_PATCHFILE)
    If Not Err = 0 Then
        wscript.echo "Could not open patch " & sMspFile
        Err.Clear
        Exit Sub
    End If
    Set MspFile = oFso.GetFile(sMspFile)
    
    ' ensure defaults
    fOfficePatch = False : fNeedObsoleted = False : fNeedMstSubStorage = False : fSupersedesPrevious = False
    
    ' init Summary dic
    Set dicSum = CreateObject("Scripting.Dictionary")
    
    ' check if Excel is available
    fXl = XLInstalled
    If NOT fXl Then
        wscript.echo "ViewPatch requires Microsoft Excel to be installed"
        wscript.quit
    End If
    InitXlWkbk fXl, XlApp, XlWkbk, iXlCalculation

    'Read basic patch data
    '---------------------
    sPatchCode = Left(MspDb.SummaryInformation.Property(PID_REVNUMBER), 38)
    dicSum.Add "PatchCode", sPatchCode
    sPatchTables = GetDatabaseTables(MspDb)
    sPatchStreams = UCase(GetDatabaseStreams(MspDb))
    arrMspTargets = Split(MspDb.SummaryInformation.Property(PID_TEMPLATE), ";")
    arrMspTransforms = Split(MspDb.SummaryInformation.Property(PID_LASTAUTHOR), ";")
    
    ' determine if this is an Office patch
    For Each MspTarget in arrMspTargets
        fOfficePatch = IsOfficeProduct(MspTarget)
        If fOfficePatch Then Exit For
    Next
    
    ' init PatchXml
    InitPatchXml sMspFile, MspFile, sXml, XmlDoc

    ' init OPUtil Xml based msp output file
    InitXmlFile FileStream, MspFile.Name, sXml

    'Summary
    '--------
    ' FileName
    dicSum.Add "FileName", MspFile.Name
    
    ' KB
    sKB = GetKB(MspDb, sPatchTables)
    dicSum.Add "KB", sKB
    
    ' Packlet
    sPacklet = GetOPacklet(fOfficePatch, MspDb)
    dicSum.Add "PatchFamily", sPacklet
    
    ' prepare Sequence field
    dicSum.Add "Sequence", ""
    
    ' prepare Supersedence field
    dicSum.Add "Supersedes previous patches", "No supersedence data available"
    
    ' SummaryInformation
    AddSumInfo MspDb, dicSum, fNeedObsoleted, arrObsoleted, fNeedMstSubStorage

    ' PatchXml
    dicSum.Add "PatchXML", sTemp & MspFile.Name & "_Patch.xml"

    'PatchTargets
    '------------
    AddPatchTargetDetails XmlDoc, arrMstDetail, arrMspTransforms

    'Oct
    '---
    If InStr(sPatchStreams, "METADATA")>0 Then
    ' add OctXml link to summary
        dicSum.Add "OCT PatchXML", sTemp & MspFile.Name & "_OCT_metadata.xml"
        WriteOctXml MspDb, MspFile
        'The rtm version of the O14 OCT used TargetVersionVerification checks.
        fOctOld = (arrMstDetail(0, COL_TVV) = True)
    End If

    'MsiPatchSequence
    '----------------
    If InStr(UCase(sPatchTables), "MSIPATCHSEQUENCE")>0 Then
        AddPatchSeqDetails MspDb, dicSum, arrSeqTable, sSequence, fSupersedesPrevious, fOfficePatch
    End If

    'MsiPatchMetaData
    '----------------
    If InStr(UCase(sPatchTables), "MSIPATCHMETADATA")>0 Then
        AddPatchMetaDetails MspDb, dicSum, arrMetaTable, fOfficePatch
    End If 'MsiPatchMetaData

    'Commit data
    '-----------
    WriteSummary XlWkbk, dicSum, MspFile
    GetPatchTargets XlWkbk, dicTransformRow, arrMspTargets, arrMstDetail, fOfficePatch, TRUE
    If fNeedObsoleted Then WriteObsoleted XlWkbk, arrObsoleted
    If IsArray(arrSeqTable) Then WritePatchSeqTable XlWkbk, arrSeqTable
    If IsArray(arrMetaTable) Then WritePatchMetaTable XlWkbk, arrMetaTable

    '========================
    'Msi table update section
    '========================

    'Scan for installed PatchTargets
    '-------------------------------
    Set dicInstalledPT = CreateObject("Scripting.Dictionary")
    ScanForInstalledPT dicInstalledPT, arrMspTargets

    'Product independent view (stapled)
    '----------------------------------
    fRet = LoadTableSchema("", "", MspDb, arrMspTargets)
    LoadPatchTransforms MspDb, "", dicTransformRow, arrMspTransforms, arrMstDetail
    iXlRowStart = 1 : iXlRowCur = 1
    Set dicColIndex = CreateObject("Scripting.Dictionary")
    AddPatchTableDetails "", "", arrMspTargets, MspDb, XlWkbk, iXlCol, iXlRowStart, iXlRowCur, dicColIndex
    ClearView MspDb, sMspFile

    'Product specific view
    '---------------------
    If IsEmpty(sViewScope) Then
        If dicTransformRow.Count > 10 Then
            sViewScope = InputBox("This patch contains  " & dicTransformRow.Count & "  PatchTargets!" & vbCrLf & vbCrLf _
                        & "To filter provide parts or full ProductCode/ProductName" & vbCrLf _
                        & "or press 'Cancel' to skip this step." & vbCrLf & vbCrLf _
                        & "NOTE: An excessive list of targets like * or ALL may cause significant increase in script runtime." _
                         , " OPUtil Patch Viewer ")
        Else
            sViewScope = "*"
        End If 'dicTransformRow.Count > 10
    End If 'IsEmpty
    
    If Len(sViewScope) > 0 Then
        iCnt = 0
        arrViewScope = Split(sViewScope, ",")
        For Each MspTarget in dicTransformRow.Keys
            iCnt = iCnt + 1
            ' check scope
            For Each ScopeItem in arrViewScope
                If InStr(MspTarget, ScopeItem) > 0 OR ScopeItem = "ALL" OR ScopeItem = "*" OR InStr(UCase(GetProductID(Left(MspTarget, 38), "")), ScopeItem) > 0 Then
                    fRet = LoadTableSchema(MspTarget, "", MspDb, arrMspTargets)
                    LoadPatchTransforms MspDb, MspTarget, dicTransformRow, arrMspTransforms, arrMstDetail
                    iXlRowStart = iXlRowCur + 1 : iXlRowCur = iXlRowStart
                    XlUpdateStatus  XlWkbk, "Analyzing PatchtargetTables " & iCnt & " of " & dicTransformRow.Count
                    AddPatchTableDetails MspTarget, "", arrMspTargets, MspDb, XlWkbk, iXlCol, iXlRowStart, iXlRowCur, dicColIndex
                    ClearView MspDb, sMspFile
                    Exit For
                End If
            Next 'ScopeItem
        Next
    End If

    'Final cleanups
    XlWkbk.BuiltinDocumentProperties(1) = SOLUTIONNAME & " v" & SCRIPTBUILD
    XlWkbk.Sheets("Status").Delete
    Set XlSheet = XlWkbk.Sheets("Summary") 

    'Hand over XL control to the user
    '--------------------------------
    XlWkbk.Worksheets("Summary").Activate
    XlApp.Calculation = iXlCalculation
    XlApp.UserControl = True
    XlApp.ScreenUpdating = True
    XlApp.Interactive = True
    XlApp.DisplayAlerts = True
    XlWkbk.Saved = True
    oWShell.AppActivate XlApp.Name & " - " & XlWkBk.Name
    Set XlApp = Nothing

End Sub 'ViewPatch
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    WriteOctXml
'-------------------------------------------------------------------------------
Sub WriteOctXml(MspDb, MspFile)
    Dim Record, FileStream
    Dim qView

    Set qView = MspDb.OpenView("SELECT * FROM _Streams") : qView.Execute
    Do
        Set Record = qView.Fetch
        If Record Is Nothing Then Exit Do
        If InStr(UCase(Record.StringData(1)), "METADATA")>0 Then
            Set FileStream = oFso.CreateTextFile(sTemp & MspFile.Name & "_OCT_" & Record.StringData(1) & ".xml")
            FileStream.Write Record.ReadStream(2, Record.DataSize(2), MSIREADSTREAM_ANSI)
            FileStream.Close
        End If
    Loop
    qView.Close

End Sub 'WriteOctXml
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    InitPatchXml
'-------------------------------------------------------------------------------
Sub InitPatchXml(sMspFile, MspFile, sXml, XmlDoc)
    Dim FileStream

    sXml = oMsi.ExtractPatchXMLData(sMspFile)
    XmlDoc.LoadXml(sXml)
    Set FileStream = oFso.CreateTextFile(sTemp & MspFile.Name & "_Patch.xml", True, True)
    FileStream.Write sXml
    FileStream.Close

End Sub 'InitPatchXml
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    InitXmlFile
'-------------------------------------------------------------------------------
Sub InitXmlFile(FileStream, sMspFileName, sXml)
    Set FileStream = oFso.CreateTextFile(sTemp & sMspFileName & "_OPUtil.xml", True, True)
    sXml = Replace(sXml, "</MsiPatch>", "")
    FileStream.Write sXml

End Sub 'InitXmlFile
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    XmlCommit
'
'    Commit the temporary xml string to file
'-------------------------------------------------------------------------------
Sub XmlCommit(FileStream, sXmlLine)
    sXmlLine = Replace(sXmlLine, " >", ">")
    sXmlLine = Replace(sXmlLine, "&", "&amp;")
    FileStream.Write sXmlLine
    sXmlLine = ""

End Sub 'XmlCommit
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    XmlAddElement
'
'    Open a new element in the OPUtil xml output
'-------------------------------------------------------------------------------
Sub XmlAddElement (sXmlLine, sElement, fCompleteElement)
    sXmlLine = sXmlLine & "<" & Ucase(sElement) & " "
    If fCompleteElement Then XmlCompleteElement sXmlLine

End Sub 'XmlAddElement
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    XmlCompleteElement
'-------------------------------------------------------------------------------
Sub XmlCompleteElement (sXmlLine)
    sXmlLine = sXmlLine & ">" & vbCrLf
End Sub 'XmlCompleteElement
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    XmlCloseSingleLine
'-------------------------------------------------------------------------------
Sub XmlCloseSingleLine (sXmlLine, sElement)
    sXmlLine = sXmlLine & "</" & sElement & ">" & vbCrLf
End Sub
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    XmlCloseElement
'
'    Open a new element in the OPUtil xml output
'-------------------------------------------------------------------------------
Sub XmlCloseElement (sXmlLine, sElement)
    sXmlLine = sXmlLine & "</" & Ucase(sElement) & ">"
End Sub 'XmlCloseElement
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    XmlAddAttr
'
'    Add attribute and values to the current element
'-------------------------------------------------------------------------------
Sub XmlAddAttr (sXmlLine, sAttribute, sValue, fCloseSingleLine, sElement)
    sXmlLine = sXmlLine & sAttribute & "=" & chr(34) & sValue & chr(34) & " "
    If fCloseSingleLine Then XmlCloseSingleLine sXmlLine, sElement

End Sub 'XmlWrite
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    InitXlWkbk

'    Updates the status message in the Excel Workbook 
'    that's shown while the script execution is in progress
'-------------------------------------------------------------------------------
Sub InitXlWkbk(fXl, XlApp, XlWkbk, iXlCalculation)
    Const xlCalculationManual = &HFFFFEFD9

    Dim XlSheet
    Dim iSheetCnt

    If NOT fXl Then Exit Sub

    ' create the XL instance
    Set XlApp = CreateObject("Excel.Application")
    XlApp.DisplayAlerts = False
    
    ' avoid blank worksheets
    iSheetCnt = XlApp.SheetsInNewWorkbook
    XlApp.SheetsInNewWorkbook = 1
    Set XlWkbk = XlApp.Workbooks.Add
    XlApp.SheetsInNewWorkbook = iSheetCnt
    
    ' remember orig. XlCalculation mode then ensure manual
    iXlCalculation = XlApp.Calculation
    XlApp.Calculation = xlCalculationManual

    'Status Sheet
    '------------
    Set XlSheet = XlWkbk.Worksheets(1)
    XlSheet.Name = "Status"
    XlSheet.Cells(1, S_PROP).Value = "Please wait"
    XlSheet.Cells(1, S_VAL).Value = "Collecting Patch Details ..."
    XlSheet.Columns.Autofit
    
    ' start the user UI experience
    XlApp.Interactive = False
    XlApp.Visible = True
    XlApp.WindowState = xlMaximized
    XlApp.ScreenUpdating = False

End Sub 'InitXlWkbk
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    XlAddSheet
'
'    Add a Sheet to the WorkBook
'-------------------------------------------------------------------------------
Sub XlAddSheet(XlWkbk, XlSheet, sSheetName)
    If NOT fXl Then Exit Sub
	If NOT XlSheetExists(XlWkbk, sSheetName) Then 
		Set XlSheet = XlWkbk.Worksheets.Add
		XlMoveSheet XlWkbk, XlSheet, ""
	Else
		Set XlSheet = XlWkbk.Sheets(sSheetName)
	End If
    XlSheet.Name = sSheetName

End Sub 'XlAddSheet
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    XlSheetExists
'
'    Check if a Sheet exists in the Excel Workbook
'-------------------------------------------------------------------------------
Function XlSheetExists(XlWkbk, sSheetName)
    Dim sheet
    
    If NOT fXl Then Exit Function
    For Each sheet in XlWkbk.Worksheets
        If LCase(sSheetName) = LCase(sheet.name) Then
            XlSheetExists = True
            Exit Function
        End If
    Next 'sheet
    XlSheetExists = False

End Function 'XlSheetExists
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    XlMoveSheet
'
'    Move a Sheet in the WorkBook
'-------------------------------------------------------------------------------
Sub XlMoveSheet(XlWkbk, XlSheet, sPosAfter)
    If NOT fXl Then Exit Sub
    If sPosAfter = "" Then XlSheet.Move , XlWkbk.Sheets(XlWkbk.Sheets.Count) Else XlSheet.Move , XlWkbk.Sheets(sPosAfter)

End Sub 'XlMoveSheet
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    XlDelSheet
'
'    Delete an Excel Sheet
'-------------------------------------------------------------------------------
Sub XlDelSheet(XlSheet)
    If NOT fXl Then Exit Sub
    XlSheet.Delete

End Sub 'XlDelSheet
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    XlDelRow
'
'    Delete a row in the Excel Sheet
'-------------------------------------------------------------------------------
Sub XlDelRow(XlSheet, iRow, Direction)
    If NOT fXl Then Exit Sub
    XlSheet.Rows(iRow).Delete Direction

End Sub 'XlDelRow
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    XlWrite
'
'    Write data to Excel
'-------------------------------------------------------------------------------
Sub XlWrite(XlSheet, row, col, sValue)
    If NOT fXl Then Exit Sub
    XlSheet.Cells(row, col).Value = sValue

End Sub 'XlWrite
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    XlArrWrite
'
'    Write array data to Excel
'-------------------------------------------------------------------------------
Sub XlArrWrite(XlSheet, sRange, Arr)
    If NOT fXl Then Exit Sub
    XlSheet.Range (sRange).NumberFormat = "@"
    XlSheet.Range (sRange).Value = Arr

End Sub 'XlWrite
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    XlRead
'
'    Read data from Excel
'-------------------------------------------------------------------------------
Function XlRead(XlSheet, row, col)
    If NOT fXl Then Exit Function
    XlRead = XlSheet.Cells(row, col).Value

End Function 'XlWrite
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    XlAddHyperlink
'
'    Add a hyperlink formatting to a cell
'-------------------------------------------------------------------------------
Sub XlAddHyperlink(XlSheet, row, col, sAddress, sSubAddress)
    If NOT fXl Then Exit Sub
    XlSheet.Hyperlinks.Add XlSheet.Cells(row, col), sAddress, sSubAddress

End Sub 'XlWrite
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    XlAddListObject
'
'    Add a range as ListObject
'-------------------------------------------------------------------------------
Sub XlAddListObject(XlSheet, sListName, sRange, fToggle)
	If NOT fXl Then Exit Sub
	Dim sTableStyle
	If fToggle Then sTableStyle = "TableStyleMedium2" Else sTableStyle = "TableStyleMedium7"
	XlSheet.ListObjects.Add(xlSrcRange, XlSheet.Range(sRange), , xlYes).Name = sListName
	XlSheet.ListObjects(sListName).TableStyle = sTableStyle
    If NOT sListName = "Scope" Then XlSheet.ListObjects(sListName).Unlist

End Sub 'XlAddListObject
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    XlAddNamedRange
'
'    Add a named list
'-------------------------------------------------------------------------------
Sub XlAddNamedRange(XlSheet, sListName, sRange)
    If NOT fXl Then Exit Sub
    XlSheet.Range(sRange).Name = sListName

End Sub 'XlAddNamedList
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    XlGetRow
'
'    Get the new row from Excel
'-------------------------------------------------------------------------------
Function XlGetRow(XlSheet, iCol, iOffSet)
    If NOT fXl Then Exit Function
    If iOffSet = "" Then iOffSet = 0
    XlGetRow = XlSheet.Columns(iCol).CurrentRegion.Rows.Count + iOffSet
End Function 'XlGetRow
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    XlUpdateStatus
'
'    Updates the status message in the Excel Workbook that's shown
'    while the script execution is in progress
'-------------------------------------------------------------------------------
Sub XlUpdateStatus(XlWkbk, sStatus)
    If NOT fXl Then Exit Sub

    Dim XlSheet
    'Set the status message and refresh the screen
    Set XlSheet = XlWkbk.Sheets("Status")
    XlSheet.Activate
    XlSheet.Cells(1, S_VAL).Value = sStatus
    XlSheet.Columns.Autofit
    XlApp.ScreenUpdating = True
    XlApp.Interactive = True
    oWShell.AppActivate XlApp.Name & " - " & XlWkBk.Name
    XlApp.ScreenUpdating = False
    XlApp.Interactive = False
    
End Sub 'XlUpdateStatus
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    XlFormatContent
'
'    Format contents in XL
'-------------------------------------------------------------------------------
Sub XlFormatContent(XlSheet)
    Const S_ROW_BASELINE = 6
    Const COL_NONPOUND = 1
    Const COL_POUND = 2
    Const xlTop = &HFFFFEFC0
    
    If NOT fXl Then Exit Sub

    XlSheet.Columns.Autofit
    XlSheet.Rows.Autofit
    
    Select Case XlSheet.Name
    Case "Computer"
        XlSheet.Columns(1).VerticalAlignment = xlTop
        XlSheet.Columns(2).NumberFormat = "@"
        XlSheet.ListObjects.Add(xlSrcRange, XlSheet.UsedRange, , xlYes).Name = XlSheet.Name
        XlSheet.ListObjects(XlSheet.Name).ShowTableStyleFirstColumn = True
        XlSheet.ListObjects(XlSheet.Name).TableStyle = "TableStyleMedium10"
    Case "Summary"
        XlSheet.Columns(1).VerticalAlignment = xlTop
        XlSheet.Columns(2).NumberFormat = "@"
        XlSheet.Cells(S_ROW_BASELINE, S_PROP).NumberFormat = "0.00"
        XlSheet.ListObjects.Add(xlSrcRange, XlSheet.UsedRange, , xlYes).Name = XlSheet.Name
        XlSheet.ListObjects(XlSheet.Name).ShowTableStyleFirstColumn = True
        XlSheet.ListObjects(XlSheet.Name).TableStyle = "TableStyleMedium10"
    Case "TransformSubStorages"
        XlSheet.Columns(1).VerticalAlignment = xlTop
        XlSheet.ListObjects.Add(xlSrcRange, XlSheet.UsedRange, , xlYes).Name = XlSheet.Name
        XlSheet.ListObjects(XlSheet.Name).TableStyle = "TableStyleMedium7"
        XlSheet.Cells(HROW, COL_NONPOUND).Value = "Non Pound Transform" & vbCrLf & "(Database Diff)"
        XlSheet.Cells(HROW, COL_POUND).Value = "Pound Transform" & vbCrLf & "(Patch Specific Tables)"
    Case "ObsoletedPatches"
        XlSheet.Columns(1).VerticalAlignment = xlTop
        XlSheet.ListObjects.Add(xlSrcRange, XlSheet.UsedRange, , xlYes).Name = XlSheet.Name
        XlSheet.ListObjects(XlSheet.Name).TableStyle = "TableStyleMedium7"
    Case "MsiPatchSequence"
        XlSheet.ListObjects.Add(xlSrcRange, XlSheet.UsedRange, , xlYes).Name = XlSheet.Name
        XlSheet.ListObjects(XlSheet.Name).TableStyle = "TableStyleMedium7"
        XlSheet.Columns.Autofit
        XlSheet.Rows.Autofit
    Case "MsiPatchMetaData"
        XlSheet.ListObjects.Add(xlSrcRange, XlSheet.UsedRange, , xlYes).Name = XlSheet.Name
        XlSheet.ListObjects(XlSheet.Name).TableStyle = "TableStyleMedium7"
    Case "PatchTargets"
        XlSheet.ListObjects.Add(xlSrcRange, XlSheet.UsedRange, , xlYes).Name = XlSheet.Name
        XlSheet.ListObjects(XlSheet.Name).TableStyle = "TableStyleMedium7"
    Case "File_Table"
        iRowCnt = XlSheet.Cells(1, 1).CurrentRegion.Rows.Count
        sRange = "$A$1:" & XlSheet.Cells(iRowCnt, F_SEQUENCE).Address
        XlSheet.ListObjects.Add(xlSrcRange, XlSheet.Range(sRange), , xlYes).Name = "TableFiles"
        XlSheet.ListObjects("TableFiles").TableStyle = "TableStyleMedium2"
        If fDeepScan Then
            sRange = XLSheet.Cells(1, F_PREDICTED).Address & ":" & XLSheet.Cells(iRowCnt, F_FILEPATH).Address
            XlSheet.ListObjects.Add(xlSrcRange, XlSheet.Range(sRange), , xlYes).Name = "ThisComputerFiles"
            XlSheet.ListObjects("ThisComputerFiles").TableStyle = "TableStyleMedium7"
        End If 'fDeepScan
    Case Else
        XlSheet.ListObjects.Add(xlSrcRange, XlSheet.UsedRange, , xlYes).Name = XlSheet.Name
        XlSheet.ListObjects(XlSheet.Name).TableStyle = "TableStyleMedium2"
    End Select

End Sub 'XlFormatContent
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    AddSumInfo
'-------------------------------------------------------------------------------
Sub AddSumInfo(MspDb, dicSum, fNeedObsoleted, arrObsoleted, fNeedMstSubStorage)
    Dim i
    Dim sMinWiVer
    
    Const PID_AUTHOR                        = 4 'Author
    Const PID_COMMENTS                      = 6 'Comments
    Const PID_WORDCOUNT                     = 15 'minimum Windows Installer version 
    Const PID_SECURITY                      = 19 'read-only flag

    'Prepare the uninstallable flag
    dicSum.Add "Uninstallable", "No"
    
    For i = 1 To 19
        Select Case i
        Case PID_TITLE
            dicSum.Add "Title", MspDb.SummaryInformation.Property(PID_TITLE)
        Case PID_AUTHOR
            dicSum.Add "Author", MspDb.SummaryInformation.Property(PID_AUTHOR)
        Case PID_SUBJECT
            dicSum.Add "Subject", MspDb.SummaryInformation.Property(PID_SUBJECT)
        Case PID_COMMENTS
            dicSum.Add "Comments", MspDb.SummaryInformation.Property(PID_COMMENTS) & vbCrLf
        Case PID_REVNUMBER 'PatchCode & Obsoletion
            dicSum.Add "Obsoletes", ""
            fNeedObsoleted = (Len(MspDb.SummaryInformation.Property(PID_REVNUMBER))>LEN_GUID)
            If fNeedObsoleted Then
                dicSum.Item("Obsoletes") = "See ObsoletedPatches Sheet"
                arrObsoleted = Split(Mid(MspDb.SummaryInformation.Property(PID_REVNUMBER), LEN_GUID+1), "}")
            End If
        Case PID_TEMPLATE 'Targets
            dicSum.Add "Targets", "See PatchTargets Sheet"
        Case PID_LASTAUTHOR 'List of mst substorages
            dicSum.Add "Transform Substorages", "See TransformSubStorages Sheet"
            fNeedMstSubStorage = True
        Case PID_WORDCOUNT 'Required WI version
            Select Case MspDb.SummaryInformation.Property(PID_WORDCOUNT)
            Case 1
                sMinWiVer = "1.0 (Type 1)"
            Case 2
                sMinWiVer = "1.2 (Type 2)"
            Case 3
                sMinWiVer = "2.0 (Type 3)"
            Case 4
                sMinWiVer = "3.0 (Type 4)"
            Case 5
                sMinWiVer = "3.1 (Type 5)"
            Case Else
                sMinWiVer = MspDb.SummaryInformation.Property(PID_WORDCOUNT)
            End Select
            dicSum.Add "WI Version Required", sMinWiVer
        Case PID_SECURITY 'Read Only Flag
            dicSum.Add "Read-Only Security", ""
            Select Case MspDb.SummaryInformation.Property(PID_SECURITY)
            Case 0
                dicSum.Item("Read-Only Security") = "No Restriction"
            Case 2
                dicSum.Item("Read-Only Security") = "Read-only recommended"
            Case 4
                dicSum.Item("Read-Only Security") = "Read-only enforced"
            Case Else
                dicSum.Item("Read-Only Security") = MspDb.SummaryInformation.Property(PID_SECURITY)
            End Select
        Case Else
            'Do Not List
        End Select
    Next 'i

End Sub 'AddSumInfo
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    AddPatchTargetDetails
'-------------------------------------------------------------------------------
Sub AddPatchTargetDetails(XmlDoc, arrMstDetail, arrMspTransforms)
    Dim Element, Elements, cn
    Dim iRow

    iRow = 1
    Redim arrMstDetail (((UBound (arrMspTransforms) - 1) / 2), 12)
    Set Elements = XmlDoc.GetElementsByTagName ("TargetProduct")
    For Each Element in Elements
        iRow = iRow + 1
        arrMstDetail (iRow - 2, COL_ROW) = iRow
        arrMstDetail (iRow - 2, COL_MST) = Mid (arrMspTransforms ((iRow - 2) * 2), 2)
        For Each cn in Element.childNodes
            Select Case cn.nodeName
            Case "TargetProductCode"
                arrMstDetail(iRow - 2, COL_TPC) = cn.text
                arrMstDetail(iRow - 2, COL_TPCV) = cn.getAttribute("Validate")
            Case "TargetVersion"
                arrMstDetail(iRow-2, COL_TV) = cn.text
                arrMstDetail(iRow-2, COL_TVV) = cn.getAttribute("Validate")
                arrMstDetail(iRow-2, COL_TVCT) = cn.getAttribute("ComparisonType")
                arrMstDetail(iRow-2, COL_TVCF) = cn.getAttribute("ComparisonFilter")
            Case "UpdatedVersion"
                arrMstDetail(iRow-2, COL_UV) = cn.text
            Case "TargetLanguage"
                arrMstDetail(iRow-2, COL_TL) = cn.text
                arrMstDetail(iRow-2, COL_TLV) = cn.getAttribute("Validate")
            Case "UpgradeCode"
                arrMstDetail(iRow-2, COL_UC) = cn.text
                arrMstDetail(iRow-2, COL_UCV) = cn.getAttribute("Validate")
            Case Else
            End Select
        Next 'Node
    Next 'Element

End Sub 'AddPatchTargetDetails
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    AddPatchSeqDetails
'-------------------------------------------------------------------------------
Sub AddPatchSeqDetails(MspDb, dicSum, arrSeqTable, sSequence, fSupersedesPrevious, fOfficePatch)
    Dim Record
    Dim iRow
    Dim qView

    iRow = -1
    Set qView = MspDb.OpenView("SELECT `PatchFamily` FROM MsiPatchSequence") : qView.Execute
    Set Record = qView.Fetch()
    Do Until Record Is Nothing
        iRow = iRow + 1
        Set Record = qView.Fetch()
    Loop
    qView.Close
    
    If iRow = -1 Then Exit Sub
    fSupersedesPrevious = False
    Redim arrSeqTable(iRow, 3)
    iRow = 0
    Set qView = MspDb.OpenView("SELECT * FROM MsiPatchSequence") : qView.Execute
    Set Record = qView.Fetch()
    Do Until Record Is Nothing
        arrSeqTable(iRow, SEQ_PATCHFAMILY-1) = Record.StringData(SEQ_PATCHFAMILY)
        arrSeqTable(iRow, SEQ_PRODUCTCODE-1) = Record.StringData(SEQ_PRODUCTCODE)
        arrSeqTable(iRow, SEQ_SEQUENCE-1) = Record.StringData(SEQ_SEQUENCE)
        If NOT InStr(sSequence, Record.StringData(SEQ_SEQUENCE))>0 Then sSequence = ";" & Record.StringData(SEQ_SEQUENCE)
        arrSeqTable(iRow, SEQ_ATTRIBUTE-1) = Record.StringData(SEQ_ATTRIBUTE)
        fSupersedesPrevious = fSupersedesPrevious OR Record.StringData(SEQ_ATTRIBUTE)="1"
        Set Record = qView.Fetch()
        iRow = iRow + 1
    Loop
    qView.Close
    If fSupersedesPrevious Then dicSum.Item("Supersedes previous patches") = "Yes" Else dicSum.Item("Supersedes previous patches") = "No"
    If Len(sSequence)>1 Then sSequence = Mid(sSequence, 2)
    If InStr(sSequence, ";")>0 Then 
        dicSum.Item("Sequence") = "Multiple sequences available"
    Else
        If sSequence = "" AND fOfficePatch Then sSequence = GetLegacyMspSeq(MspDb)
        If sSequence = "" Then dicSum.Item("Sequence") = "No sequence data available" Else dicSum.Item("Sequence") = sSequence
    End If

End Sub 'AddPatchSeqDetails
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    AddPatchMetaDetails
'-------------------------------------------------------------------------------
Sub AddPatchMetaDetails(MspDb, dicSum, arrMetaTable, fOfficePatch)
    Dim Record
    Dim sPacklet
    Dim iCnt, iRow
    Dim qView

    iCnt = -1
    Set qView = MspDb.OpenView("SELECT `Property` FROM MsiPatchMetadata") : qView.Execute
    Set Record = qView.Fetch()
    Do Until Record Is Nothing
        iCnt = iCnt + 1
        Set Record = qView.Fetch()
    Loop
    qView.Close

    If iCnt = -1 Then Exit Sub
    Redim arrMetaTable(iCnt, 2)
    iRow = 0
    Set qView = MspDb.OpenView("SELECT * FROM MsiPatchMetadata") : qView.Execute
    Set Record = qView.Fetch()
    Do Until Record Is Nothing
        arrMetaTable(iRow, MET_COMPANY - 1) = Record.StringData(MET_COMPANY)
        arrMetaTable(iRow, MET_PROPERTY - 1) = Record.StringData(MET_PROPERTY)
        arrMetaTable(iRow, MET_VALUE - 1) = Record.StringData(MET_VALUE)
        If UCase(Record.StringData(MET_PROPERTY)) = "STDPACKAGENAME" AND fOfficePatch Then
            'Overwrite the Packlet information if data is available in the MetaData table
            sPacklet = ""
            sPacklet = Record.StringData(MET_VALUE)
            If InStr(sPacklet, ".")>0 Then sPacklet = Left(sPacklet, InStr(sPacklet, ".")-1)
            If NOT sPacklet = "" Then dicSum.Item("PatchFamily") = sPacklet
        End If
        If UCase(Record.StringData(MET_PROPERTY)) = "ALLOWREMOVAL" Then
            If Record.StringData(MET_VALUE) = 1 Then dicSum.Item("Uninstallable") = "Yes"
        End If
        Set Record = qView.Fetch()
        iRow = iRow + 1
    Loop
    qView.Close

End Sub 'AddPatchMetaDetails
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    AddPatchTableDetails
'-------------------------------------------------------------------------------
Sub AddPatchTableDetails(ProductCode, sTable, arrMspTargets, MspDb, XlWkbk, iXlCol, iXlRowStart, iXlRowCur, dicColIndex)
    Dim sProductCode
    Dim i, iVM
    Dim arrColHeaders, arrTable
	Dim fToggle

    iXlCol = 2
    sProductCode = ProductCode
    If InStr(ProductCode, "_") > 0 Then sProductCode = Left(ProductCode, 38)
    If sProductCode = "" Then iVM = GetVersionMajor(arrMspTargets(0)) Else iVM = GetVersionMajor(sProductCode)
	fToggle = True
    For i = 0 To UBound(arrSchema)
        If (LCase(arrSchema(i, 0)) = LCase(sTable) OR sTable = "") AND (InStr(arrSchema(i, 2), iVM) > 0 OR (Len(sProductCode) = 38 AND InStr(arrSchema(i, 2), sProductCode) > 0)) Then 
            arrColHeaders = GetTableColumnHeadersFromDef(arrSchema(i, 1))
            arrTable = GetPatchTableDetails(MspDb, arrSchema(i, 0), arrColHeaders)
			WritePatchTable XlWkbk, arrSchema(i, 0), arrColHeaders, arrTable, ProductCode, iXlCol, iXlRowStart, iXlRowCur, fToggle, dicColIndex
        End If
    Next 'i

End Sub 'AddPatchTableDetails
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    GetPatchTableDetails
'-------------------------------------------------------------------------------
Function GetPatchTableDetails(MspDb, tbl, arrColHeaders)
    Dim Record
    Dim sKey
    Dim iCnt, iRow, iCol
    Dim arrTable, arrLine
    Dim qView

    GetPatchTableDetails = arrTable
    ' defaults
    iCnt = 0 : iRow = 0 : sKey = ""
    
    ' initiate the view for counting
    On Error Resume Next
    Set qView = MspDb.OpenView("SELECT * FROM `_TransformView` WHERE `Table` = '" & tbl & "' ORDER BY `Row`")
    If NOT Err = 0 Then
        Err.Clear
        Exit Function
    End If
    On Error Goto 0
    qView.Execute()
    Set Record = qView.Fetch
    
    ' get row count
    If NOT Record Is Nothing Then
        sKey = Record.StringData(3)
        iCnt = iCnt + 1
    End If
    Do Until Record Is Nothing
        'Next Row?
        If NOT sKey = Record.StringData(3) Then 
            iCnt = iCnt + 1
            sKey = Record.StringData(3)
        End If
        Set Record = qView.Fetch
    Loop

    ' exit if patch does not modify this table
    If NOT iCnt > 0 Then Exit Function
    
    ' create array
    ReDim arrTable(iCnt - 1, UBound(arrColHeaders))
    ' collect data
    Set qView = MspDb.OpenView("SELECT * FROM `_TransformView` WHERE `Table` = '" & tbl & "' ORDER BY `Row`")
    qView.Execute()
    Set Record = qView.Fetch
    ' get the first Row
    If NOT Record Is Nothing Then
        sKey = Record.StringData(3)
        arrTable(iRow, 0) = Record.StringData(3)
    End If
    ' loop all records
    Do Until Record Is Nothing
        'Next Row?
        If NOT sKey = Record.StringData(3) Then 
            iRow = iRow + 1
            sKey = Record.StringData(3)
            arrTable(iRow, 0) = Record.StringData(3)
        End If
        'Add data from _TransformView
        Select Case Record.StringData(2)
        Case "CREATE"
        Case "DELETE"
        Case "DROP"
        Case "INSERT"
            arrLine = Split(Record.StringData(3), chr(9))
            For iCol = 0 To UBound(arrLine)
                arrTable(iRow, iCol) = arrLine(iCol)
            Next
        Case Else
            For iCol = 0 To UBound(arrColHeaders)
                If Record.StringData(2) = arrColHeaders(iCol) Then
                    arrTable(iRow, iCol) = Record.StringData(4)
                    Exit For
                End If
            Next 'iCol
        End Select
        Set Record = qView.Fetch
    Loop
    
    ' return the table contents
    GetPatchTableDetails = arrTable

End Function 'GetPatchTableDetails
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    WriteSummary
'
'    Write Contents to the output formats
'-------------------------------------------------------------------------------
Sub WriteSummary(XlWkbk, dicSum, MspFile)
    Dim iRow, iCol
    Dim sElement, sKey, sXmlLine
    Dim XlSheet
    Dim arrCacheLog

    Const S_ROW_NAME = 2
    Const S_ROW_KB = 3
    Const S_ROW_PACKLET = 4
    Const S_ROW_SEQUENCE = 5
    Const S_ROW_SUPERSEDENCE = 6
    Const S_ROW_UNINSTALLABLE = 7
    Const S_ROW_TITLE = 8
    Const S_ROW_AUTHOR = 9
    Const S_ROW_SUBJECT = 10
    Const S_ROW_COMMENTS = 11
    Const S_ROW_PATCHCODE = 12
    Const S_ROW_TARGETS = 13
    Const S_ROW_OBSOLETES = 14
    Const S_ROW_PATCHTYPE = 15
    Const S_ROW_SECURITY = 16
    Const S_ROW_PATCHXML = 17
    Const S_ROW_OCTPATCHXML = 18

    'Summary
    iRow = 1 : iCol = 1
    sElement = "Summary"
    XmlAddElement sXmlLine, sElement, False
    XlAddSheet XlWkbk, XlSheet, sElement
    XlWrite XlSheet, 1, S_PROP, "Property"
    XlWrite XlSheet, 1, S_VAL, "Value"

    sKey = "FileName"
    XlWrite XlSheet, S_ROW_NAME, S_PROP, sKey
    XlWrite XlSheet, S_ROW_NAME, S_VAL, dicSum.Item(sKey)
    XmlAddAttr sXmlLine, "FileName", MspFile.Name, False, sElement

    sKey = "KB"
    XmlAddAttr sXmlLine, "KB", dicSum.Item(sKey), False, sElement
    XlWrite XlSheet, S_ROW_KB, S_PROP, sKey
    If Len(dicSum.Item(sKey))>0 Then
        XmlAddAttr sXmlLine, "KbUrl", "http://support.microsoft.com/kb/" & dicSum.Item(sKey), False, sElement
        XlAddHyperlink XlSheet, S_ROW_KB, S_VAL, "http://support.microsoft.com/kb/" & dicSum.Item(sKey), ""
    Else
        XmlAddAttr sXmlLine, "KbUrl", "", False, sElement
    End If

    sKey = "PatchFamily"
    XlWrite XlSheet, S_ROW_PACKLET, S_PROP, sKey
    XlWrite XlSheet, S_ROW_PACKLET, S_VAL, dicSum.Item(sKey)
    XmlAddAttr sXmlLine, sKey, dicSum.Item(sKey), False, sElement

    sKey = "Sequence"
    XlWrite XlSheet, S_ROW_SEQUENCE, S_PROP, sKey
    XlWrite XlSheet, S_ROW_SEQUENCE, S_VAL, dicSum.Item(sKey)
    If InStr(dicSum.Item(sKey), "Multiple")>0 Then XlAddHyperlink XlSheet, S_ROW_SEQUENCE, S_VAL, "", "MsiPatchSequence!$A$1"

    sKey = "Supersedes previous patches"
    XlWrite XlSheet, S_ROW_SUPERSEDENCE, S_PROP, sKey
    XlWrite XlSheet, S_ROW_SUPERSEDENCE, S_VAL, dicSum.Item(sKey)

    skey = "Title"
    XlWrite XlSheet, S_ROW_TITLE, S_PROP, sKey
    XlWrite XlSheet, S_ROW_TITLE, S_VAL, dicSum.Item(sKey)

    skey = "Author"
    XlWrite XlSheet, S_ROW_AUTHOR, S_PROP, sKey
    XlWrite XlSheet, S_ROW_AUTHOR, S_VAL, dicSum.Item(sKey)

    skey = "Subject"
    XlWrite XlSheet, S_ROW_SUBJECT, S_PROP, sKey
    XlWrite XlSheet, S_ROW_SUBJECT, S_VAL, dicSum.Item(sKey)

    skey = "Comments"
    XlWrite XlSheet, S_ROW_COMMENTS, S_PROP, sKey
    XlWrite XlSheet, S_ROW_COMMENTS, S_VAL, dicSum.Item(sKey)

    skey = "PatchCode"
    XlWrite XlSheet, S_ROW_PATCHCODE, S_PROP, sKey
    XlWrite XlSheet, S_ROW_PATCHCODE, S_VAL, dicSum.Item(sKey)

    skey = "Obsoletes"
    XlWrite XlSheet, S_ROW_OBSOLETES, S_PROP, sKey
    XlWrite XlSheet, S_ROW_OBSOLETES, S_VAL, dicSum.Item(sKey)
    If NOT dicSum.Item(sKey) = "" Then XlAddHyperlink XlSheet, S_ROW_OBSOLETES, S_VAL, "", "ObsoletedPatches!$A$1"

    skey = "Targets"
    XlWrite XlSheet, S_ROW_TARGETS, S_PROP, sKey
    XlWrite XlSheet, S_ROW_TARGETS, S_VAL, dicSum.Item(sKey)
    XlAddHyperlink XlSheet, S_ROW_TARGETS, S_VAL, "", "PatchTargets!$A$1"

    skey = "WI Version Required"
    XlWrite XlSheet, S_ROW_PATCHTYPE, S_PROP, sKey
    XlWrite XlSheet, S_ROW_PATCHTYPE, S_VAL, dicSum.Item(sKey)

    skey = "Read-Only Security"
    XlWrite XlSheet, S_ROW_SECURITY, S_PROP, sKey
    XlWrite XlSheet, S_ROW_SECURITY, S_VAL, dicSum.Item(sKey)

    skey = "PatchXML"
    XlWrite XlSheet, S_ROW_PATCHXML, S_PROP, sKey
    XlWrite XlSheet, S_ROW_PATCHXML, S_VAL, dicSum.Item(sKey)
    XlAddHyperlink XlSheet, S_ROW_PATCHXML, S_VAL, dicSum.Item(sKey), ""

    skey = "OCT PatchXML"
    If NOT dicSum.Item(sKey) = "" Then 
        XlWrite XlSheet, S_ROW_OCTPATCHXML, S_PROP, sKey
        XlWrite XlSheet, S_ROW_OCTPATCHXML, S_VAL, dicSum.Item(sKey)
        XlAddHyperlink XlSheet, S_ROW_OCTPATCHXML, S_VAL, dicSum.Item(sKey), ""
    End If

    sKey = "Uninstallable"
    XlWrite XlSheet, S_ROW_UNINSTALLABLE, S_PROP, sKey
    XlWrite XlSheet, S_ROW_UNINSTALLABLE, S_VAL, dicSum.Item(sKey)

    XlFormatContent XlSheet

End Sub 'WriteSummary
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    GetPatchTargets
'
'    Obtain PatchTarget details and commit for ViewPatch mode
'-------------------------------------------------------------------------------
Sub GetPatchTargets (XlWkbk, dicTransformRow, arrMspTargets, arrMstDetail, fOfficePatch, fViewPatchMode)
    Dim XlSheet
    Dim key, prod
    Dim sValidate, sRange
    Dim i, iRow, iTargetName, iFamilyVer, iLicense, iPlatform, iNonPound, iPound
    Dim iTargetGuid, iTargetVer, iUpdatedVer, iLcid, iCulture, iVTargetProd, iVTargetVer, iVTargetLang, iVTargetUpg
    Dim arrTargetVersion, arrLogCache

    If fViewPatchMode Then
        ' create the PatchTarget sheet
        XlAddSheet XlWkbk, XlSheet, "PatchTargets"
        XlMoveSheet XlWkbk, XlSheet, "Summary"
    End If
    
    ' count loop
    dicProdMst.RemoveAll
    For Each prod in arrMspTargets
        For i = 0 To UBound(arrMstDetail)
            If (prod = arrMstDetail(i, COL_TPC)) OR (NOT arrMstDetail(i, COL_TPCV)) Then
                If NOT dicProdMst.Exists(prod & "_" & i) Then dicProdMst.Add prod & "_" & i, i
            End If
        Next 'i
    Next 'prod
    
    ' determine columns
    iTargetGuid = 0
    iTargetVer = iTargetGuid + 1
    If fOfficePatch Then
        iTargetName = iTargetGuid + 1
        iFamilyVer = iTargetName + 1
        iTargetVer = iFamilyVer + 1
    End If 'fOfficePatch
    iUpdatedVer = iTargetVer + 1
    iLcid = iUpdatedVer + 1
    iCulture = iLcid + 1
    iVTargetProd = iCulture + 1
    If fOfficePatch Then
        iLicense = iCulture + 1
        iPlatform = iLicense + 1
        iVTargetProd = iPlatform + 1
    End If
    iVTargetVer = iVTargetProd + 1
    iVTargetLang = iVTargetVer + 1
    iVTargetUpg = iVTargetLang + 1
    iNonPound = iVTargetUpg + 1
    iPound = iNonPound + 1
    
    ' init the LogCache array
    ReDim arrLogCache(dicProdMst.Count, iPound) 'iPound -> last column
    
    ' cache column headers
    arrLogCache(0, iTargetGuid) = "ProductCode"
    arrLogCache(0, iTargetVer) = "Target Baseline" & vbCrLf & "(Targeted Build)"
    arrLogCache(0, iUpdatedVer) = "Updated Baseline" & vbCrLf & "(Build Version After Patch)"
    arrLogCache(0, iLcid) = "LCID"
    arrLogCache(0, iCulture) = "Culture"
    arrLogCache(0, iVTargetProd) = "Validate ProductCode"
    arrLogCache(0, iVTargetVer) = "Validate ProductVersion"
    arrLogCache(0, iVTargetLang) = "Validate Language"
    arrLogCache(0, iVTargetUpg) = "Validate UpgradeCode"
    arrLogCache(0, iNonPound) = "Non Pound Transform" & vbCrLf & "(Database Diff)"
    arrLogCache(0, iPound) = "Pound Transform" & vbCrLf & "(Patch Specific Tables)"
    If fOfficePatch Then
        arrLogCache(0, iTargetName) = "ProductName"
        arrLogCache(0, iFamilyVer) = "Office Family"
        arrLogCache(0, iLicense) = "License"
        arrLogCache(0, iPlatform) = "Platform"
    End If
    
    ' collect PatchTarget details
    iRow = 0
    Set dicTransformRow = Nothing
    Set dicTransformRow = CreateObject("Scripting.Dictionary")
    For Each key in dicProdMst.Keys
        iRow = iRow + 1
        prod = Left(key, 38)
        i = dicProdMst.Item(key)
        
        ' ProductCode
        arrLogCache(iRow, iTargetGuid) = prod
        If arrMstDetail(i, COL_TPCV) Then  arrLogCache(iRow, iVTargetProd) = "TRUE" Else arrLogCache(iRow, iVTargetProd) = "FALSE"
        
        ' Office specifics
        If fOfficePatch Then
            ' Office ProductName
            arrLogCache(iRow, iTargetName) = GetProductID(prod, GetVersionMajor(prod))
            ' Office Family
            arrLogCache(iRow, iFamilyVer) = GetOFamilyVer(prod)
            ' Office License Channel
            arrLogCache(iRow, iLicense) = GetReleaseType(CInt(Mid(prod, 3, 1)))
            ' Architecture
            If Mid(prod, 21, 1) = "1" Then arrLogCache(iRow, iPlatform) = "x64" Else arrLogCache(iRow, iPlatform) = "x86"
            ' LCID
            Select Case GetVersionMajor(prod)
            Case 9, 10, 11
                arrMstDetail(i, COL_TL) = CInt("&h" & Mid(prod, 6, 4))
            Case Else
                arrMstDetail(i, COL_TL) = CInt("&h" & Mid(prod, 16, 4))
            End Select
        End If 'fOfficePatch
        
        ' TargetVersion
        If arrMstDetail(i, COL_TVV) Then 
            arrTargetVersion = Split(arrMstDetail(i, COL_TV), ".")
            arrLogCache(iRow, iVTargetVer) = "TRUE: " & arrMstDetail(i, COL_TVCT) & " " & GetComparisonFilter(arrMstDetail(i, COL_TVCF), arrTargetVersion) & " (" & arrMstDetail(i, COL_TVCF) & ")"
            ' add SP level
            If fOfficePatch Then arrLogCache(iRow, iTargetVer) = arrMstDetail(i, COL_TV) & GetSpLevel(arrMstDetail(i, COL_TV)) Else arrLogCache(iRow, iTargetVer) = arrMstDetail(i, COL_TV)
        Else 
            arrLogCache(iRow, iTargetVer) = "Baselineless (" & arrMstDetail(i, COL_TV) & ")"
            arrLogCache(iRow, iVTargetVer) = "FALSE"
        End If
        
        ' UpdatedVersion
        If NOT IsEmpty(arrMstDetail(i, COL_UV)) Then arrLogCache(iRow, iUpdatedVer) = arrMstDetail(i, COL_UV) Else arrLogCache(iRow, iUpdatedVer) = "Not Updated (" & arrMstDetail(i, COL_TV) & ")"
        
        ' TargetLanguage
        arrLogCache(iRow, iLcid) = arrMstDetail(i, COL_TL)
        arrLogCache(iRow, iCulture) = GetCultureInfo(arrMstDetail(i, COL_TL))
        arrLogCache(iRow, iVTargetLang) = arrMstDetail(i, COL_TLV)
        
        ' UpgradeCode
        If arrMstDetail(i, COL_UCV) Then arrLogCache(iRow, iVTargetUpg) = "TRUE: " & arrMstDetail(i, COL_UC) Else arrLogCache(iRow, iVTargetUpg) = "FALSE"
        
        ' Transforms
        arrLogCache(iRow, iNonPound) = ":" & arrMstDetail(i, COL_MST)
        arrLogCache(iRow, iPound) = ":#" & arrMstDetail(i, COL_MST)
        
        ' dicTransformRow
        If NOT dicTransformRow.Exists( prod & "_" & arrLogCache(iRow, iTargetVer)) Then
            dicTransformRow.Add prod & "_" & arrLogCache(iRow, iTargetVer), arrMstDetail(i, COL_MST)
        Else
            dicTransformRow.Item (prod & "_" & arrLogCache(iRow, iTargetVer)) = dicTransformRow.Item (prod & "_" & arrLogCache(iRow, iTargetVer)) & ";" & arrMstDetail(i, COL_MST)
        End If
    Next 'prod
    
    If fViewPatchMode Then
        ' commit the cached contents to file
        If fXL Then sRange = "$A$1:" & XlSheet.Cells(UBound(arrLogCache, 1) + 1, iPound + 1).Address
        xlArrWrite XlSheet, sRange, arrLogCache
        XlFormatContent XlSheet
    End If

End Sub 'GetPatchTargets
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    WriteObsoleted
'-------------------------------------------------------------------------------
Sub WriteObsoleted(XlWkbk, arrObsoleted)
    Dim XlSheet
    Dim i, iRow, iCol

    'Create the Obsoleted sheet
    XlAddSheet XlWkbk, XlSheet, "ObsoletedPatches"
    XlMoveSheet XlWkbk, XlSheet, "Summary"
    
    'Fill the sheet
    iRow = 1 : iCol = 1
    XlWrite XlSheet, HROW, iCol, "Obsoleted Patches"
    For i = 0 To UBound(arrObsoleted)-1
        XlWrite XlSheet, i+2, iCol, arrObsoleted(i) & "}"
    Next 'i

    XlFormatContent XlSheet

End Sub 'WriteObsoleted
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    WritePatchSeqTable
'-------------------------------------------------------------------------------
Sub WritePatchSeqTable(XlWkbk, arrSeqTable)
    Dim XlSheet, MspSeq
    Dim iRow
    
    'Create the MsiPatchSequence sheet
    XlAddSheet XlWkbk, XlSheet, "MsiPatchSequence"
    XlMoveSheet XlWkbk, XlSheet, "Summary"
    
    'Fill the sheet
    XlWrite XlSheet, HROW, SEQ_PATCHFAMILY, "PatchFamily"
    XlWrite XlSheet, HROW, SEQ_PRODUCTCODE, "ProductCode"
    XlWrite XlSheet, HROW, SEQ_SEQUENCE, "Sequence"
    XlWrite XlSheet, HROW, SEQ_ATTRIBUTE, "Attribute" & vbCrLf & "(Supersedence Flag)"
    For iRow = 0 To UBound(arrSeqTable, 1)
        XlWrite XlSheet, iRow + 2, SEQ_PATCHFAMILY, arrSeqTable(iRow, SEQ_PATCHFAMILY - 1)
        XlWrite XlSheet, iRow + 2, SEQ_PRODUCTCODE, arrSeqTable(iRow, SEQ_PRODUCTCODE - 1)
        XlWrite XlSheet, iRow + 2, SEQ_SEQUENCE, arrSeqTable(iRow, SEQ_SEQUENCE - 1)
        XlWrite XlSheet, iRow + 2, SEQ_ATTRIBUTE, arrSeqTable(iRow, SEQ_ATTRIBUTE - 1)
    Next
    XlFormatContent XlSheet

End Sub 'WritePatchSeqTable
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    WritePatchMetaTable
'-------------------------------------------------------------------------------
Sub WritePatchMetaTable(XlWkbk, arrMetaTable)
    Dim XlSheet, MetaData
    Dim iRow

    'Create the MsiPatchMetaData sheet
    XlAddSheet XlWkbk, XlSheet, "MsiPatchMetaData"
    XlMoveSheet XlWkbk, XlSheet, "Summary"
    
    'Fill the sheet
    XlWrite XlSheet, HROW, MET_COMPANY, "Company"
    XlWrite XlSheet, HROW, MET_PROPERTY, "Property"
    XlWrite XlSheet, HROW, MET_VALUE, "Value"
    For iRow = 0 To UBound(arrMetaTable, 1)
        XlWrite XlSheet, iRow + 2, MET_COMPANY, arrMetaTable(iRow, MET_COMPANY - 1)
        XlWrite XlSheet, iRow + 2, MET_PROPERTY, arrMetaTable(iRow, MET_PROPERTY - 1)
        XlWrite XlSheet, iRow + 2, MET_VALUE, arrMetaTable(iRow, MET_VALUE - 1)
    Next
    XlFormatContent XlSheet

End Sub 'WritePatchMetaTable
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    WritePatchTable
'-------------------------------------------------------------------------------
Sub WritePatchTable(XlWkbk, sTable, arrColHeaders, arrTable, ProductCode, iXlCol, iXlRowStart, iXlRowCur, fToggle, dicColIndex)
    Dim XlSheet
    Dim sSheetName, sPrefix, sRange, sScope, sProductCode
    Dim i, j, iVM, iThisCnt

    If NOT IsArray(arrTable) Then Exit Sub
    sProductCode = ProductCode
    If InStr(ProductCode, "_") > 0 Then sProductCode = Left(ProductCode, 38)
    iVM = GetVersionMajor(sProductCode)
    iXlCol = iXlCol + 1
    If dicColIndex.Exists(sTable) Then iXlCol = dicColIndex.Item(sTable) Else dicColIndex.Add sTable, iXlCol
    If iVM = 0 Then 
        sPreFix = ""
        sScope = "Combined_AllProducts_View"
    Else
        If iVM < 12 Then sPrefix = Mid(sProductCode, 4, 2) & "_" Else sPrefix = Mid(sProductCode, 11, 4)
        sScope = GetProductID(sProductCode, iVM) & "_" & ProductCode
    End If
    
    ' create the Xl sheet if needed and get the sheet handle
    sSheetName = "PatchTables"
    XlAddSheet XlWkbk, XlSheet, sSheetName
    sRange = XlSheet.Cells(iXlRowStart+1, iXlCol).Address & ":" & XlSheet.Cells(iXlRowStart+1, iXlCol + UBound(arrColHeaders) ).Address
    XlArrWrite XlSheet, sRange, arrColHeaders
    
    ' add Contents from array
    sRange = XlSheet.Cells(iXlRowStart+2, iXlCol).Address & ":" & XlSheet.Cells(iXlRowStart+2 + UBound(arrTable, 1), iXlCol + UBound(arrTable, 2)).Address
    XlArrWrite XlSheet, sRange, arrTable
    
    ' update product filter column
    iThisCnt = iXlRowStart + UBound(arrTable, 1) + 2
    If iXlRowCur < iThisCnt Then
        If iXlRowCur = 1 Then
            sRange = XlSheet.Cells(1, 1).Address & ":" & XlSheet.Cells(iXlRowCur, 1).Address
            XlWrite XlSheet, 1, 1, "Scope"
            XlAddListObject XlSheet, "Scope", sRange, fToggle
            XlSheet.Columns(2).ColumnWidth = 1
        ElseIf iXlCol = 3 Then
            XlWrite XlSheet, iXlRowCur, 1, " "
        End If
        iXlRowCur = iThisCnt
        sRange = XlSheet.Cells(iXlRowStart+1, 1).Address & ":" & XlSheet.Cells(iXlRowCur, 1).Address
        XlArrWrite XlSheet, sRange, sScope
    End If
    
    ' add a named range with formatting
    sRange = XlSheet.Cells(iXlRowStart+1, iXlCol).Address & ":" & XlSheet.Cells(iXlRowStart+2 + UBound(arrTable, 1), iXlCol + UBound(arrTable, 2)).Address
    fToggle = NOT fToggle
    XlAddListObject XlSheet, "foo", sRange, fToggle
    XlSheet.Columns.Autofit
    XlSheet.Rows.Autofit
    
    ' write header
    If iXlRowStart = 1 Then
        XlWrite XlSheet, 1, iXlCol, sTable
        sRange = XlSheet.Cells(1, iXlCol).Address & ":" & XlSheet.Cells(1, iXlCol + UBound(arrColHeaders) ).Address
        XlAddNamedRange XlSheet, sTable, sRange
    End If
    
    ' update counter
    iXlCol = iXlCol + UBound(arrTable, 2) + 1
    
    ' format columns separator
    XlSheet.Columns(iXlCol).ColumnWidth = 1

End Sub 'WritePatchTable
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    GetKB
'
'    Extract the patch KB number
'-------------------------------------------------------------------------------
Function GetKB(MspDb, sPatchTables)
    Dim Record
    Dim sKB, sTitle, sSiTmp, sChar
    Dim iSiCnt
    Dim arrSi
    Dim qView

    sKB = ""
    If InStr(sPatchTables, "MsiPatchMetadata")>0 Then
        Set qView = MspDb.OpenView("SELECT `Property`, `Value` FROM MsiPatchMetadata WHERE `Property`='KBArticle Number'")
        qView.Execute : Set Record = qView.Fetch()
        If Not Record Is Nothing Then
            sKB = UCase(Record.StringData(2))
            sKB = Replace(sKB, "KB", "")
        Else
            sKB = ""
        End If
        qView.Close
    End If
    
    If sKB = "" Then
        'Scan the SummaryInformation data for the KB
        For iSiCnt = 1 To 2
            Select Case iSiCnt
            Case 1
                arrSi = Split(MspDb.SummaryInformation.Property(PID_SUBJECT), ";")
            Case 2
                arrSi = Split(MspDb.SummaryInformation.Property(PID_TITLE), ";")
            End Select
        
            If IsArray(arrSi) Then
                For Each sTitle in arrSi
                    sSiTmp = ""
                    sSiTmp = Replace(UCase(sTitle), " ", "")
                    If InStr(sSiTmp, "KB")>0 Then
                        'Strip the KB
                        sSiTmp = Mid(sSiTmp, InStr(sSiTmp, "KB")+2)
                        For i = 1 To Len(sSiTmp)
                            sChar = ""
                            sChar = Mid(sSiTmp, i, 1)
                            If (Asc(sChar) >= 48 AND Asc(sChar) <= 57) Then sKB=sKB & sChar
                        Next 'i
                        'Ensure a valid length
                        If Len(sKB)<5 Then sKB="" Else Exit For
                    End If
                Next
                If Len(sKB)>4 Then Exit For
            End If 'IsArray(arrSi)
        Next 'iSiCnt
    End If
    GetKB = sKB

End Function 'GetKB
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    GetOPacklet
'
'    Extract the Office Patch Packlet name
'-------------------------------------------------------------------------------
Function GetOPacklet(fOfficePatch, MspDb)
    Dim arrPacklet, sPacklet

    sPacklet = ""
    If fOfficePatch Then
        arrPacklet = Split(MspDb.SummaryInformation.Property(PID_TITLE), ";")
        If IsArray(arrPacklet) Then
            If UBound(arrPacklet)>0 Then
                sPacklet = arrPacklet(1)
                If InStr(sPacklet, ".")>0 Then sPacklet = Left(sPacklet, InStr(sPacklet, ".")-1)
            End If
        End If
    End If
    GetOPacklet = sPacklet

End Function 'GetOPacklet
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'    ScanForInstalledPT
'
'    Check if PatchTargets are installed 
'-------------------------------------------------------------------------------
Sub ScanForInstalledPT(dicInstalledPT, arrMspTargets)
    Dim prod, target

    On Error Resume Next
    For Each prod  in oMsi.Products
        For Each target in arrMspTargets
            If target = prod Then
                If NOT dicInstalledPT.Exists(prod) Then
                    'Split the add to allow error handling
                    dicInstalledPT.Add prod, prod
                    dicInstalledPT.Item(prod) = oMsi.ProductInfo(prod, "LocalPackage")
                End If
                'Exit inner loop
                Exit For
            End If
        Next 'target
    Next 'prod

End Sub
'-------------------------------------------------------------------------------

'                ---------------------------------
'                Windows Installer Helper Routines
'                ---------------------------------

'-------------------------------------------------------------------------------

'Initializes msi tables schema array
Sub InitSchema(sDefExt)
    
    Dim sDef
    Dim i
    Dim arrLines, arrLine

    If NOT IsArray(arrSchema) OR NOT sDefExt = "" Then
        'Build the schema array
		'arrSchema columns: TableName - SQL - Targeted Version Major
        sDef = "_Validation;`_Validation` (`Table` CHAR(32) NOT NULL, `Column` CHAR(32) NOT NULL, `Nullable` CHAR(4) NOT NULL, `MinValue` LONG, `MaxValue` LONG, `KeyTable` CHAR(255), `KeyColumn` SHORT, `Category` CHAR(32), `Set` CHAR(255), `Description` CHAR(255) PRIMARY KEY `Table`, `Column`);14;12;11;10;9;15\AdvtExecuteSequence;`AdvtExecuteSequence` (`Action` CHAR(72) NOT NULL, `Condition` CHAR(255), `Sequence` SHORT PRIMARY KEY `Action`);14;12;11;10;9;15\Condition;`Condition` (`Feature_` CHAR(38) NOT NULL, `Level` SHORT NOT NULL, `Condition` CHAR(255) PRIMARY KEY `Feature_`, `Level`);14;12;15\AppId;`AppId` (`AppId` CHAR(38) NOT NULL, `RemoteServerName` CHAR(255), `LocalService` CHAR(255), `ServiceParameters` CHAR(255), `DllSurrogate` CHAR(255), `ActivateAtStorage` SHORT, `RunAsInteractiveUser` SHORT PRIMARY KEY `AppId`);14;12;11;10;9;15\AppSearch;`AppSearch` (`Property` CHAR(72) NOT NULL, `Signature_` CHAR(72) NOT NULL PRIMARY KEY `Property`, `Signature_`);14;12;11;10;9;15\Property;`Property` (`Property` CHAR(72) NOT NULL, `Value` LONGCHAR NOT NULL  LOCALIZABLE PRIMARY KEY `Property`);14;12;11;10;15\Binary;`Binary` (`Name` CHAR(72) NOT NULL, `Data` OBJECT NOT NULL PRIMARY KEY `Name`);14;12;11;10;9;15\Class;`Class` (`CLSID` CHAR(38) NOT NULL, `Context` CHAR(32) NOT NULL, `Component_` CHAR(72) NOT NULL, `ProgId_Default` CHAR(255), `Description` CHAR(255) LOCALIZABLE, `AppId_` CHAR(38), `FileTypeMask` CHAR(255), `Icon_` CHAR(72), `IconIndex` SHORT, `DefInprocHandler` CHAR(32), `Argument` CHAR(255), `Feature_` CHAR(38) NOT NULL, `Attributes` SHORT PRIMARY KEY `CLSID`, `Context`, `Component_`);14;12;11;15\Component;`Component` (`Component` CHAR(72) NOT NULL, `ComponentId` CHAR(38), `Directory_` CHAR(72) NOT NULL, `Attributes` SHORT NOT NULL, `Condition` CHAR(255), `KeyPath` CHAR(72) PRIMARY KEY `Component`);14;12;11;10;9;15\ProgId;`ProgId` (`ProgId` CHAR(255) NOT NULL, `ProgId_Parent` CHAR(255), `Class_` CHAR(38), `Description` CHAR(255) LOCALIZABLE, `Icon_` CHAR(72), `IconIndex` SHORT PRIMARY KEY `ProgId`);14;12;11;10;9;15\Icon;`Icon` (`Name` CHAR(72) NOT NULL, `Data` OBJECT NOT NULL PRIMARY KEY `Name`);14;12;11;10;9;15\Feature;`Feature` (`Feature` CHAR(38) NOT NULL, `Feature_Parent` CHAR(38), `Title` CHAR(64) LOCALIZABLE, `Description` CHAR(255) LOCALIZABLE, `Display` SHORT, `Level` SHORT NOT NULL, `Directory_` CHAR(72), `Attributes` SHORT NOT NULL PRIMARY KEY `Feature`);14;12;15\CompLocator;`CompLocator` (`Signature_` CHAR(72) NOT NULL, `ComponentId` CHAR(38) NOT NULL, `Type` SHORT PRIMARY KEY `Signature_`);14;12;11;10;9;15\Directory;`Directory` (`Directory` CHAR(72) NOT NULL, `Directory_Parent` CHAR(72), `DefaultDir` CHAR(255) NOT NULL LOCALIZABLE PRIMARY KEY `Directory`);14;12;11;10;9;15\CreateFolder;`CreateFolder` (`Directory_` CHAR(72) NOT NULL, `Component_` CHAR(72) NOT NULL PRIMARY KEY `Directory_`, `Component_`);14;12;11;10;9;15\CustomAction;`CustomAction` (`Action` CHAR(72) NOT NULL, `Type` SHORT NOT NULL, `Source` CHAR(72), `Target` CHAR(255), `ExtendedType` LONG PRIMARY KEY `Action`);14;15\DrLocator;`DrLocator` (`Signature_` CHAR(72) NOT NULL, `Parent` CHAR(72), `Path` CHAR(255), `Depth` SHORT PRIMARY KEY `Signature_`, `Parent`, `Path`);14;12;11;10;9;15\Error;`Error` (`Error` SHORT NOT NULL, `Message` LONGCHAR  LOCALIZABLE PRIMARY KEY `Error`);14;12;11;10;15\EventManifest;`EventManifest` (`Component_` CHAR(72) NOT NULL, `File` CHAR(72) NOT NULL PRIMARY KEY `Component_`, `File`);14;15\File;`File` (`File` CHAR(72) NOT NULL, `Component_` CHAR(72) NOT NULL, `FileName` CHAR(255) NOT NULL LOCALIZABLE, `FileSize` LONG NOT NULL, `Version` CHAR(72), `Language` CHAR(20), `Attributes` SHORT, `Sequence` LONG NOT NULL PRIMARY KEY `File`);14;12;15\Extension;`Extension` (`Extension` CHAR(255) NOT NULL, `Component_` CHAR(72) NOT NULL, `ProgId_` CHAR(255), `MIME_` CHAR(64), `Feature_` CHAR(38) NOT NULL PRIMARY KEY `Extension`, `Component_`);14;12;11;10;15\MIME;`MIME` (`ContentType` CHAR(64) NOT NULL, `Extension_` CHAR(255) NOT NULL, `CLSID` CHAR(38) PRIMARY KEY `ContentType`);14;12;11;10;9;15\FeatureComponents;`FeatureComponents` (`Feature_` CHAR(38) NOT NULL, `Component_` CHAR(72) NOT NULL PRIMARY KEY `Feature_`, `Component_`);14;12;11;10;15\Font;`Font` (`File_` CHAR(72) NOT NULL, `FontTitle` CHAR(128) PRIMARY KEY `File_`);14;12;11;10;9;15\HelpFile;`HelpFile` (`HelpFileKey` CHAR(72) NOT NULL, `HelpFileName` CHAR(72) NOT NULL, `LangID` SHORT NOT NULL, `File_HxS` CHAR(72) NOT NULL, `File_HxI` CHAR(72), `File_HxQ` CHAR(72), `File_HxR` CHAR(72), `File_Samples` CHAR(72) PRIMARY KEY `HelpFileKey`);14;12;15\HelpFileToNamespace;`HelpFileToNamespace` (`HelpFile_` CHAR(72) NOT NULL, `HelpNamespace_` CHAR(72) NOT NULL PRIMARY KEY `HelpFile_`, `HelpNamespace_`);14;12;10;15\HelpNamespace;`HelpNamespace` (`NamespaceKey` CHAR(72) NOT NULL, `NamespaceName` LONGCHAR NOT NULL, `File_Collection` CHAR(72) NOT NULL, `Description` LONGCHAR  LOCALIZABLE PRIMARY KEY `NamespaceKey`);14;12;15\InstallExecuteSequence;`InstallExecuteSequence` (`Action` CHAR(72) NOT NULL, `Condition` CHAR(255), `Sequence` SHORT PRIMARY KEY `Action`);14;12;9;15\Licenses;`Licenses` (`Name` CHAR(38) NOT NULL, `Component` CHAR(255) NOT NULL, `Data` OBJECT NOT NULL PRIMARY KEY `Name`);14;15\LicenseSetData;`LicenseSetData` (`Name` CHAR(255) NOT NULL, `Data` OBJECT NOT NULL, `Origin` CHAR(255) PRIMARY KEY `Name`);14;15\LicenseSets;`LicenseSets` (`ACID` CHAR(38) NOT NULL, `Version` CHAR(38) NOT NULL, `PL` CHAR(255), `PHN` CHAR(255), `OOB` CHAR(255) NOT NULL, `PPDLIC` CHAR(255), `RAC-Private` CHAR(255), `RAC-Public` CHAR(255), `Option` CHAR(255) PRIMARY KEY `ACID`);14;15\LockPermissions;`LockPermissions` (`LockObject` CHAR(72) NOT NULL, `Table` CHAR(32) NOT NULL, `Domain` CHAR(255), `User` CHAR(255) NOT NULL, `Permission` LONG PRIMARY KEY `LockObject`, `Table`, `Domain`, `User`);14;12;11;10;15\Media;`Media` (`DiskId` SHORT NOT NULL, `LastSequence` LONG NOT NULL, `DiskPrompt` CHAR(64) LOCALIZABLE, `Cabinet` CHAR(255), `VolumeLabel` CHAR(32), `Source` CHAR(72) PRIMARY KEY `DiskId`);14;12;15\ModuleComponents;`ModuleComponents` (`Component` CHAR(72) NOT NULL, `ModuleID` CHAR(72) NOT NULL, `Language` SHORT NOT NULL PRIMARY KEY `Component`, `ModuleID`, `Language`);14;12;15\ModuleSignature;`ModuleSignature` (`ModuleID` CHAR(72) NOT NULL, `Language` SHORT NOT NULL, `Version` CHAR(32) NOT NULL PRIMARY KEY `ModuleID`, `Language`);14;12;15\ModuleDependency;`ModuleDependency` (`ModuleID` CHAR(72) NOT NULL, `ModuleLanguage` SHORT NOT NULL, `RequiredID` CHAR(72) NOT NULL, `RequiredLanguage` SHORT NOT NULL, `RequiredVersion` CHAR(32) PRIMARY KEY `ModuleID`, `ModuleLanguage`, `RequiredID`, `RequiredLanguage`);14;12;15\ModuleExclusion;`ModuleExclusion` (`ModuleID` CHAR(72) NOT NULL, `ModuleLanguage` SHORT NOT NULL, `ExcludedID` CHAR(72) NOT NULL, `ExcludedLanguage` SHORT NOT NULL, `ExcludedMinVersion` CHAR(32), `ExcludedMaxVersion` CHAR(32) PRIMARY KEY `ModuleID`, `ModuleLanguage`, `ExcludedID`, `ExcludedLanguage`);14;12;15\MoveFile;`MoveFile` (`FileKey` CHAR(72) NOT NULL, `Component_` CHAR(72) NOT NULL, `SourceName` CHAR(255) LOCALIZABLE, `DestName` CHAR(255) LOCALIZABLE, `SourceFolder` CHAR(72), `DestFolder` CHAR(72) NOT NULL, `Options` SHORT NOT NULL PRIMARY KEY `FileKey`);14;12;11;10;9;15\MsiAssembly;`MsiAssembly` (`Component_` CHAR(72) NOT NULL, `Feature_` CHAR(38) NOT NULL, `File_Manifest` CHAR(72), `File_Application` CHAR(72), `Attributes` SHORT PRIMARY KEY `Component_`);14;12;11;15\MsiAssemblyName;`MsiAssemblyName` (`Component_` CHAR(72) NOT NULL, `Name` CHAR(255) NOT NULL, `Value` CHAR(255) NOT NULL PRIMARY KEY `Component_`, `Name`);14;12;11;15\MsiFileHash;`MsiFileHash` (`File_` CHAR(72) NOT NULL, `Options` SHORT NOT NULL, `HashPart1` LONG NOT NULL, `HashPart2` LONG NOT NULL, `HashPart3` LONG NOT NULL, `HashPart4` LONG NOT NULL PRIMARY KEY `File_`);14;12;11;10;15\MsiSFCBypass;`MsiSFCBypass` (`File_` CHAR(72) NOT NULL PRIMARY KEY `File_`);14;12;15\MsiShortcutProperty;`MsiShortcutProperty` (`MsiShortcutProperty` CHAR(72) NOT NULL, `Shortcut_` CHAR(72) NOT NULL, `PropertyKey` LONGCHAR NOT NULL, `PropVariantValue` LONGCHAR NOT NULL PRIMARY KEY `MsiShortcutProperty`);14;15\Shortcut;`Shortcut` (`Shortcut` CHAR(72) NOT NULL, `Directory_` CHAR(72) NOT NULL, `Name` CHAR(128) NOT NULL LOCALIZABLE, `Component_` CHAR(72) NOT NULL, `Target` CHAR(72) NOT NULL, `Arguments` CHAR(255), `Description` CHAR(255) LOCALIZABLE, `Hotkey` SHORT, `Icon_` CHAR(72), `IconIndex` SHORT, `ShowCmd` SHORT, `WkDir` CHAR(72), `DisplayResourceDLL` CHAR(255), `DisplayResourceId` SHORT, `DescriptionResourceDLL` CHAR(255), `DescriptionResourceId` SHORT PRIMARY KEY `Shortcut`);14;15\Shortcut;`Shortcut` (`Shortcut` CHAR(72) NOT NULL, `Directory_` CHAR(72) NOT NULL, `Name` CHAR(128) NOT NULL LOCALIZABLE, `Component_` CHAR(72) NOT NULL, `Target` CHAR(72) NOT NULL, `Arguments` CHAR(255), `Description` CHAR(255) LOCALIZABLE, `Hotkey` SHORT, `Icon_` CHAR(72), `IconIndex` SHORT, `ShowCmd` SHORT, `WkDir` CHAR(72) PRIMARY KEY `Shortcut`);12;11;10;9\NetFxNativeImage;`NetFxNativeImage` (`NetFxNativeImage` CHAR(72) NOT NULL, `File_` CHAR(72) NOT NULL, `Priority` SHORT NOT NULL, `Attributes` LONG NOT NULL, `File_Application` CHAR(72), `Directory_ApplicationBase` CHAR(72) PRIMARY KEY `NetFxNativeImage`);14;15\NiceRegistry;`NiceRegistry` (`NiceRegistry` CHAR(72) NOT NULL, `Root` SHORT NOT NULL, `Key` CHAR(255) NOT NULL LOCALIZABLE, `Name` CHAR(255) LOCALIZABLE, `Value` CHAR(255) LOCALIZABLE, `Component_` CHAR(72) NOT NULL, `TargetComponent` CHAR(72) NOT NULL, `Action` CHAR(10) NOT NULL, `IgnoreValue` CHAR(255) PRIMARY KEY `NiceRegistry`);14;15\OUpgrade;`OUpgrade` (`UpgradeCode` CHAR(38) NOT NULL, `VersionMin` CHAR(20), `VersionMax` CHAR(20), `Language` CHAR(255), `Attributes` LONG NOT NULL, `Remove` CHAR(255), `ActionProperty` CHAR(72) NOT NULL, `OPCAttributes` SHORT PRIMARY KEY `UpgradeCode`, `VersionMin`, `VersionMax`, `Language`, `Attributes`);14;12;15\PublishComponent;`PublishComponent` (`ComponentId` CHAR(38) NOT NULL, `Qualifier` CHAR(255) NOT NULL, `Component_` CHAR(72) NOT NULL, `AppData` LONGCHAR  LOCALIZABLE, `Feature_` CHAR(38) NOT NULL PRIMARY KEY `ComponentId`, `Qualifier`, `Component_`);14;12;11;10;15\Registry;`Registry` (`Registry` CHAR(72) NOT NULL, `Root` SHORT NOT NULL, `Key` CHAR(255) NOT NULL LOCALIZABLE, `Name` CHAR(255) LOCALIZABLE, `Value` LONGCHAR  LOCALIZABLE, `Component_` CHAR(72) NOT NULL PRIMARY KEY `Registry`);14;12;11;10;9;15\RegistryOnUninstall;`RegistryOnUninstall` (`RegistryOnUninstall` CHAR(72) NOT NULL, `Root` SHORT NOT NULL, `Key` CHAR(255) NOT NULL LOCALIZABLE, `Name` CHAR(255) LOCALIZABLE, `Value` CHAR(255) LOCALIZABLE, `Component_` CHAR(72) NOT NULL PRIMARY KEY `RegistryOnUninstall`);14;12;15\RegLocator;`RegLocator` (`Signature_` CHAR(72) NOT NULL, `Root` SHORT NOT NULL, `Key` CHAR(255) NOT NULL, `Name` CHAR(255), `Type` SHORT PRIMARY KEY `Signature_`);14;12;11;10;9;15\RemoveFile;`RemoveFile` (`FileKey` CHAR(72) NOT NULL, `Component_` CHAR(72) NOT NULL, `FileName` CHAR(255) LOCALIZABLE, `DirProperty` CHAR(72) NOT NULL, `InstallMode` SHORT NOT NULL PRIMARY KEY `FileKey`);14;12;11;10;9;15\RemoveRegistry;`RemoveRegistry` (`RemoveRegistry` CHAR(72) NOT NULL, `Root` SHORT NOT NULL, `Key` CHAR(255) NOT NULL LOCALIZABLE, `Name` CHAR(255) LOCALIZABLE, `Component_` CHAR(72) NOT NULL PRIMARY KEY `RemoveRegistry`);14;12;11;10;9;15\ReserveCost;`ReserveCost` (`ReserveKey` CHAR(72) NOT NULL, `Component_` CHAR(72) NOT NULL, `ReserveFolder` CHAR(72), `ReserveLocal` LONG NOT NULL, `ReserveSource` LONG NOT NULL PRIMARY KEY `ReserveKey`);14;12;11;10;9;15\SecureObjects;`SecureObjects` (`SecureObject` CHAR(72) NOT NULL, `Table` CHAR(32) NOT NULL, `Domain` CHAR(255), `User` CHAR(255) NOT NULL, `Permission` LONG, `Component_` CHAR(72) NOT NULL PRIMARY KEY `SecureObject`, `Table`, `Domain`, `User`);14;15\ServiceControl;`ServiceControl` (`ServiceControl` CHAR(72) NOT NULL, `Name` CHAR(255) NOT NULL LOCALIZABLE, `Event` SHORT NOT NULL, `Arguments` CHAR(255) LOCALIZABLE, `Wait` SHORT, `Component_` CHAR(72) NOT NULL PRIMARY KEY `ServiceControl`);14;12;11;10;9;15\ServiceInstall;`ServiceInstall` (`ServiceInstall` CHAR(72) NOT NULL, `Name` CHAR(255) NOT NULL, `DisplayName` CHAR(255) LOCALIZABLE, `ServiceType` LONG NOT NULL, `StartType` LONG NOT NULL, `ErrorControl` LONG NOT NULL, `LoadOrderGroup` CHAR(255), `Dependencies` CHAR(255), `StartName` CHAR(255), `Password` CHAR(255), `Arguments` CHAR(255), `Component_` CHAR(72) NOT NULL, `Description` CHAR(255) LOCALIZABLE PRIMARY KEY `ServiceInstall`);14;12;11;15\Signature;`Signature` (`Signature` CHAR(72) NOT NULL, `FileName` CHAR(255) NOT NULL, `MinVersion` CHAR(20), `MaxVersion` CHAR(20), `MinSize` LONG, `MaxSize` LONG, `MinDate` LONG, `MaxDate` LONG, `Languages` CHAR(255) PRIMARY KEY `Signature`);14;12;11;10;9;15\SxsMsmGenComponents;`SxsMsmGenComponents` (`Component_` CHAR(72) NOT NULL, `Guid` CHAR(38) NOT NULL PRIMARY KEY `Component_`);14;12;15\Upgrade;`Upgrade` (`UpgradeCode` CHAR(38) NOT NULL, `VersionMin` CHAR(20), `VersionMax` CHAR(20), `Language` CHAR(255), `Attributes` LONG NOT NULL, `Remove` CHAR(255), `ActionProperty` CHAR(72) NOT NULL PRIMARY KEY `UpgradeCode`, `VersionMin`, `VersionMax`, `Language`, `Attributes`);14;12;15\Verb;`Verb` (`Extension_` CHAR(255) NOT NULL, `Verb` CHAR(32) NOT NULL, `Sequence` SHORT, `Command` CHAR(255) LOCALIZABLE, `Argument` CHAR(255) LOCALIZABLE PRIMARY KEY `Extension_`, `Verb`);14;12;11;10;9;15\XEVAppsForGenericXML;`XEVAppsForGenericXML` (`ProgID` CHAR(255) NOT NULL, `Verbs` SHORT NOT NULL, `Overwrite` SHORT NOT NULL, `Priority` SHORT NOT NULL, `RegisteredCondition` CHAR(255) NOT NULL, `Component_` CHAR(72) NOT NULL PRIMARY KEY `ProgID`);14;12;11;15\XmlConfig;`XmlConfig` (`XmlConfig` CHAR(72) NOT NULL, `File` CHAR(255) NOT NULL LOCALIZABLE, `ElementPath` LONGCHAR NOT NULL  LOCALIZABLE, `VerifyPath` LONGCHAR  LOCALIZABLE, `Name` CHAR(255) LOCALIZABLE, `Value` CHAR(255) LOCALIZABLE, `Flags` LONG NOT NULL, `Component_` CHAR(72) NOT NULL, `Sequence` SHORT PRIMARY KEY `XmlConfig`);14\XmlFile;`XmlFile` (`XmlFile` CHAR(72) NOT NULL, `File` CHAR(255) NOT NULL LOCALIZABLE, `ElementPath` LONGCHAR NOT NULL  LOCALIZABLE, `Name` CHAR(255) LOCALIZABLE, `Value` LONGCHAR  LOCALIZABLE, `Flags` LONG NOT NULL, `Component_` CHAR(72) NOT NULL, `Sequence` SHORT PRIMARY KEY `XmlFile`);14;15\_sqlAction;`_sqlAction` (`CustomAction` CHAR(72) NOT NULL, `Entrypoint` CHAR(72), `Condition` CHAR(72), `Retryable` SHORT NOT NULL, `Fatal` SHORT NOT NULL, `HelpId` SHORT NOT NULL, `DisplayFlags` SHORT PRIMARY KEY `CustomAction`);14;12;15\_sqlAssembly;`_sqlAssembly` (`File_` CHAR(72) NOT NULL PRIMARY KEY `File_`);14;12\_sqlFollowComponents;`_sqlFollowComponents` (`Component_` CHAR(72) NOT NULL, `ParentComponent_` CHAR(72) NOT NULL PRIMARY KEY `Component_`, `ParentComponent_`);14;15\_sqlVerIndependentProgID;`_sqlVerIndependentProgID` (`Registry_` CHAR(72) NOT NULL PRIMARY KEY `Registry_`);14;15\ActionText;`ActionText` (`Action` CHAR(72) NOT NULL, `Description` LONGCHAR  LOCALIZABLE, `Template` LONGCHAR  LOCALIZABLE PRIMARY KEY `Action`);14;12;15\ManifestMsiResource;`ManifestMsiResource` (`File_` CHAR(72) NOT NULL PRIMARY KEY `File_`);14;12\NativeImage;`NativeImage` (`File_` CHAR(72) NOT NULL, `Attribute` SHORT NOT NULL, `Priority` SHORT NOT NULL, `CfgExecFileKey` CHAR(72), `Chip` CHAR(72) PRIMARY KEY `File_`);14\IniFile;`IniFile` (`IniFile` CHAR(72) NOT NULL, `FileName` CHAR(255) NOT NULL LOCALIZABLE, `DirProperty` CHAR(72), `Section` CHAR(96) NOT NULL LOCALIZABLE, `Key` CHAR(128) NOT NULL LOCALIZABLE, `Value` CHAR(255) NOT NULL LOCALIZABLE, `Action` SHORT NOT NULL, `Component_` CHAR(72) NOT NULL PRIMARY KEY `IniFile`);14;12;9;15\CustomAction;`CustomAction` (`Action` CHAR(72) NOT NULL, `Type` SHORT NOT NULL, `Source` CHAR(72), `Target` CHAR(255) PRIMARY KEY `Action`);12\HelpFilter;`HelpFilter` (`FilterKey` CHAR(72) NOT NULL, `Description` LONGCHAR NOT NULL, `QueryString` LONGCHAR PRIMARY KEY `FilterKey`);12;10\HelpFilterToNamespace;`HelpFilterToNamespace` (`HelpFilter_` CHAR(72) NOT NULL, `HelpNamespace_` CHAR(72) NOT NULL PRIMARY KEY `HelpFilter_`, `HelpNamespace_`);12;10\HelpPlugin;`HelpPlugin` (`HelpNamespace_` CHAR(72) NOT NULL, `HelpNamespace_Parent` CHAR(72) NOT NULL, `File_HxT` CHAR(72), `File_HxA` CHAR(72), `File_ParentHxT` CHAR(72) PRIMARY KEY `HelpNamespace_`, `HelpNamespace_Parent`);12;15\SensSubscription;`SensSubscription` (`SubscriptionID` CHAR(38) NOT NULL, `Component_` CHAR(72) NOT NULL, `SubscriptionName` CHAR(72), `PublisherID` CHAR(38), `EventClassID` CHAR(38) NOT NULL, `MethodName` CHAR(72), `SubscriberCLSID` CHAR(38) NOT NULL, `SubscriberInterface` SHORT NOT NULL, `PerUser` SHORT NOT NULL, `OwnerSID` CHAR(72), `Enabled` SHORT NOT NULL, `Description` CHAR(72), `MachineName` CHAR(72), `InterfaceID` CHAR(38) NOT NULL PRIMARY KEY `SubscriptionID`);12;11\SecureObjects;`SecureObjects` (`SecureObject` CHAR(72) NOT NULL, `Table` CHAR(32) NOT NULL, `Domain` CHAR(255), `User` CHAR(255) NOT NULL, `Permission` LONG PRIMARY KEY `SecureObject`, `Table`, `Domain`, `User`);12;11\DuplicateFile;`DuplicateFile` (`FileKey` CHAR(72) NOT NULL, `Component_` CHAR(72) NOT NULL, `File_` CHAR(72) NOT NULL, `DestName` CHAR(255) LOCALIZABLE, `DestFolder` CHAR(72) PRIMARY KEY `FileKey`);12;11;10\_sqlActionsModule;`_sqlActionsModule` (`Action` CHAR(72) NOT NULL, `ModuleID` CHAR(72) NOT NULL PRIMARY KEY `Action`);12\_sqlFollowComponents;`_sqlFollowComponents` (`FollowComponent` CHAR(72) NOT NULL, `Component_` CHAR(72) NOT NULL, `ParentComponent_` CHAR(72) NOT NULL PRIMARY KEY `FollowComponent`);12\NativeImage;`NativeImage` (`File_` CHAR(72) NOT NULL, `Attribute` SHORT, `Priority` SHORT PRIMARY KEY `File_`);12\_sqlVerIndependentProgID;`_sqlVerIndependentProgID` (`VerIndependentProgID` CHAR(72) NOT NULL, `Registry_` CHAR(72) NOT NULL PRIMARY KEY `VerIndependentProgID`);12\BindImage;`BindImage` (`File_` CHAR(72) NOT NULL, `Path` CHAR(255) PRIMARY KEY `File_`);12;11;10;9\UIText;`UIText` (`Key` CHAR(72) NOT NULL, `Text` CHAR(255) LOCALIZABLE PRIMARY KEY `Key`);12;11;10;9\AdvtUISequence;`AdvtUISequence` (`Action` CHAR(72) NOT NULL, `Condition` CHAR(255), `Sequence` SHORT PRIMARY KEY `Action`);12;10;9\Complus;`Complus` (`Component_` CHAR(72) NOT NULL, `ExpType` SHORT PRIMARY KEY `Component_`, `ExpType`);12;10\Environment;`Environment` (`Environment` CHAR(72) NOT NULL, `Name` CHAR(255) NOT NULL LOCALIZABLE, `Value` CHAR(255) LOCALIZABLE, `Component_` CHAR(72) NOT NULL PRIMARY KEY `Environment`);12;10;9;15\InitializationSequence;`InitializationSequence` (`Action` CHAR(72) NOT NULL, `Condition` CHAR(255), `Sequence` SHORT PRIMARY KEY `Action`);12\NGenExecCfg;`NGenExecCfg` (`AssemblyFileKey` CHAR(72) NOT NULL, `CfgExecFileKey` CHAR(72) NOT NULL PRIMARY KEY `AssemblyFileKey`, `CfgExecFileKey`);12\IniLocator;`IniLocator` (`Signature_` CHAR(72) NOT NULL, `FileName` CHAR(255) NOT NULL, `Section` CHAR(96) NOT NULL, `Key` CHAR(128) NOT NULL, `Field` SHORT, `Type` SHORT PRIMARY KEY `Signature_`);12;10;9\RemoveIniFile;`RemoveIniFile` (`RemoveIniFile` CHAR(72) NOT NULL, `FileName` CHAR(255) NOT NULL LOCALIZABLE, `DirProperty` CHAR(72), `Section` CHAR(96) NOT NULL LOCALIZABLE, `Key` CHAR(128) NOT NULL LOCALIZABLE, `Value` CHAR(255) LOCALIZABLE, `Action` SHORT NOT NULL, `Component_` CHAR(72) NOT NULL PRIMARY KEY `RemoveIniFile`);12;11;10;9\SelfReg;`SelfReg` (`File_` CHAR(72) NOT NULL, `Cost` SHORT PRIMARY KEY `File_`);12;11;10;9\TypeLib;`TypeLib` (`LibID` CHAR(38) NOT NULL, `Language` SHORT NOT NULL, `Component_` CHAR(72) NOT NULL, `Version` LONG, `Description` CHAR(128) LOCALIZABLE, `Directory_` CHAR(72), `Feature_` CHAR(38) NOT NULL, `Cost` LONG PRIMARY KEY `LibID`, `Language`, `Component_`);12\AssemblyPrivate;`AssemblyPrivate` (`Component_` CHAR(72) NOT NULL, `Feature_` CHAR(38) NOT NULL, `File_Manifest` CHAR(72), `File_Application` CHAR(72), `Attributes` SHORT PRIMARY KEY `Component_`);12\ManifestIntDependency;`ManifestIntDependency` (`File` CHAR(72) NOT NULL, `DependentFile` CHAR(72) NOT NULL, `Discoverable` CHAR(1), `Optional` CHAR(1) PRIMARY KEY `File`, `DependentFile`);12\ManifestExtDependency;`ManifestExtDependency` (`File` CHAR(72) NOT NULL, `DependentFileName` CHAR(255) NOT NULL, `DestinationPath` CHAR(255), `SourcePath` CHAR(255), `Discoverable` CHAR(1), `Optional` CHAR(1) PRIMARY KEY `File`, `DependentFileName`);12\ManifestMSIComponents;`ManifestMSIComponents` (`Component` CHAR(72) NOT NULL, `ManComp` CHAR(255) NOT NULL PRIMARY KEY `Component`);12\LaunchCondition;`LaunchCondition` (`Condition` CHAR(255) NOT NULL, `Description` CHAR(255) NOT NULL LOCALIZABLE PRIMARY KEY `Condition`);12;14;15\UninstallRemoveRegistry;`UninstallRemoveRegistry` (`RemoveRegistry` CHAR(72) NOT NULL, `Root` SHORT NOT NULL, `Key` CHAR(255) NOT NULL LOCALIZABLE, `Name` CHAR(255) LOCALIZABLE, `Component_` CHAR(72) NOT NULL PRIMARY KEY `RemoveRegistry`);12;11;10\NetFxNativeImage;`NetFxNativeImage` (`File_` CHAR(72) NOT NULL, `NetFxNativeImage` CHAR(72) NOT NULL, `Priority` SHORT NOT NULL, `Attributes` SHORT NOT NULL, `File_Application` CHAR(72), `Directory_ApplicationBase` CHAR(72) PRIMARY KEY `File_`);12\ODBCDataSource;`ODBCDataSource` (`DataSource` CHAR(72) NOT NULL, `Component_` CHAR(72) NOT NULL, `Description` CHAR(255) NOT NULL, `DriverDescription` CHAR(255) NOT NULL, `Registration` SHORT NOT NULL PRIMARY KEY `DataSource`);12\ODBCSourceAttribute;`ODBCSourceAttribute` (`DataSource_` CHAR(72) NOT NULL, `Attribute` CHAR(32) NOT NULL, `Value` CHAR(255) LOCALIZABLE PRIMARY KEY `DataSource_`, `Attribute`);12;9\ActionText;`ActionText` (`Action` CHAR(72) NOT NULL, `Description` CHAR(64) LOCALIZABLE, `Template` CHAR(128) LOCALIZABLE PRIMARY KEY `Action`);11;10;9\CCPSearch;`CCPSearch` (`Signature_` CHAR(72) NOT NULL PRIMARY KEY `Signature_`);11;10;9\AdminExecuteSequence;`AdminExecuteSequence` (`Action` CHAR(72) NOT NULL, `Condition` CHAR(255), `Sequence` SHORT PRIMARY KEY `Action`);11;10;9;14;15\Condition;`Condition` (`Feature_` CHAR(38) NOT NULL, `Level` SHORT NOT NULL, `Condition` LONGCHAR PRIMARY KEY `Feature_`, `Level`);11;10\AdminUISequence;`AdminUISequence` (`Action` CHAR(72) NOT NULL, `Condition` CHAR(255), `Sequence` SHORT PRIMARY KEY `Action`);11;10;9;14;15\CheckBox;`CheckBox` (`Property` CHAR(72) NOT NULL, `Value` CHAR(64) PRIMARY KEY `Property`);11;10;9;14;15\Control;`Control` (`Dialog_` CHAR(72) NOT NULL, `Control` CHAR(50) NOT NULL, `Type` CHAR(20) NOT NULL, `X` SHORT NOT NULL, `Y` SHORT NOT NULL, `Width` SHORT NOT NULL, `Height` SHORT NOT NULL, `Attributes` LONG, `Property` CHAR(50), `Text` LONGCHAR  LOCALIZABLE, `Control_Next` CHAR(50), `Help` CHAR(50) LOCALIZABLE PRIMARY KEY `Dialog_`, `Control`);11;10;9\ListBox;`ListBox` (`Property` CHAR(72) NOT NULL, `Order` SHORT NOT NULL, `Value` CHAR(64) NOT NULL, `Text` CHAR(64) LOCALIZABLE PRIMARY KEY `Property`, `Order`);11;10;9;14;15\ControlCondition;`ControlCondition` (`Dialog_` CHAR(72) NOT NULL, `Control_` CHAR(50) NOT NULL, `Action` CHAR(50) NOT NULL, `Condition` LONGCHAR NOT NULL PRIMARY KEY `Dialog_`, `Control_`, `Action`, `Condition`);11;10\ControlEvent;`ControlEvent` (`Dialog_` CHAR(72) NOT NULL, `Control_` CHAR(50) NOT NULL, `Event` CHAR(50) NOT NULL, `Argument` CHAR(255) NOT NULL, `Condition` LONGCHAR, `Ordering` SHORT PRIMARY KEY `Dialog_`, `Control_`, `Event`, `Argument`, `Condition`);11;10\CustomAction;`CustomAction` (`Action` CHAR(72) NOT NULL, `Type` SHORT NOT NULL, `Source` CHAR(64), `Target` LONGCHAR PRIMARY KEY `Action`);11;10\RegLookup;`RegLookup` (`Registry` LONGCHAR NOT NULL, `Root` SHORT, `Key` LONGCHAR NOT NULL, `Name` LONGCHAR, `Property` LONGCHAR NOT NULL PRIMARY KEY `Registry`);11;10\FeatureDependency;`FeatureDependency` (`Feature_` CHAR(38) NOT NULL, `Feature_Lead` CHAR(38) NOT NULL, `Attributes` LONG, `Sequence` SHORT PRIMARY KEY `Feature_`, `Feature_Lead`);11;10\Dialog;`Dialog` (`Dialog` CHAR(72) NOT NULL, `HCentering` SHORT NOT NULL, `VCentering` SHORT NOT NULL, `Width` SHORT NOT NULL, `Height` SHORT NOT NULL, `Attributes` LONG, `Title` CHAR(128) LOCALIZABLE, `Control_First` CHAR(50) NOT NULL, `Control_Default` CHAR(50), `Control_Cancel` CHAR(50), `Condition` LONGCHAR, `Ordering` SHORT, `WizardAttributes` LONG PRIMARY KEY `Dialog`);11\EventMapping;`EventMapping` (`Dialog_` CHAR(72) NOT NULL, `Control_` CHAR(50) NOT NULL, `Event` CHAR(50) NOT NULL, `Attribute` CHAR(50) NOT NULL PRIMARY KEY `Dialog_`, `Control_`, `Event`);11;10;9;14;15\Feature;`Feature` (`Feature` CHAR(38) NOT NULL, `Feature_Parent` CHAR(38), `Title` CHAR(128) LOCALIZABLE, `Description` CHAR(255) LOCALIZABLE, `Display` SHORT, `Level` SHORT NOT NULL, `Directory_` CHAR(72), `Attributes` SHORT NOT NULL PRIMARY KEY `Feature`);11;10\FeatureCabinets;`FeatureCabinets` (`Feature_` LONGCHAR NOT NULL, `Cabinet_` LONGCHAR NOT NULL PRIMARY KEY `Feature_`, `Cabinet_`);11\CabinetDetail;`CabinetDetail` (`Cabinet` CHAR(255) NOT NULL, `Size` LONG NOT NULL, `MD5` CHAR(32) NOT NULL, `Attributes` SHORT NOT NULL PRIMARY KEY `Cabinet`);11\File;`File` (`File` CHAR(72) NOT NULL, `Component_` CHAR(72) NOT NULL, `FileName` CHAR(255) NOT NULL LOCALIZABLE, `FileSize` LONG NOT NULL, `Version` CHAR(72), `Language` CHAR(20), `Attributes` SHORT, `Sequence` SHORT NOT NULL PRIMARY KEY `File`);11;10;9\InstallExecuteSequence;`InstallExecuteSequence` (`Action` CHAR(72) NOT NULL, `Condition` LONGCHAR, `Sequence` SHORT PRIMARY KEY `Action`);11;10\InstallUISequence;`InstallUISequence` (`Action` CHAR(72) NOT NULL, `Condition` LONGCHAR, `Sequence` SHORT PRIMARY KEY `Action`);11;10\LaunchCondition;`LaunchCondition` (`Condition` CHAR(255) NOT NULL, `Description` LONGCHAR NOT NULL  LOCALIZABLE PRIMARY KEY `Condition`);11;10\Media;`Media` (`DiskId` SHORT NOT NULL, `LastSequence` SHORT NOT NULL, `DiskPrompt` CHAR(64) LOCALIZABLE, `Cabinet` CHAR(255), `VolumeLabel` CHAR(32), `Source` CHAR(32) PRIMARY KEY `DiskId`);11;10;9\ModuleComponents;`ModuleComponents` (`Component` CHAR(72) NOT NULL, `ModuleID` CHAR(144) NOT NULL, `Language` SHORT NOT NULL PRIMARY KEY `Component`, `ModuleID`, `Language`);11;10\ModuleDependency;`ModuleDependency` (`ModuleID` CHAR(144) NOT NULL, `ModuleLanguage` SHORT NOT NULL, `RequiredID` CHAR(72) NOT NULL, `RequiredLanguage` SHORT NOT NULL, `RequiredVersion` CHAR(32) PRIMARY KEY `ModuleID`, `ModuleLanguage`, `RequiredID`, `RequiredLanguage`);11;10\ModuleSignature;`ModuleSignature` (`ModuleID` CHAR(144) NOT NULL, `Language` SHORT NOT NULL, `Version` CHAR(32) NOT NULL PRIMARY KEY `ModuleID`, `Language`);11;10\ODBCDataSource;`ODBCDataSource` (`DataSource` CHAR(72) NOT NULL, `Component_` CHAR(72) NOT NULL, `Description` CHAR(255) NOT NULL LOCALIZABLE, `DriverDescription` CHAR(255) NOT NULL LOCALIZABLE, `Registration` SHORT NOT NULL PRIMARY KEY `DataSource`);11;10;9\RadioButton;`RadioButton` (`Property` CHAR(72) NOT NULL, `Order` SHORT NOT NULL, `Value` CHAR(64) NOT NULL, `X` SHORT NOT NULL, `Y` SHORT NOT NULL, `Width` SHORT NOT NULL, `Height` SHORT NOT NULL, `Text` LONGCHAR  LOCALIZABLE, `Help` CHAR(50) LOCALIZABLE PRIMARY KEY `Property`, `Order`);11;10\Upgrade;`Upgrade` (`UpgradeCode` LONGCHAR NOT NULL, `VersionMin` LONGCHAR, `VersionMax` LONGCHAR, `Language` LONGCHAR, `Attributes` SHORT NOT NULL, `Remove` LONGCHAR, `ActionProperty` LONGCHAR, `OPCAttributes` SHORT PRIMARY KEY `UpgradeCode`, `VersionMin`, `VersionMax`, `Language`, `Attributes`);11;10\TypeLib;`TypeLib` (`LibID` CHAR(38) NOT NULL, `Language` SHORT NOT NULL, `Component_` CHAR(72) NOT NULL, `Version` SHORT, `Description` CHAR(128) LOCALIZABLE, `Directory_` CHAR(72), `Feature_` CHAR(38) NOT NULL, `Cost` LONG PRIMARY KEY `LibID`, `Language`, `Component_`);11;10\Registry2;`Registry2` (`Registry2` CHAR(72) NOT NULL, `Root` SHORT NOT NULL, `Key` CHAR(255) LOCALIZABLE, `Name` CHAR(255) LOCALIZABLE, `Value` CHAR(255) LOCALIZABLE, `Component_` CHAR(72) NOT NULL PRIMARY KEY `Registry2`);11\RegistryOnUninstall;`RegistryOnUninstall` (`RegistryOnUninstall` CHAR(72) NOT NULL, `Root` SHORT NOT NULL, `Key` CHAR(255) LOCALIZABLE, `Name` CHAR(255) LOCALIZABLE, `Value` CHAR(255) LOCALIZABLE, `Component_` CHAR(72) NOT NULL PRIMARY KEY `RegistryOnUninstall`);11\TextStyle;`TextStyle` (`TextStyle` CHAR(72) NOT NULL, `FaceName` CHAR(32) NOT NULL, `Size` SHORT NOT NULL, `Color` LONG, `StyleBits` SHORT PRIMARY KEY `TextStyle`);11;10;9;14;15\IniFile;`IniFile` (`IniFile` CHAR(72) NOT NULL, `FileName` CHAR(255) NOT NULL, `DirProperty` CHAR(72), `Section` CHAR(96) NOT NULL, `Key` CHAR(128) NOT NULL, `Value` CHAR(255) NOT NULL LOCALIZABLE, `Action` SHORT NOT NULL, `Component_` CHAR(72) NOT NULL PRIMARY KEY `IniFile`);11;10\BBControl;`BBControl` (`Billboard_` CHAR(50) NOT NULL, `BBControl` CHAR(50) NOT NULL, `Type` CHAR(50) NOT NULL, `X` SHORT NOT NULL, `Y` SHORT NOT NULL, `Width` SHORT NOT NULL, `Height` SHORT NOT NULL, `Attributes` LONG, `Text` LONGCHAR  LOCALIZABLE PRIMARY KEY `Billboard_`, `BBControl`);10\Billboard;`Billboard` (`Billboard` CHAR(50) NOT NULL, `Feature_` CHAR(38) NOT NULL, `Action` CHAR(50), `Ordering` SHORT PRIMARY KEY `Billboard`);10\Class;`Class` (`CLSID` CHAR(38) NOT NULL, `Context` CHAR(32) NOT NULL, `Component_` CHAR(72) NOT NULL, `ProgId_Default` CHAR(255), `Description` CHAR(255) LOCALIZABLE, `AppId_` CHAR(38), `FileTypeMask` CHAR(255), `Icon_` CHAR(72), `IconIndex` SHORT, `DefInprocHandler` CHAR(32), `Argument` CHAR(255), `Feature_` CHAR(38) NOT NULL PRIMARY KEY `CLSID`, `Context`, `Component_`);10\Dialog;`Dialog` (`Dialog` CHAR(72) NOT NULL, `HCentering` SHORT NOT NULL, `VCentering` SHORT NOT NULL, `Width` SHORT NOT NULL, `Height` SHORT NOT NULL, `Attributes` LONG, `Title` CHAR(128) LOCALIZABLE, `Control_First` CHAR(50) NOT NULL, `Control_Default` CHAR(50), `Control_Cancel` CHAR(50) PRIMARY KEY `Dialog`);10;9;14;15\HelpFile;`HelpFile` (`HelpFileKey` CHAR(72) NOT NULL, `HelpFileName` CHAR(72) NOT NULL, `LangID` SHORT, `File_HxS` CHAR(72) NOT NULL, `File_HxI` CHAR(72), `File_HxQ` CHAR(72), `File_HxR` CHAR(72), `Component_Samples` CHAR(72) PRIMARY KEY `HelpFileKey`);10\HelpNamespace;`HelpNamespace` (`NamespaceKey` CHAR(72) NOT NULL, `NamespaceName` LONGCHAR NOT NULL, `File_Collection` CHAR(72) NOT NULL, `Description` LONGCHAR PRIMARY KEY `NamespaceKey`);10\HelpPlugin;`HelpPlugin` (`HelpNamespace_` CHAR(72) NOT NULL, `HelpNamespace_Parent` CHAR(72) NOT NULL, `File_HxT` CHAR(72), `File_HxA` CHAR(72), `File_ParentHxt` CHAR(72) PRIMARY KEY `HelpNamespace_`, `HelpNamespace_Parent`);10\HHContent;`HHContent` (`Content_Key` CHAR(72) NOT NULL, `Name` LONGCHAR NOT NULL, `Hxs_` CHAR(72) NOT NULL, `LangID` LONGCHAR NOT NULL, `HXI_` CHAR(72), `HXQ_` CHAR(72), `SAMPLEDIRECTORY_` CHAR(72) PRIMARY KEY `Content_Key`);10\HHFilter;`HHFilter` (`Key` CHAR(72) NOT NULL, `Description` LONGCHAR, `QueryString` LONGCHAR PRIMARY KEY `Key`);10\HHNameSpace;`HHNameSpace` (`Namespace_Key` CHAR(72) NOT NULL, `Namespace_Name` LONGCHAR NOT NULL, `Collection_` CHAR(72) NOT NULL, `Description` LONGCHAR NOT NULL, `Hxt_` CHAR(72), `Hxa_` CHAR(72), `ParentNS` LONGCHAR PRIMARY KEY `Namespace_Key`);10\HHNameSpaceToFile;`HHNameSpaceToFile` (`HHContentKey` CHAR(72) NOT NULL, `HHNameSpaceKey` CHAR(72) NOT NULL PRIMARY KEY `HHContentKey`, `HHNameSpaceKey`);10\HHNameSpaceToFilter;`HHNameSpaceToFilter` (`HHNameSpaceKey` CHAR(72) NOT NULL, `HHFilterKey` CHAR(72) NOT NULL PRIMARY KEY `HHNameSpaceKey`, `HHFilterKey`);10\ModuleExclusion;`ModuleExclusion` (`ModuleID` CHAR(144) NOT NULL, `ModuleLanguage` SHORT NOT NULL, `ExcludedID` CHAR(72) NOT NULL, `ExcludedLanguage` SHORT NOT NULL, `ExcludedMinVersion` CHAR(32), `ExcludedMaxVersion` CHAR(32) PRIMARY KEY `ModuleID`, `ModuleLanguage`, `ExcludedID`, `ExcludedLanguage`);10\ServiceInstall;`ServiceInstall` (`ServiceInstall` CHAR(72) NOT NULL, `Name` CHAR(255) NOT NULL, `DisplayName` CHAR(255) LOCALIZABLE, `ServiceType` LONG NOT NULL, `StartType` LONG NOT NULL, `ErrorControl` LONG NOT NULL, `LoadOrderGroup` CHAR(255), `Dependencies` CHAR(255), `StartName` CHAR(255), `Password` CHAR(255), `Arguments` CHAR(255), `Component_` CHAR(72) NOT NULL PRIMARY KEY `ServiceInstall`);10;9\Condition;`Condition` (`Feature_` CHAR(32) NOT NULL, `Level` SHORT NOT NULL, `Condition` CHAR(255) PRIMARY KEY `Feature_`, `Level`);9\Property;`Property` (`Property` CHAR(72) NOT NULL, `Value` CHAR(128) NOT NULL LOCALIZABLE PRIMARY KEY `Property`);9\BBControl;`BBControl` (`Billboard_` CHAR(50) NOT NULL, `BBControl` CHAR(50) NOT NULL, `Type` CHAR(50) NOT NULL, `X` SHORT NOT NULL, `Y` SHORT NOT NULL, `Width` SHORT NOT NULL, `Height` SHORT NOT NULL, `Attributes` LONG, `Text` CHAR(50) LOCALIZABLE PRIMARY KEY `Billboard_`, `BBControl`);9\Billboard;`Billboard` (`Billboard` CHAR(50) NOT NULL, `Feature_` CHAR(32) NOT NULL, `Action` CHAR(50), `Ordering` SHORT PRIMARY KEY `Billboard`);9\Feature;`Feature` (`Feature` CHAR(32) NOT NULL, `Feature_Parent` CHAR(32), `Title` CHAR(64) LOCALIZABLE, `Description` CHAR(255) LOCALIZABLE, `Display` SHORT, `Level` SHORT NOT NULL, `Directory_` CHAR(72), `Attributes` SHORT NOT NULL PRIMARY KEY `Feature`);9\Class;`Class` (`CLSID` CHAR(38) NOT NULL, `Context` CHAR(32) NOT NULL, `Component_` CHAR(72) NOT NULL, `ProgId_Default` CHAR(255), `Description` CHAR(255) LOCALIZABLE, `AppId_` CHAR(38), `FileTypeMask` CHAR(255), `Icon_` CHAR(72), `IconIndex` SHORT, `DefInprocHandler` CHAR(32), `Argument` CHAR(255), `Feature_` CHAR(32) NOT NULL PRIMARY KEY `CLSID`, `Context`, `Component_`);9\ComboBox;`ComboBox` (`Property` CHAR(72) NOT NULL, `Order` SHORT NOT NULL, `Value` CHAR(64) NOT NULL, `Text` CHAR(64) LOCALIZABLE PRIMARY KEY `Property`, `Order`);9\ControlCondition;`ControlCondition` (`Dialog_` CHAR(72) NOT NULL, `Control_` CHAR(50) NOT NULL, `Action` CHAR(50) NOT NULL, `Condition` CHAR(255) NOT NULL PRIMARY KEY `Dialog_`, `Control_`, `Action`, `Condition`);9;14;15\ControlEvent;`ControlEvent` (`Dialog_` CHAR(72) NOT NULL, `Control_` CHAR(50) NOT NULL, `Event` CHAR(50) NOT NULL, `Argument` CHAR(255) NOT NULL, `Condition` CHAR(255), `Ordering` SHORT PRIMARY KEY `Dialog_`, `Control_`, `Event`, `Argument`, `Condition`);9;14;15\CustomAction;`CustomAction` (`Action` CHAR(72) NOT NULL, `Type` SHORT NOT NULL, `Source` CHAR(64), `Target` CHAR(255) PRIMARY KEY `Action`);9\DuplicateFile;`DuplicateFile` (`FileKey` CHAR(72) NOT NULL, `Component_` CHAR(72) NOT NULL, `File_` CHAR(72) NOT NULL, `DestName` CHAR(255) LOCALIZABLE, `DestFolder` CHAR(32) PRIMARY KEY `FileKey`);9\Error;`Error` (`Error` SHORT NOT NULL, `Message` CHAR(255) LOCALIZABLE PRIMARY KEY `Error`);9\Extension;`Extension` (`Extension` CHAR(255) NOT NULL, `Component_` CHAR(72) NOT NULL, `ProgId_` CHAR(255), `MIME_` CHAR(64), `Feature_` CHAR(32) NOT NULL PRIMARY KEY `Extension`, `Component_`);9\FeatureComponents;`FeatureComponents` (`Feature_` CHAR(32) NOT NULL, `Component_` CHAR(72) NOT NULL PRIMARY KEY `Feature_`, `Component_`);9\InstallUISequence;`InstallUISequence` (`Action` CHAR(72) NOT NULL, `Condition` CHAR(255), `Sequence` SHORT PRIMARY KEY `Action`);9;14;15\LaunchCondition;`LaunchCondition` (`Condition` CHAR(255) NOT NULL, `Description` CHAR(255) LOCALIZABLE PRIMARY KEY `Condition`);9\ListView;`ListView` (`Property` CHAR(72) NOT NULL, `Order` SHORT NOT NULL, `Value` CHAR(64) NOT NULL, `Text` CHAR(64) LOCALIZABLE, `Binary_` CHAR(72) PRIMARY KEY `Property`, `Order`);9\ODBCAttribute;`ODBCAttribute` (`Driver_` CHAR(72) NOT NULL, `Attribute` CHAR(40) NOT NULL, `Value` CHAR(255) LOCALIZABLE PRIMARY KEY `Driver_`, `Attribute`);9\ODBCDriver;`ODBCDriver` (`Driver` CHAR(72) NOT NULL, `Component_` CHAR(72) NOT NULL, `Description` CHAR(255) NOT NULL LOCALIZABLE, `File_` CHAR(72) NOT NULL, `File_Setup` CHAR(72) PRIMARY KEY `Driver`);9\ODBCTranslator;`ODBCTranslator` (`Translator` CHAR(72) NOT NULL, `Component_` CHAR(72) NOT NULL, `Description` CHAR(255) NOT NULL LOCALIZABLE, `File_` CHAR(72) NOT NULL, `File_Setup` CHAR(72) PRIMARY KEY `Translator`);9\Patch;`Patch` (`File_` CHAR(72) NOT NULL, `Sequence` SHORT NOT NULL, `PatchSize` LONG NOT NULL, `Attributes` SHORT NOT NULL, `Header` OBJECT NOT NULL PRIMARY KEY `File_`, `Sequence`);9\PatchPackage;`PatchPackage` (`PatchId` CHAR(38) NOT NULL, `Media_` SHORT NOT NULL PRIMARY KEY `PatchId`);9\PublishComponent;`PublishComponent` (`ComponentId` CHAR(38) NOT NULL, `Qualifier` CHAR(255) NOT NULL, `Component_` CHAR(72) NOT NULL, `AppData` CHAR(255) LOCALIZABLE, `Feature_` CHAR(32) NOT NULL PRIMARY KEY `ComponentId`, `Qualifier`, `Component_`);9\RadioButton;`RadioButton` (`Property` CHAR(72) NOT NULL, `Order` SHORT NOT NULL, `Value` CHAR(64) NOT NULL, `X` SHORT NOT NULL, `Y` SHORT NOT NULL, `Width` SHORT NOT NULL, `Height` SHORT NOT NULL, `Text` CHAR(64) LOCALIZABLE, `Help` CHAR(50) LOCALIZABLE PRIMARY KEY `Property`, `Order`);9\RegLookup;`RegLookup` (`Registry` CHAR(32) NOT NULL, `Root` SHORT, `Key` CHAR(255) NOT NULL, `Name` CHAR(255), `Property` CHAR(32) NOT NULL PRIMARY KEY `Registry`);9\TypeLib;`TypeLib` (`LibID` CHAR(38) NOT NULL, `Language` SHORT NOT NULL, `Component_` CHAR(72) NOT NULL, `Version` SHORT, `Description` CHAR(128) LOCALIZABLE, `Directory_` CHAR(72), `Feature_` CHAR(32) NOT NULL, `Cost` LONG PRIMARY KEY `LibID`, `Language`, `Component_`);9\Control;`Control` (`Dialog_` CHAR(72) NOT NULL, `Control` CHAR(50) NOT NULL, `Type` CHAR(20) NOT NULL, `X` SHORT NOT NULL, `Y` SHORT NOT NULL, `Width` SHORT NOT NULL, `Height` SHORT NOT NULL, `Attributes` LONG, `Property` CHAR(72), `Text` LONGCHAR  LOCALIZABLE, `Control_Next` CHAR(50), `Help` CHAR(50) LOCALIZABLE PRIMARY KEY `Dialog_`, `Control`);14;15\ServiceConfig;`ServiceConfig` (`ServiceName` CHAR(72) NOT NULL, `Component_` CHAR(72) NOT NULL, `NewService` SHORT NOT NULL, `FirstFailureActionType` CHAR(32) NOT NULL, `SecondFailureActionType` CHAR(32) NOT NULL, `ThirdFailureActionType` CHAR(32) NOT NULL, `ResetPeriodInDays` LONG, `RestartServiceDelayInSeconds` LONG, `ProgramCommandLine` CHAR(255), `RebootMessage` CHAR(255) PRIMARY KEY `ServiceName`);14;15\PerfmonManifest;`PerfmonManifest` (`Component_` CHAR(72) NOT NULL, `File` CHAR(72) NOT NULL, `ResourceFileDirectory` CHAR(255) NOT NULL PRIMARY KEY `Component_`, `File`, `ResourceFileDirectory`);14;15\SoftwareIdentificationTag;`SoftwareIdentificationTag` (`File_` CHAR(72) NOT NULL, `Regid` LONGCHAR NOT NULL, `UniqueId` LONGCHAR NOT NULL, `Type` LONGCHAR NOT NULL PRIMARY KEY `File_`);15\XmlConfig;`XmlConfig` (`XmlConfig` CHAR(72) NOT NULL, `File` CHAR(255) NOT NULL LOCALIZABLE, `ElementPath` LONGCHAR NOT NULL  LOCALIZABLE, `VerifyPath` LONGCHAR  LOCALIZABLE, `Name` CHAR(255) LOCALIZABLE, `Value` LONGCHAR  LOCALIZABLE, `Flags` LONG NOT NULL, `Component_` CHAR(72) NOT NULL, `Sequence` SHORT PRIMARY KEY `XmlConfig`);15"
        sDef = sDef & sDefExt
        arrLines = Split(sDef, "\")
        ReDim arrSchema(UBound(arrLines), 2)
        For i = 0 To UBound(arrLines)
            arrLine = Split(arrLines(i), ";", 3)
            arrSchema(i, 0) = arrLine(0)
            arrSchema(i, 1) = arrLine(1)
            arrSchema(i, 2) = arrLine(2)
        Next
    End If 'arrSchema

End Sub 'InitSchema
'-------------------------------------------------------------------------------

'Load a msi table schema definition
Function LoadTableSchema(ProductCode, sTables, MspDb, arrMspTargets)
    
    Dim sProductCode
    Dim i

' preset as result to False
    LoadTableSchema = False
' handle ProductCode
    sProductCode = ProductCode
    If sProductCode = "" Then
        For i = 0 To UBound(arrMspTargets)
            If LoadKnownTableSchema(arrMspTargets(i), sTables, "", MspDb) Then
                LoadTableSchema = True
                Exit Function
            End If
        Next 'i
    ElseIf InStr(sProductCode, "_") > 0 Then 
        sProductCode = Left(sProductCode, 38)
    End If

    LoadTableSchema = True
	'First attempt is to obtain from known definitions
    If NOT LoadKnownTableSchema(sProductCode, sTables, "", MspDb) Then
        'Second attempt is to check if a schema file has been provided
        If NOT MsiProvidedAsFile(MspDb, sProductCode, sTables, sMspFile, arrMspTargets) Then
            'Third attempt is to look for a match with installed products
            If NOT LoadSchemaFromInstalled (MspDb, sProductCode, sTables, arrMspTargets) Then
                'Last option is to use a generic definition by defaulting to O14 ProPlus
                'NOTE: A generic schema can produce wrong results in the report!
                LoadTableSchema = NOT LoadKnownTableSchema("{90140000-0011-0000-0000-0000000FF1CE}", sTables, "", MspDb)
            End If 'LoadSchemaFromInstalled
        End If 'MsiProvidedAsFile
    End If 'LoadKnownTableSchema
End Function 'LoadTableSchema
'-------------------------------------------------------------------------------

Function LoadKnownTableSchema(sProductCode, sTables, sDefExt, MspDb)
    
    Dim i, iVM

    'Default to False
    LoadKnownTableSchema = False
    
    If NOT IsArray(arrSchema) OR NOT sDefExt = "" Then InitSchema sDefExt
    iVM = GetVersionMajor(sProductCode)
    For i = 0 To UBound(arrSchema)
        If (LCase(arrSchema(i, 0)) = LCase(sTables) OR sTables = "") AND (InStr(arrSchema(i, 2), iVM) > 0 OR (Len(sProductCode) = 38 AND InStr(arrSchema(i, 2), sProductCode) > 0)) Then 
			MspDb.OpenView("CREATE TABLE " & arrSchema(i, 1)).Execute 
            LoadKnownTableSchema = True
        End If
    Next
End Function 'LoadKnownTableSchema
'-------------------------------------------------------------------------------

'Clear the View by renewal of the database handle
Sub ClearView (MspDb, sMspFile)
    Set MspDb = Nothing
    Set MspDb = oMsi.OpenDatabase(sMspFile, MSIOPENDATABASEMODE_PATCHFILE)
End Sub 'ClearView
'-------------------------------------------------------------------------------

Function MsiProvidedAsFile(MspDb, sProductCode, sTables, fsoFile, arrMspTargets)

Dim f, mspTarget, Prod, MsiDb, tbl
Dim sDef, sDefExt
Dim arrTables
Dim i
Dim fKnownDef

MsiProvidedAsFile = False
sDefExt = ""

'Check if we have a valid .msi file in the .msp folder location
For Each f in fsoFile.ParentFolder.Files
    If Right(LCase(f), 4)=".msi" Then
        Prod = GetMsiProductCode(f.Path)
        For Each mspTarget in arrMspTargets
            If Prod = mspTarget Then
                MsiProvidedAsFile = True
                fMsiProvidedAsFile = True
                Set MsiDb = Msi.OpenDatabase(f.Path, msiOpenDatabaseModeReadOnly)
                arrTables = Split(GetDatabaseTables(MsiDb), ",")
                For Each tbl in arrTables
                    sDef = "" : fKnownDef = False
                    sDef = "`" & tbl & "` (" & GetTableColumnDef(MsiDb, tbl) & " PRIMARY KEY " & GetPrimaryTableKeys(MsiDb, tbl) & ")"
                    For i = 0 To UBound(arrSchema)
                        If arrSchema(i, 1) = sDef Then
                            fKnownDef = True
                            If NOT InStr(arrSchema(i, 2), Prod) > 0 Then arrSchema(i, 2) = arrSchema(i, 2) & ";" & Prod
                        End If
                    Next 'i
                    If NOT fKnownDef Then sDefExt = sDefExt & "\" & sDef
                Next 'tbl
                If NOT sDefExt = "" Then LoadKnownTableSchema sProductCode, sTables, sDefExt, MspDb
                Exit Function
            End If 'Prod
        Next 'mspTarget
    End If '.msi
Next 'f

End Function 'MsiProvidedAsFile
'-------------------------------------------------------------------------------

'Try to dynamically obtain the schema from installed products
Function LoadSchemaFromInstalled (MspDb, sProductCode, sTables, arrMspTargets)

Dim Prod, MsiDb, mspTarget, tbl
Dim sDef, sDefExt
Dim arrTables
Dim i
Dim fKnownDef

LoadSchemaFromInstalled = False
For Each Prod in oMsi.Products
    For Each mspTarget in arrMspTargets
        If Prod = mspTarget Then
            LoadSchemaFromInstalled = True
            Set MsiDb = oMsi.OpenDatabase(oMsi.ProductInfo(Prod, "LocalPackage"), MSIOPENDATABASEMODE_READONLY)
            arrTables = Split(GetDatabaseTables(MsiDb), ",")
            For Each tbl in arrTables
                sDef = "" : fKnownDef = False
                sDef = "`" & tbl & "` (" & GetTableColumnDef(MsiDb, tbl) & " PRIMARY KEY " & GetPrimaryTableKeys(MsiDb, tbl) & ")"
                For i = 0 To UBound(arrSchema)
                    If arrSchema(i, 1) = sDef Then
                        fKnownDef = True
                        If NOT InStr(arrSchema(i, 2), Prod) > 0 Then arrSchema(i, 2) = arrSchema(i, 2) & ";" & Prod
                    End If
                Next 'i
                If NOT fKnownDef Then sDefExt = sDefExt & "\" & sDef
            Next 'tbl
            If NOT sDefExt = "" Then LoadKnownTableSchema sProductCode, sTables, sDefExt, MspDb
            Exit Function
        End If 'Prod = mspTarget
    Next 'mspTarget
Next 'Prod

End Function 'LoadSchemaFromInstalled
'-------------------------------------------------------------------------------

Sub LoadPatchTransforms(MspDb, ProductCode, dicTransformRow, arrMspTransforms, arrMstDetail)
    Const MSITRANSFORMERROR_ALL = 319
    
    Dim mst, sProductCode
    Dim i
    Dim arrMst

    On Error Resume Next
    ' handle ProductCode
    sProductCode = ProductCode
    If sProductCode = "" Then
        'Load all transforms
        For Each mst in arrMspTransforms
            MspDb.ApplyTransform mst, MSITRANSFORMERROR_ALL
        Next 'mst
    Else
        ' only load matching transforms
        sProductCode = Left(ProductCode, 38)
        If dicTransformRow.Exists(ProductCode) Then
            arrMst = Split(dicTransformRow.Item(ProductCode), ";")
            For i = 0  To UBound(arrMst)
                MspDb.ApplyTransform ":" & arrMst(i), MSITRANSFORMERROR_ALL
                MspDb.ApplyTransform ":#" & arrMst(i), MSITRANSFORMERROR_ALL
            Next 'i
        Else
            For i = 0 To UBound(arrMstDetail)
                If (sProductCode = arrMstDetail(i, COL_TPC)) OR (NOT arrMstDetail(i, COL_TPCV)) Then
                    MspDb.ApplyTransform ":" & arrMstDetail(i, COL_MST), MSITRANSFORMERROR_ALL
                    MspDb.ApplyTransform ":#" & arrMstDetail(i, COL_MST), MSITRANSFORMERROR_ALL
                End If
            Next 'i
        End If
    End If

End Sub 'LoadPatchTransforms
'-------------------------------------------------------------------------------

'Scans for a dynamic (optimized) SUPdateLocation folder structure
'Builds a global dictionary with identified folders
Sub DiscoverDynSUpdateFolders
    Dim Product
    Dim sRelPath, sCulture

    Set dicDynCultFolders = CreateObject ("Scripting.Dictionary")
    For Each Product in oMsi.Products
        If Len (Product) = 38 Then
            If IsOfficeProduct (Product) Then
                sRelPath = ""
                sRelPath = GetVersionMajor (Product) & ".0"
                sCulture = LCase (GetCultureInfo (Product))
                Select Case sRelPath
                Case "12.0", "14.0"
                    If Mid (Product, 11, 1) = "1" Then sRelPath = sRelPath & "\Server" Else sRelPath = sRelPath & "\Client"
                    If Mid (Product, 21, 1) = "1" Then sRelPath = sRelPath & "\x64" Else sRelPath = sRelPath & "\x86"
                End Select
                If sCulture = "neutral" Then sCulture = "x-none"
                sRelPath = sRelPath & "\" & sCulture
                If NOT sCulture = "" Then
                    If NOT dicDynCultFolders.Exists (sCulture) Then dicDynCultFolders.Add sCulture, sCulture
                End If
            End If 'IsOfficeProduct
        End If '38
    Next 'Product

    fDynSUpdateDiscovered = True
End Sub 'DiscoverSUpdateFolders
'-------------------------------------------------------------------------------

'Returns a boolean to determine if Excel is installed on the computer
Function XLInstalled()
    Dim Product

    XLInstalled = False
    For Each Product In oMsi.Products
        If oMsi.FeatureState(Product, "EXCELFiles") = MSIINSTALLSTATE_LOCAL Then
            XLInstalled = True : Exit For
        End If
    Next 'Product
End Function
'-------------------------------------------------------------------------------

'Returns the Msi file hash values as comma separated string list
Function GetMsiFileHash(sFullFileName)
    Dim Record
    
    On Error Resume Next
    
    GetMsiFileHash = ""
    Set Record = oMsi.FileHash(sFullFileName, 0)
    GetMsiFileHash = Record.StringData(1) & ", " & Record.StringData(2) & ", " & Record.StringData(3) & ", " & Record.StringData(4)
    
End Function
'-------------------------------------------------------------------------------

'Obtain the ProductCode (GUID) from a .msi package
'The function will open the .msi database and query the 'Property' table to retrieve the ProductCode
Function GetMsiProductCode(sMsiFile)
    
    Dim MsiDb, Record
    Dim qView
    
    On Error Resume Next
    
    GetMsiProductCode = ""
    Set Record = Nothing
    
    Set MsiDb = oMsi.OpenDatabase(sMsiFile, MSIOPENDATABASEMODE_READONLY)
    Set qView = MsiDb.OpenView("SELECT `Value` FROM Property WHERE `Property` = 'ProductCode'")
    qView.Execute
    Set Record = qView.Fetch
    GetMsiProductCode = Record.StringData(1)
    qView.Close

End Function 'GetMsiProductCode
'-------------------------------------------------------------------------------

'Obtain the ProductVersion from a .msi package
'The function will open the .msi database and query the 'Property' table to retrieve the ProductCode
Function GetMsiProductVersion(sMsiFile)
    
    Dim MsiDb, Record
    Dim qView
    
    On Error Resume Next
    
    GetMsiProductVersion = ""
    Set Record = Nothing
    
    Set MsiDb = oMsi.OpenDatabase(sMsiFile, MSIOPENDATABASEMODE_READONLY)
    Set qView = MsiDb.OpenView("SELECT `Value` FROM Property WHERE `Property` = 'ProductVersion'")
    qView.Execute
    Set Record = qView.Fetch
    If NOT Record Is Nothing Then GetMsiProductVersion = Record.StringData(1)
    qView.Close

End Function 'GetMsiProductVersion
'-------------------------------------------------------------------------------

'Obtain the PackageCode (GUID) from a .msi package
'The function will the .msi'S SummaryInformation stream
Function GetMsiPackageCode(sMsiFile)

    On Error Resume Next
    
    GetMsiPackageCode = ""
    GetMsiPackageCode = oMsi.SummaryInformation(sMsiFile, MSIOPENDATABASEMODE_READONLY).Property(PID_REVNUMBER)

End Function 'GetMsiPackageCode
'-------------------------------------------------------------------------------

'Returns a string with the patch sequence data
Function GetLegacyMspSeq(Msp)

Dim i
Dim sSeq
Dim arrTitle

sSeq = ""
arrTitle = Split(Msp.SummaryInformation.Property(PID_TITLE), ";")
If IsArray(arrTitle) Then
    If UBound(arrTitle)>1 Then
        sSeq = arrTitle(2)
        For i = 1 To Len(sSeq)
            If NOT (Asc(Mid(sSeq, i, 1)) >= 48 AND Asc(Mid(sSeq, i, 1)) <= 57) Then 
                sSeq = ""
                Exit For
            End If
        Next 'i
    End If
End If

GetLegacyMspSeq = sSeq

End Function 'GetMspSequence
'-------------------------------------------------------------------------------

'Detect the real product build number based on .msi and .msp build information
'to allow verification of the registered build number
Function GetRealBuildVersion(sAppliedPatches, sProductCode)

    Dim Element, Elements
    
    Dim sProductVersionReg, sProductVersionMsi

    Dim iIndex

    On Error Resume Next

    sProductVersionReg  = oMsi.ProductInfo(sProductCode, "VersionString")
    sProductVersionMsi  = GetMsiProductVersion(oMsi.ProductInfo(sProductCode, "LocalPackage"))
    sProductVersionReal = sProductVersionMsi

    If IsArray(arrSUpdatesAll) Then 
        For iIndex = 0 To UBound(arrSUpdatesAll)
            If (InStr(arrSUpdatesAll(iIndex, COL_TARGETS), sProductCode)>0) Then
                If InStr(sAppliedPatches, arrSUpdatesAll(iIndex, COL_PATCHCODE))>0 Then
                    If IsMinorUpdate(sProductCode, arrSUpdatesAll(iIndex, COL_PATCHXML)) Then
                        XmlDoc.LoadXml(arrSUpdatesAll(iIndex, COL_PATCHXML))
                        Set Elements = XmlDoc.GetElementsByTagName("TargetProduct")
                        For Each Element in Elements
                            If Element.selectSingleNode("TargetProductCode").text = sProductCode Then
                                If Element.selectSingleNode("UpdatedVersion").text > sProductVersionReal Then _
                                   sProductVersionReal = Element.selectSingleNode("UpdatedVersion").text
                            End If
                        Next 'Element
                    End If
                End If
            End If 'InStr(arrSUpdatesAll...
        Next 'iIndex
    End If 'IsArray
    
    If NOT Err = 0 Then GetRealBuildVersion = sProductVersionReg Else GetRealBuildVersion = sProductVersionReal

End Function 'GetRealBuildVersion
'-------------------------------------------------------------------------------

'Checks if a ProductCode belongs to an Office family
Function IsOfficeProduct(sProductCode)
    On Error Resume Next
    
    IsOfficeProduct = False
    If InStr(OFFICE_ALL, UCase(Right(sProductCode, 28))) > 0 OR _
       InStr(OFFICEID, UCase(Right(sProductCode, 17))) > 0 OR _
       InStr(OFFICEDBGID, UCase(Right(sProductCode, 17))) > 0 Then
           If Not Err = 0 Then Exit Function
           IsOfficeProduct = True
    End If

End Function 'IsOfficeProduct
'-------------------------------------------------------------------------------

'Checks if a PatchCode belongs to an Office family
Function IsOfficePatch(sPatchTargets)
    Dim arrPatchTargets
    Dim Target
    
    On Error Resume Next
    IsOfficePatch = False
    If NOT Len(sPatchTargets)>1 Then Exit Function
    
    arrPatchTargets = Split(sPatchTargets, ";")
    For Each Target in arrPatchTargets
        If InStr(OFFICE_ALL, UCase(Right(Target, 28))) > 0 OR _
           InStr(OFFICEID, UCase(Right(Target, 17))) > 0 OR _
           InStr(OFFICEDBGID, UCase(Right(Target, 17))) > 0 Then
               If Not Err = 0 Then Exit Function
               IsOfficePatch = True
               Exit Function
        End If
    Next 'Target
End Function 'IsOfficePatch
'==============================================================================================

'Verify Windows Installer metadata are in a healthy state and initiate fixup if needed
Sub EnsurePatchMetadata (Patch, sProductUserSid)
    Const MSISOURCETYPE_NETWORK = 1
    Const MSITRANSFORMERROR_ALL = 319

    
    Dim RegItem, Folder, File, RegItems, MspDb, Record
    Dim dicTransforms
    dim qView
    Dim fGlobalConfigPatchExists, fGlobalConfigProductExists, fConfigPatchExists
    Dim fConfigProductExists, fLocalPatchPackageDataExist, fNoError
    Dim sPatchCodeCompressed, sRegUserSID, sHive, sRegHive, sRegKey, sRegName, sTmp, sProductCode
    Dim sProductCodeCompressed, sItem, sLocalMSP, sSqlCreateTable, sMst, sDiskPrompt, sVolumeLabel
    Dim sDiskId, sKey, sMspClasses, sPackageName
    Dim iContext
    
    On Error Resume Next
    fLocalPatchPackageDataExist = False
    fGlobalConfigPatchExists = False
    fGlobalConfigProductExists = False
    fConfigPatchExists = False
    fConfigProductExists = False
    fNoError = True

    sPatchCodeCompressed = GetCompressedGuid(Patch.PatchCode)
    sProductCode = Patch.ProductCode
    sProductCodeCompressed = GetCompressedGuid(sProductCode)

    Err.Clear
    sLocalMSP = ""
    sLocalMSP = Patch.Patchproperty("LocalPackage")
    Select Case (Err.number)
    Case 0
        fLocalPatchPackageDataExist = True
    
    Case -2147023249
        'MSI API Error -> Failed to get value for local package
        fNoError = False
        Err.Clear

    Case Else
        'Unexpected Error
        fNoError = False
        Err.Clear

    End Select

    'Prepare UserSID variable needed for registry operations
    sRegUserSID = "S-1-5-18\"
    If NOT sProductUserSid = "" Then sRegUserSID = sProductUserSid & "\"

    'Check Global Config location
    '============================
    sHive = HKLM
    sRegHive = "HKEY_LOCAL_MACHINE\"

    'Check local package registration
    '--------------------------------
    'REG_GLOBALCONFIG = "Software\Microsoft\Windows\CurrentVersion\Installer\UserData\"
    sRegKey = REG_GLOBALCONFIG & sRegUserSID & "Patches\" & sPatchCodeCompressed & "\"
    sRegName = "LocalPackage"
    fGlobalConfigPatchExists = RegValExists(sHive, sRegKey, sRegName)
    If Not fGlobalConfigPatchExists Then 
        sTmp =  vbTab & "Missing patch metadata. Failed to read value from: " & sRegHive & sRegKey & sRegName
        Log sTmp
        LogSummary sProductCode, sTmp
        fNoError = False
    End If

    'Check if patchkey exists for the product
    '----------------------------------------
    'REG_GLOBALCONFIG = "Software\Microsoft\Windows\CurrentVersion\Installer\UserData\"
    sRegKey = REG_GLOBALCONFIG & sRegUserSID & "Products\" & sProductCodeCompressed & "\Patches\"
    fGlobalConfigProductExists = RegKeyExists (sHive, sRegKey & sPatchCodeCompressed)
    If Not fGlobalConfigProductExists Then 
        sTmp =  vbTab & "Missing patch metadata. Failed to locate key: " & sRegHive & sRegKey & sPatchCodeCompressed & "\"
        Log sTmp
        LogSummary sProductCode, sTmp
        'This could be related to a Windows Installer Upgrade scenario
    End If

    'Check per-user/per-machine/managed location
    '===========================================
    Select Case (Patch.Context)
    Case 1
        'Context = "USERMANAGED"
        sHive = HKLM
        sRegHive = "HKEY_LOCAL_MACHINE\"
        'REG_PRODUCTPERUSERMANAGED = "Software\Microsoft\Windows\CurrentVersion\Installer\Managed\"
        sRegKey = REG_PRODUCTPERUSERMANAGED & sRegUserSID & "\Installer\Patches\" & sPatchCodeCompressed & "\"
    Case 2
        'Context = "USER UNMANAGED"
        sHive = HKU
        sRegHive = "HKEY_USERS\"
        'REG_PRODUCT = "Software\Classes\Installer\"
        sRegKey = sRegUserSID & REG_PRODUCT & "Patches\" & sPatchCodeCompressed & "\"
    Case Else ' = Case 4
        'Context = "MANAGED"
        sHive = HKLM
        sRegHive = "HKEY_LOCAL_MACHINE\"
        'REG_PRODUCT = "Software\Classes\Installer\"
        sRegKey = REG_PRODUCT & "Patches\" & sPatchCodeCompressed & "\"
    End Select
    sMspClasses = sRegKey

    'Check registration in 'Patches' section
    '---------------------------------------
    fConfigPatchExists = RegKeyExists(sHive, sRegKey & "SourceList")
    If Not fConfigPatchExists Then 
        sTmp =  vbTab & "Missing patch metadata. Failed to locate key: " & sRegHive & sRegKey & "SourceList\"
        Log sTmp
        LogSummary sProductCode, sTmp
        fNoError = False
    End If

    'Check if patchkey exists for the product
    '----------------------------------------
    Select Case (Patch.Context)
    Case 1
        'Context = "USERMANAGED"
        'REG_PRODUCTPERUSERMANAGED = "Software\Microsoft\Windows\CurrentVersion\Installer\Managed\"
        sRegKey = REG_PRODUCTPERUSERMANAGED & sRegUserSID & "\Installer\Products\"& sProductCodeCompressed & "\Patches\"
    Case 2
        'Context = "USER UNMANAGED"
        'REG_PRODUCTPERUSER = "Software\Microsoft\Installer\"
        sRegKey = sRegUserSID & REG_PRODUCTPERUSER & "Products\" & sProductCodeCompressed & "\Patches\"
    Case Else ' = Case 4
        'Context = "MANAGED"
        'REG_PRODUCT = "Software\Classes\Installer\"
        sRegKey = REG_PRODUCT & "Products\" & sProductCodeCompressed & "\Patches\"
    End Select
    sRegName = sPatchCodeCompressed
    fConfigProductExists = RegValExists(sHive, sRegKey, sRegName)
    If Not fConfigProductExists Then 
        sTmp = vbTab & "Missing patch metadata. Failed to read value from: " & sRegHive & sRegKey & sRegName
        Log sTmp
    End If

    sRegName = "Patches"
    sTmp = ""
    RegItems = oWShell.RegRead(sRegHive & sRegKey & sRegName)
    For Each sItem In RegItems
        sTmp = sTmp & sItem
    Next

    fConfigProductExists = (InStr(sTmp, sPatchCodeCompressed) > 0)

    If NOT fNoError AND NOT fDetectOnly Then 'FixMspReg Patch
        sTmp = vbTab & "Repair: '" & Patch.PatchCode & "'. Fixing patch registration."
        If fDetectOnly Then sTmp = vbTab & "Error: Registration broken for '" & Patch.PatchCode & "'. Patch registration would be fixed."
        Log sTmp
        LogSummary Patch.PatchCode, sTmp

        If NOT fLocalPatchPackageDataExist Then FixMspGlobalReg Patch.PatchCode
        If NOT fConfigPatchExists Then
            Set MspDb = oMsi.OpenDatabase(sLocalMSP, MSIOPENDATABASEMODE_PATCHFILE)
            sSqlCreateTable = "CREATE TABLE `Media` (`DiskId` SHORT NOT NULL, `LastSequence` LONG NOT NULL, `DiskPrompt` CHAR(64) LOCALIZABLE, `Cabinet` CHAR(255), `VolumeLabel` CHAR(32), `Source` CHAR(72) PRIMARY KEY `DiskId`)"
            MspDb.OpenView(sSqlCreateTable).Execute

            'Get the patch embedded transforms
            Set dicTransforms = CreateObject("Scripting.Dictionary")
            Set qView = MspDb.OpenView("SELECT `Name` FROM `_Storages` ORDER BY `Name`") : qView.Execute
            Set Record = qView.Fetch
            Do Until Record Is Nothing
                dicTransforms.Add Record.StringData(1), Record.StringData(1)
                Set Record = qView.Fetch
            Loop
            qView.Close

            'Apply the patch transforms to the patch itself
            For Each sMst in dicTransforms.Keys
                MspDb.ApplyTransform ":" & sMst, MSITRANSFORMERROR_ALL
                Set TestSumInfo = MspDb.SummaryInformation
            Next 'sMst

            'Obtain the DiskPrompt and VolumeLabel
            Set qView = MspDb.OpenView("SELECT * FROM `_TransformView` WHERE `Table` = 'Media' ORDER BY `Row`")
            qView.Execute()
            Set Record = qView.Fetch
            If NOT Record Is Nothing Then
                sKey = Record.StringData(3)
            End If
            Do Until Record Is Nothing
                'Next FTK?
                If NOT sKey = Record.StringData(3) Then Exit Do
                'Add data from _TransformView
                Select Case Record.StringData(2)
                Case "DiskId"
                    sDiskId = Record.StringData(4)
                Case "DiskPrompt"
                    sDiskPrompt = Record.StringData(4)
                Case "VolumeLabel"
                    sVolumeLabel = Record.StringData(4)
                Case "CREATE"
                Case "DELETE"
                Case "DROP"
                Case "INSERT"
                Case Else
                End Select
                Set Record = qView.Fetch
            Loop
            qView.Close
            
            'StdPackageName
            Set qView = MspDb.OpenView("SELECT `Property`, `Value` FROM MsiPatchMetadata WHERE `Property`='StdPackageName'")
            qView.Execute : Set Record = qView.Fetch()
            If Not Record Is Nothing Then
                sPackageName = Record.StringData(2)
            Else
                sPackageName = ""
            End If
            qView.Close

            Patch.SourceListAddSource MSISOURCETYPE_NETWORK, sWICacheDir, 0
            Patch.SourceListInfo("DiskPrompt") = oMsi.ProductInfo(Patch.ProductCode, "ProductName")
            oWShell.RegWrite sRegHive & sMspClasses & "SourceList\PackageName", sPackageName, "REG_SZ"
            oWShell.RegWrite sRegHive & sMspClasses & "SourceList\Media\100", sVolumeLabel & ";" & sDiskPrompt, "REG_SZ"
        End If

    End If 'FixMspReg

End Sub 'EnsurePatchMetadata
'==============================================================================================


Sub FixMspGlobalReg(sPatchCode)

    Dim sPatchCodeCompressed, sGlobalPatchKey, sValue

    On Error Resume Next
    
    sPatchCodeCompressed = GetCompressedGuid(sPatchCode)
    sGlobalPatchKey = REG_GLOBALCONFIG & "S-1-5-18\Patches\" & sPatchCodeCompressed & "\"
    'Create the registry key
    If NOT RegKeyExists(HKLM, REG_GLOBALCONFIG & "S-1-5-18\Patches\") Then oReg.CreateKey HKLM, REG_GLOBALCONFIG & "S-1-5-18\Patches\"
    If NOT RegKeyExists(HKLM, sGlobalPatchKey) Then oReg.CreateKey HKLM, sGlobalPatchKey
    
    'Obtain a filename.
    'If the file already exists in the installer cache - use that one. If not use a random filename
    If dicRepair.Exists(sPatchCode) Then 
        If InStr(LCase(dicRepair.Item(sPatchCode)), LCase(sWICacheDir))>0 Then sValue = dicRepair.Item(sPatchCode) Else sValue = GetRandomMspName
    Else 
        sValue = GetRandomMspName
    End If

    'Create the registry value
    oReg.SetStringValue HKLM, sGlobalPatchKey, "LocalPackage", sValue
    

End Sub 'FixMspGlobalReg
'-------------------------------------------------------------------------------

'Only supports per-machine installations!
Sub UpdateProductVersion(sProductCode, sProductVersion)

    Dim sProductCodeCompressed, sHive, sGlobalConfigKey

    On Error Resume Next
    
    sProductCodeCompressed = GetCompressedGuid(sProductCode)
    sHive = HKLM
    sGlobalConfigKey = REG_GLOBALCONFIG & "S-1-5-18\Products\" & sProductCodeCompressed & "\InstallProperties\"
    If RegKeyExists(sHive, sGlobalConfigKey) Then oReg.SetStringValue sHive, sGlobalConfigKey, "DisplayVersion", sProductVersion

End Sub 'UpdateProductVersion
'-------------------------------------------------------------------------------

Sub UnregisterPatch(Patch)

    Dim PatchRef, PatchR, value
    Dim sHive, sKey, sPatchCodeCompressed, sProductCodeCompressed, sUserSid, sPatchKey, sProductKey, sPatchList
    Dim sGlobalConfigKey, sGlobalPatchKey
    Dim i
    Dim fReturn
    Dim arrMultiSzValues, arrMultiSzNewValues, arrTest
    
    On Error Resume Next

    'Ensure empty variables
    sHive = "" : sKey = "" : sPatchCodeCompressed = "" : sProductCodeCompressed = "" : sUserSid = ""
    sPatchKey = "" : sProductKey = "" : sPatchList = "" 
    ReDim arrMultiSzNewValues(-1)
    i = -1

    'Fill variables
    sPatchCodeCompressed = GetCompressedGuid(Patch.PatchCode)
    sProductCodeCompressed = GetCompressedGuid(Patch.ProductCode)
    sUserSid = Patch.UserSid : If sUserSid = "" Then sUserSid = "S-1-5-18\" Else sUserSid = sUserSid & "\"
    sGlobalConfigKey = REG_GLOBALCONFIG & sUserSid & "Products\" & sProductCodeCompressed & "\Patches\" 
    sGlobalPatchKey = REG_GLOBALCONFIG & sUserSid & "Patches\" & sPatchCodeCompressed & "\"
    
    If Err <> 0 Then Exit Sub
    
    Select Case (Patch.Context)
    Case MSIINSTALLCONTEXT_USERMANAGED '1
        sHive = HKLM
        sPatchKey =  REG_PRODUCTPERUSERMANAGED & sUserSid & "Installer\Patches\" & sPatchCodeCompressed & "\"
        sProductKey = REG_PRODUCTPERUSERMANAGED & "Products\" & sProductCodeCompressed & "\Patches\" 
    
    Case MSIINSTALLCONTEXT_USERUNMANAGED '2
        sHive = HKCU
        sPatchKey = REG_PRODUCTPERUSER & "Patches\" & sPatchCodeCompressed & "\"
        sProductKey = REG_PRODUCTPERUSER & "Products\" & sProductCodeCompressed & "\Patches\" 
    
    Case Else 'Case MSIINSTALLCONTEXT_MACHINE '4 (Managed)
        sHive = HKLM
        sPatchKey = REG_PRODUCT & "Patches\" & sPatchCodeCompressed & "\"
        sProductKey = REG_PRODUCT & "Products\" & sProductCodeCompressed & "\Patches\" 
    End Select

    'Unregister the patch from ProductKey
    If RegReadMultiStringValue(sHive, sProductKey, "Patches", arrMultiSzValues) Then
        For Each value in arrMultiSzValues
            If Not value = sPatchCodeCompressed Then 
                i = i + 1 
                ReDim Preserve arrMultiSzNewValues(i)
                arrMultiSzNewValues(i) = value
            End If
        Next 'Value
    End If
    fReturn = oReg.GetMultiStringValue(sHive, sProductKey, "Patches", arrTest)
    If fReturn = 0 Then
        Log vbTab & vbTab & "Updating value " & HiveString(sHive) & "\" & sProductKey & "Patches"
        If NOT fDetectOnly Then oReg.SetMultiStringValue sHive, sProductKey, "Patches", arrMultiSzNewValues
    Else
        If fx64 Then
            fReturn = oReg.GetMultiStringValue(sHive, Wow64Key(sHive, sProductKey), "Patches", arrMultiSzNewValues)
            If fReturn = 0 Then
                Log vbTab & vbTab & "Updating value " & HiveString(sHive) & "\" & Wow64Key(sHive, sProductKey) & "Patches"
                If NOT fDetectOnly Then oReg.SetMultiStringValue sHive, Wow64Key(sHive, sProductKey), "Patches", arrMultiSzNewValues
            End If
        End If 'fx64
    End If
    RegDeleteValue sHive, sProductKey, sPatchCodeCompressed

    'Unregister PatchKey
    RegDeleteKey sHive, sPatchKey

    'Unregister GlobalConfigKey
    ReDim arrMultiSzNewValues(-1)
    i = -1
    If RegReadMultiStringValue(sHive, sGlobalConfigKey, "AllPatches", arrMultiSzValues) Then
        For Each Value in arrMultiSzValues
            If Not Value = sPatchCodeCompressed Then 
                i = i + 1 
                ReDim Preserve arrMultiSzNewValues(i)
                arrMultiSzNewValues(i) = Value
            End If
        Next 'Value
    End If
    fReturn = oReg.GetMultiStringValue(sHive, sGlobalConfigKey, "AllPatches", arrTest)
    If fReturn = 0 Then
        Log vbTab & vbTab & "Updating value " & HiveString(sHive) & "\" & sGlobalConfigKey & "AllPatches"
        If NOT fDetectOnly Then oReg.SetMultiStringValue sHive, sGlobalConfigKey, "AllPatches", arrMultiSzNewValues
    Else
        If fx64 Then
            fReturn = oReg.GetMultiStringValue(sHive, Wow64Key(sHive, sGlobalConfigKey), "AllPatches", arrMultiSzNewValues)
            If fReturn = 0 Then
                Log vbTab & vbTab & "Updating value " & HiveString(sHive) & "\" & Wow64Key(sHive, sGlobalConfigKey) & "AllPatches"
                If NOT fDetectOnly Then oReg.SetMultiStringValue sHive, Wow64Key(sHive, sGlobalConfigKey), "AllPatches", arrMultiSzNewValues
            End If
        End If 'fx64
    End If
    
    RegDeleteKey HKLM, sGlobalConfigKey & sPatchCodeCompressed & "\"

    'Unregister sGlobalPatchKey
    RegDeleteKey HKLM, sGlobalPatchKey
End Sub
'-------------------------------------------------------------------------------

' Fills the dicFeatureStates dictionary for the current product
Function GetFeatureStates (Product)
    Dim feature, features
    Dim sAbsent
    
    Set features = oMsi.Features (Product)
    For Each feature in features
        If NOT dicFeatureStates.Exists (feature) Then
            dicFeatureStates.Add feature, oMsi.FeatureState (Product, feature)
        End If
    Next 'feature
End Function 'GetFeatureStates
'-------------------------------------------------------------------------------

' Set the feature control flag for this patch
' The feature control needs to be enabled if the patch adds a new component
' Returns true or false
Function GetFeatureControl (sProductCode, sPatchFile)
    Dim MspDb, key
    Dim sXml
    Dim i, iRow, iCol
    Dim fReturn, fFoo
    Dim arrMspTargets, arrMspTransforms, arrMstDetail, arrTable, arrColHeaders
    Dim dicTransformRow

    On Error Resume Next

    fReturn = False
    If arrSUpdatesAll (iIndex, COL_PACKAGE) = "OCT" Then
        'keep the default
    Else
        'check if the patch contains a 'Component' table
        ' load the patch for inspection
        Set MspDb = oMsi.OpenDatabase(sPatchFile, MSIOPENDATABASEMODE_PATCHFILE)
        ' init required patch data
        arrMspTargets = Split(MspDb.SummaryInformation.Property(PID_TEMPLATE), ";")
        arrMspTransforms = Split(MspDb.SummaryInformation.Property(PID_LASTAUTHOR), ";")
        sXml = oMsi.ExtractPatchXMLData(sPatchFile)
        XmlDoc.LoadXml(sXml)
        ' init the 'arrMstDetail array
        AddPatchTargetDetails XmlDoc, arrMstDetail, arrMspTransforms
        ' call GetPatchTargets to fill the global dicProdMst
        GetPatchTargets "", dicTransformRow, arrMspTargets, arrMstDetail, TRUE, FALSE
        ' load the schema for the Component table
        fFoo = LoadTableSchema(sProductCode, "Component", MspDb, arrMspTargets)
        ' load the patch transforms
        LoadPatchTransforms MspDb, sProductCode, dicTransformRow, arrMspTransforms, arrMstDetail
        ' read the patch table details
        ' drill down into the details to check if there's a new component added
        Set arrTable = Nothing
        For i = 0 To UBound(arrSchema)
            ' ensure filter on Component table
            If (LCase(arrSchema(i, 0)) = LCase("component")) Then 
                ' load the table details into an array
                arrColHeaders = GetTableColumnHeadersFromDef(arrSchema(i, 1))
                arrTable = GetPatchTableDetails(MspDb, arrSchema(i, 0), arrColHeaders)
                ' enumerate the details of the array
                If IsArray(arrTable) Then
                    For iRow = 0 To UBound (arrTable, 1)
                        For iCol = 0 To UBound (arrTable, 2)
                            ' the second column represents the ComponentCode
                            ' an existing component won't have a value in there, a new one does list the ComponentCode
                            fReturn = fReturn OR (NOT IsEmpty(arrTable(iRow, 1)) AND NOT arrTable(iRow, 1) = "")
                        Next 'iCol
                    Next 'iRow
                End If
            End If
        Next 'i
        ClearView MspDb, sPatchFile
    End If

    GetFeatureControl = fReturn

End Function 'GetFeatureControl
'-------------------------------------------------------------------------------

Function ApplyPatch (sProductCode, sPatchFile, fFeatureControl)
    Dim feature
    Dim sReturn, sCmd, sAbsent, sAddLocal, sReinstall
    Dim arrFeatureMsp

    On Error Resume Next

    If fFeatureControl Then
        ' prepare the Features list
        For Each feature in dicFeatureStates.Keys
            Select Case dicFeatureStates.Item (feature)
            Case 1  'msiInstallStateAdvertised
            Case 2  'msiInstallStateAbsent
                sAbsent = sAbsent & "," & feature
            Case 3  'msiInstallStateLocal
                sReinstall = sReinstall & "," & feature
            Case 4  'msiInstallStateSource
            Case -2 'msiInstallStateInvalidArg
            Case -1 'msiInstallStateUnknown
            Case -6 'msiInstallStateBadConfig
            End Select
        Next 'feature

        ' add new Features to ADDLOCAL
        sAddLocal = GetAddLocal (sProductCode, sPatchFile)
        arrFeatureMsp = Split (sAddLocal, ",")
        ' ensure this is really a new feature
        For Each feature in arrFeatureMsp
            If dicFeatureStates.Exists (feature) Then
                sAddLocal = Replace (sAddLocal, "," & feature, "")
            End If
        Next
    End If 'fFeatureControl

    If Len (sAddLocal) > 0 Then sAddLocal = " ADDLOCAL=" & sAddLocal Else sAddLocal = ""
    If Len (sAbsent) > 0 Then sAbsent = " REMOVE=" & Mid (sAbsent, 2) Else sAbsent = ""
    If Len (sReinstall) > 0 Then
        sReinstall = " REINSTALL=" & Mid (sReinstall, 2) & " REINSTALLMODE=omu"
        sAbsent = "" ' no need to have both configured REINSTALL and ABSENT
    Else
        sReinstall = ""
    End If
    
    ' build the patch apply command
    sCmd = "msiexec.exe /i " & sProductCode & _
              " PATCH=" & chr(34) & sPatchFile & chr(34) & _
              " REBOOT=ReallySuppress" & _
              " MSIRESTARTMANAGERCONTROL=Disable" & _
              sAddLocal & _
              sAbsent & _
              sReinstall & _
              " /qb-" & _
              " /l*v " & chr(34) & sPathOutputFolder & sComputerName & "_" & sProductCode & "_MspApply.log" & chr(34)
    If NOT fDisableRestartManager Then sCmd = Replace (sCmd, " MSIRESTARTMANAGERCONTROL=Disable", "")
    sTmp = "Calling msiexec to apply patch: " & sCmd
    If fDetectOnly Then
        sPatchFile = Replace (sPatchFile, ";", vbCrLf & vbTab & vbTab)
        sTmp = "Applicable patch: " & sPatchFile
    End If 'fDetectOnly
    Log vbTab & "Debug:  " & sTmp
    LogSummary sProductCode, vbTab & sTmp
    If fCscript Then wscript.echo vbTab & vbTab & sTmp
    
    'Execute the patch apply command
    If NOT fDetectOnly Then 
        sReturn = CStr (oWShell.Run (sCmd, 0, True))
        ApplyPatch = MsiexecRetVal (sReturn)
        sTmp = "Msiexec returned with code: " & sReturn & " " & MsiexecRetval (sReturn)
        Log vbTab & "Debug:  " & sTmp
        LogSummary sProductCode, vbTab & sTmp
        If fCscript Then wscript.echo vbTab & vbTab & sTmp
        fRebootRequired = fRebootRequired OR (sReturn = "3010")
    End If 'NOT fDetectOnly

End Function 'ApplyPatch
'-------------------------------------------------------------------------------

Function GetAddLocal (Product, MspFile)
    Dim MsiDb, MspDb, tbl, targetProduct, feature
    Dim i, iPos, h
    Dim sDef, sXml, sMst, sErr, sReturn, sProductVersion
    Dim arrTables, arrMst, arrMspTransforms, arrColHeaders, arrTable

    sProductVersion = oMsi.ProductInfo (Product, "VersionString")
    ' get the table schema
    sDef = ""
    sReturn = ""
    Set MsiDb = oMsi.OpenDatabase(oMsi.ProductInfo (Product, "LocalPackage"), MSIOPENDATABASEMODE_READONLY)
    arrTables = Split (GetDatabaseTables (MsiDb), ",")
    For Each tbl in arrTables
        If tbl = "Feature" Then
            sDef = "`" & tbl & "` (" & GetTableColumnDef (MsiDb, tbl) & " PRIMARY KEY " & GetPrimaryTableKeys (MsiDb, tbl) & ")"
        End If
    Next 'tbl
    If sDef = "" Then
        GetAddLocal = sDef
        Exit Function
    End If
    
    ' determine the index position of the the patch transform by counting the product position in the PatchXml
    sXml = oMsi.ExtractPatchXMLData(MspFile)
    If IsValidVersion (Product, sXml, sProductVersion, sErr, iPos) Then
        iPos = iPos * 2

        ' open a handle to the patch 
        Set MspDb = oMsi.OpenDatabase(MspFile, MSIOPENDATABASEMODE_PATCHFILE)

        ' create the table view
        MspDb.OpenView ("CREATE TABLE " & sDef).Execute

        ' get the transform name
        arrMspTransforms = Split(MspDb.SummaryInformation.Property(PID_LASTAUTHOR), ";")
        sMst = Mid (arrMspTransforms (iPos), 2)

        ' apply the patch transform to the patch
        MspDb.ApplyTransform ":" & sMst, 319 'MSITRANSFORMERROR_ALL
        MspDb.ApplyTransform ":#" & sMst, 319 'MSITRANSFORMERROR_ALL

        ' get items from Feature table (if any)
        arrColHeaders = GetTableColumnHeadersFromDef (sDef)
        arrTable = GetPatchTableDetails (MspDb, "Feature", arrColHeaders)
        ' filter on FeatureNames
        If IsArray(arrTable) Then
            For i = 0 To UBound (arrTable, 1)
                If Len (arrTable (i, 0)) > 0 Then
                    wscript.echo "adding AddLocal feature: " & arrTable (i, 0) 
                    sReturn = sReturn & "," & arrTable (i, 0)
                End If
            Next
        End If
    End If 'IsValidVersion
    If Len (sReturn) > 0 Then sReturn = Mid (sReturn, 2)

    GetAddLocal = sReturn

End Function 'GetAddLocal
'-------------------------------------------------------------------------------

'Return the primary keys of a table by using the PrimaryKeys property of the database object
'in SQL ready syntax 
Function GetPrimaryTableKeys(MsiDb, sTable)
    Dim iKeyCnt
    Dim sPrimaryTmp
    Dim PrimaryKeys
    On Error Resume Next

    sPrimaryTmp = ""
    Set PrimaryKeys = MsiDb.PrimaryKeys(sTable)
    For iKeyCnt = 1 To PrimaryKeys.FieldCount
        sPrimaryTmp = sPrimaryTmp & "`" & PrimaryKeys.StringData(iKeyCnt) & "`, "
    Next 'iKeyCnt
    GetPrimaryTableKeys = Left(sPrimaryTmp, Len(sPrimaryTmp)-2)
End Function 'GetPrimaryTableKeys
'-------------------------------------------------------------------------------

'Return the Column schema definition of a table in SQL ready syntax
Function GetTableColumnDef(MsiDb, sTable)
    Const MSICOLUMNINFONAMES                = 0
    Const MSICOLUMNINFOTYPES                = 1

    Dim sQuery, sColDefTmp
    Dim View, ColumnNames, ColumnTypes
    Dim iColCnt

    On Error Resume Next
    'Get the ColumnInfo details
    sColDefTmp = ""
    sQuery = "SELECT * FROM " & sTable
    Set View = MsiDb.OpenView(sQuery)
    View.Execute
    Set ColumnNames = View.ColumnInfo(MSICOLUMNINFONAMES)
    Set ColumnTypes = View.ColumnInfo(MSICOLUMNINFOTYPES)
    For iColCnt = 1 To ColumnNames.FieldCount
        sColDefTmp = sColDefTmp & ColDefToSql(ColumnNames.StringData(iColCnt), ColumnTypes.StringData(iColCnt)) & ", "
    Next 'iColCnt
    View.Close
    
    GetTableColumnDef = Left(sColDefTmp, Len(sColDefTmp)-2)
    
End Function 'GetTableColumnDef
'-------------------------------------------------------------------------------

'Return the Column header names
Function GetTableColumnHeaders(MsiDb, sTable)
    Const MSICOLUMNINFONAMES                = 0
    
    Dim sQuery, sColDefTmp
    Dim View, ColumnNames, ColumnTypes
    Dim iColCnt
    
    On Error Resume Next
    GetTableColumnHeaders = ""

    'Get the ColumnInfo details
    sColDefTmp = ""
    sQuery = "SELECT * FROM " & sTable
    Set View = MsiDb.OpenView(sQuery)
    View.Execute
    Set ColumnNames = View.ColumnInfo(MSICOLUMNINFONAMES)
    For iColCnt = 1 To ColumnNames.FieldCount
        sColDefTmp = sColDefTmp & "," & ColumnNames.StringData(iColCnt)
    Next 'iColCnt
    View.Close
    
    If NOT sColDefTmp="" Then GetTableColumnHeaders = Mid(sColDefTmp, 2)
    
End Function 'GetTableColumnHeaders
'-------------------------------------------------------------------------------

'Return array with column header names from Sql definition string
Function GetTableColumnHeadersFromDef (ByVal sSqlDef)

    'ColumnHeaders start after the opening bracket and end with "PRIMARY KEY"
    'Each column is separated with a ","
    'Each column header is then enclosed in "`" -> `<ColumnName>`

    Dim i
	Dim sTable
    Dim arrHeaders

    GetTableColumnHeadersFromDef = ""

	sTable = Mid(sSqlDef, 2, InStr(sSqlDef, "` (")-2)
    'strip intro
    sSqlDef = Mid(sSqlDef, InStr(sSqlDef, "(")+1)
    'crop end
    sSqlDef = Left(sSqlDef, InStr(sSqlDef, " PRIMARY KEY"))
    ' exception handler
	Select Case sTable
	Case "File"
		sSqlDef = sSqlDef & ", `Predicted Action`, `ComponentState`, `FileSize On Disk`, `Version On Disk`, `Hash On Disk`, `FilePath`"
	Case Else
	End Select
	'split into array
    arrHeaders = Split(sSqlDef, ",")
    For i = 0 To UBound(arrHeaders)
        'trim blanks
        arrHeaders(i) = Trim(arrHeaders(i))
        'strip opening `
        arrHeaders(i) = Mid(arrHeaders(i), 2)
        'crop at closing `
        arrHeaders(i) = Left(arrHeaders(i), InStr(arrHeaders(i), "`")-1)
    Next

    GetTableColumnHeadersFromDef = arrHeaders

End Function 'GetKnownTableColumnHeaders
'-------------------------------------------------------------------------------
'Translate the column definition fields into SQL syntax
Function ColDefToSql(sColName, sColType)
    On Error Resume Next
    
    Dim iLen
    Dim sRight, sLeft, sSqlTmp

    iLen = Len(sColType)
    sRight = Right(sColType, iLen-1)
    sLeft = Left(sColType, 1)
    sSqlTmp = "`" & sColName & "`"
    Select Case sLeft
    Case "s", "S"
        's? String, variable length (?=1-255) -> CHAR(#) or CHARACTER(#)
        's0 String, variable length -> LONGCHAR
        If sRight="0" Then sSqlTmp = sSqlTmp & " LONGCHAR" Else sSqlTmp = sSqlTmp & " CHAR(" & sRight & ")"
        If sLeft = "s" Then sSqlTmp = sSqlTmp & " NOT NULL"
    Case "l", "L"
        'CHAR(#) LOCALIZABLE or CHARACTER(#) LOCALIZABLE
        If sRight="0" Then sSqlTmp = sSqlTmp & " LONGCHAR" Else sSqlTmp = sSqlTmp & " CHAR(" & sRight & ")"
        If sLeft = "l" Then sSqlTmp = sSqlTmp & " NOT NULL"
        If sRight="0" Then sSqlTmp = sSqlTmp & "  LOCALIZABLE" Else sSqlTmp = sSqlTmp & " LOCALIZABLE"
    Case "i", "I"
        'i2 Short integer 
        'i4 Long integer 
        If sRight="2" Then sSqlTmp = sSqlTmp & " SHORT" Else sSqlTmp = sSqlTmp & " LONG"
        If sLeft = "i" Then sSqlTmp = sSqlTmp & " NOT NULL"
    Case "v", "V"
        'v0 Binary Stream 
        sSqlTmp = sSqlTmp & " OBJECT"
        If sLeft = "v" Then sSqlTmp = sSqlTmp & " NOT NULL"
    Case "g", "G"
        'g? Temporary string (?=0-255)
    Case "j", "J"
        'j? Temporary integer (?=0, 1, 2, 4)) 
    Case "o", "O"
        'O0 Temporary object 
    Case Else
    End Select

    ColDefToSql = sSqlTmp

End Function 'ColDefToSql
'-------------------------------------------------------------------------------


'Registry Helper Routines
'------------------------

'-------------------------------------------------------------------------------

'Register context menu
Sub RegisterShellExt

    'Ensure to unregister old contents first
    UnRegisterShellExt
    
    'Register 
    oReg.CreateKey HKCR, "Msi.Patch"
    oReg.CreateKey HKCR, "Msi.Patch\shell"
    
'    oReg.CreateKey HKCR, "Msi.Patch\shell\OPUtil ApplyPatch"
'    oReg.CreateKey HKCR, "Msi.Patch\shell\OPUtil ApplyPatch\command"
'    oReg.SetStringValue HKCR, "Msi.Patch\shell\OPUtil ApplyPatch\command", , "wscript " & chr(34) & wscript.ScriptFullName & chr(34) & " /ContextMenu /ApplyPatch=" & chr(34) & "%1%" & chr(34)
    
    oReg.CreateKey HKCR, "Msi.Patch\shell\OPUtil CabExtract"
    oReg.CreateKey HKCR, "Msi.Patch\shell\OPUtil CabExtract\command"
    oReg.SetStringValue HKCR, "Msi.Patch\shell\OPUtil CabExtract\command", , "wscript " & chr(34) & wscript.ScriptFullName & chr(34) & " /ContextMenu /CabExtract=" & chr(34) & "%1%" & chr(34)

'    oReg.CreateKey HKCR, "Msi.Patch\shell\OPUtil RemovePatch"
'    oReg.CreateKey HKCR, "Msi.Patch\shell\OPUtil RemovePatch\command"
'    oReg.SetStringValue HKCR, "Msi.Patch\shell\OPUtil RemovePatch\command", , "wscript " & chr(34) & wscript.ScriptFullName & chr(34) & " /ContextMenu /RemovePatch=" & chr(34) & "%1%" & chr(34)

    oReg.CreateKey HKCR, "Msi.Patch\shell\OPUtil ViewPatch"
    oReg.CreateKey HKCR, "Msi.Patch\shell\OPUtil ViewPatch\command"
    oReg.SetStringValue HKCR, "Msi.Patch\shell\OPUtil ViewPatch\command", , "wscript " & chr(34) & wscript.ScriptFullName & chr(34) & " /ContextMenu /ViewPatch=" & chr(34) & "%1%" & chr(34)

'    oReg.CreateKey HKCR, "Msi.Patch\shell\OPUtil ViewPatch (DeepScan)"
'    oReg.CreateKey HKCR, "Msi.Patch\shell\OPUtil ViewPatch (DeepScan)\command"
'    oReg.SetStringValue HKCR, "Msi.Patch\shell\OPUtil ViewPatch (DeepScan)\command", , "wscript " & chr(34) & wscript.ScriptFullName & chr(34) & " /ContextMenu /DeepScan /ViewPatch=" & chr(34) & "%1%" & chr(34)
    
End Sub 'RegisterShellExt
'-------------------------------------------------------------------------------

'Register context menu
Sub UnRegisterShellExt

    Dim arrKeys
    Dim Key
    Dim sSubKeyName

    sSubKeyName = "Msi.Patch\shell\"
    If (oReg.EnumKey(HKCR, sSubKeyName, arrKeys)=0) AND IsArray(arrKeys) Then
        For Each Key in arrKeys
            If InStr(Key, "OPUtil")>0 Then RegDeleteKey HKCR, sSubKeyName & Key
        Next 'Key
    End If

End Sub 'RegisterShellExt
'-------------------------------------------------------------------------------

Function RegKeyExists(hDefKey, sSubKeyName)
    On Error Resume Next
    Dim arrKeys
    RegKeyExists = False
    If oReg.EnumKey(hDefKey, sSubKeyName, arrKeys) = 0 Then RegKeyExists = True
End Function
'-------------------------------------------------------------------------------

Function HiveString(hDefKey)
    On Error Resume Next
    Select Case hDefKey
        Case HKCR : HiveString = "HKEY_CLASSES_ROOT"
        Case HKCU : HiveString = "HKEY_CURRENT_USER"
        Case HKLM : HiveString = "HKEY_LOCAL_MACHINE"
        Case HKU  : HiveString = "HKEY_USERS"
        Case Else : HiveString = hDefKey
    End Select
End Function
'-------------------------------------------------------------------------------

Function RegValExists(hDefKey, sSubKeyName, sName)
    Dim arrValueTypes, arrValueNames
    Dim i

    On Error Resume Next
    RegValExists = False
    If Not RegKeyExists(hDefKey, sSubKeyName) Then Exit Function
    If oReg.EnumValues(hDefKey, sSubKeyName, arrValueNames, arrValueTypes) = 0 AND IsArray(arrValueNames) Then
        For i = 0 To UBound(arrValueNames) 
            If LCase(arrValueNames(i)) = Trim(LCase(sName)) Then RegValExists = True
        Next 
    End If 'oReg.EnumValues
End Function
'-------------------------------------------------------------------------------

Function RegReadMultiStringValue(hDefKey, sSubKeyName, sName, arrValues)
    Dim RetVal

    On Error Resume Next
    RetVal = oReg.GetMultiStringValue(hDefKey, sSubKeyName, sName, arrValues)
    If Not RetVal = 0 AND fx64 Then RetVal = oReg.GetMultiStringValue(hDefKey, Wow64Key(hDefKey, sSubKeyName), sName, arrValues)
    
    RegReadMultiStringValue = (RetVal = 0 AND IsArray(arrValues))
End Function 'RegReadMultiStringValue
'-------------------------------------------------------------------------------

'Enumerate a registry key to return all values
Function RegEnumValues(hDefKey, sSubKeyName, arrNames, arrTypes)
    Dim RetVal, RetVal64
    Dim arrNames32, arrNames64, arrTypes32, arrTypes64
    
    On Error Resume Next
    If fx64 Then
        RetVal = oReg.EnumValues(hDefKey, sSubKeyName, arrNames32, arrTypes32)
        RetVal64 = oReg.EnumValues(hDefKey, Wow64Key(hDefKey, sSubKeyName), arrNames64, arrTypes64)
        If (RetVal = 0) AND (Not RetVal64 = 0) AND IsArray(arrNames32) AND IsArray(arrTypes32) Then 
            arrNames = arrNames32
            arrTypes = arrTypes32
        End If
        If (Not RetVal = 0) AND (RetVal64 = 0) AND IsArray(arrNames64) AND IsArray(arrTypes64) Then 
            arrNames = arrNames64
            arrTypes = arrTypes64
        End If
        If (RetVal = 0) AND (RetVal64 = 0) AND IsArray(arrNames32) AND IsArray(arrNames64) AND IsArray(arrTypes32) AND IsArray(arrTypes64) Then 
            arrNames = RemoveDuplicates(Split((Join(arrNames32, "\") & "\" & Join(arrNames64, "\")), "\"))
            arrTypes = RemoveDuplicates(Split((Join(arrTypes32, "\") & "\" & Join(arrTypes64, "\")), "\"))
        End If
    Else
        RetVal = oReg.EnumValues(hDefKey, sSubKeyName, arrNames, arrTypes)
    End If 'fx64
    RegEnumValues = ((RetVal = 0) OR (RetVal64 = 0)) AND IsArray(arrNames) AND IsArray(arrTypes)
End Function 'RegEnumValues
'-------------------------------------------------------------------------------

'Enumerate a registry key to return all subkeys
Function RegEnumKey(hDefKey, sSubKeyName, arrKeys)
    Dim RetVal, RetVal64
    Dim arrKeys32, arrKeys64
    
    On Error Resume Next
    If fx64 Then
        RetVal = oReg.EnumKey(hDefKey, sSubKeyName, arrKeys32)
        RetVal64 = oReg.EnumKey(hDefKey, Wow64Key(hDefKey, sSubKeyName), arrKeys64)
        If (RetVal = 0) AND (Not RetVal64 = 0) AND IsArray(arrKeys32) Then arrKeys = arrKeys32
        If (Not RetVal = 0) AND (RetVal64 = 0) AND IsArray(arrKeys64) Then arrKeys = arrKeys64
        If (RetVal = 0) AND (RetVal64 = 0) Then 
            If IsArray(arrKeys32) AND IsArray (arrKeys64) Then 
                arrKeys = RemoveDuplicates(Split((Join(arrKeys32, "\") & "\" & Join(arrKeys64, "\")), "\"))
            ElseIf IsArray(arrKeys64) Then
                arrKeys = arrKeys64
            Else
                arrKeys = arrKeys32
            End If
        End If
    Else
        RetVal = oReg.EnumKey(hDefKey, sSubKeyName, arrKeys)
    End If 'fx64
    RegEnumKey = ((RetVal = 0) OR (RetVal64 = 0)) AND IsArray(arrKeys)
End Function 'RegEnumKey
'-------------------------------------------------------------------------------

'Wrapper around oReg.DeleteValue to handle 64 bit
Sub RegDeleteValue(hDefKey, sSubKeyName, sName)
    Dim sWow64Key
    
    On Error Resume Next
    If RegValExists(hDefKey, sSubKeyName, sName) Then
        On Error Resume Next
        Log vbTab & vbTab & "Deleting value " & HiveString(hDefKey) & "\" & sSubKeyName & sName
        If NOT fDetectOnly Then oReg.DeleteValue hDefKey, sSubKeyName, sName
        On Error Goto 0
    End If 'RegValExists
    If fx64 Then 
        sWow64Key = Wow64Key(hDefKey, sSubKeyName)
        If RegValExists(hDefKey, sWow64Key, sName) Then
            On Error Resume Next
            Log vbTab & vbTab & "Deleting value " & HiveString(hDefKey) & "\" & sWow64Key & sName
            If NOT fDetectOnly Then oReg.DeleteValue hDefKey, sWow64Key, sName
            On Error Goto 0
        End If 'RegValExists
    End If
End Sub 'RegDeleteValue
'-------------------------------------------------------------------------------

'Wrappper around RegDeleteKeyEx to handle 64bit scenarios
Sub RegDeleteKey(hDefKey, sSubKeyName)
    Dim sWow64Key
    
    On Error Resume Next
    If RegKeyExists(hDefKey, sSubKeyName) Then
    'Get the list of patches for the product
        
        On Error Resume Next
        RegDeleteKeyEx hDefKey, sSubKeyName
        On Error Goto 0
    End If 'RegKeyExists
    If fx64 Then 
        sWow64Key = Wow64Key(hDefKey, sSubKeyName)
        If RegKeyExists(hDefKey, sWow64Key) Then
            On Error Resume Next
            RegDeleteKeyEx hDefKey, sWow64Key
            On Error Goto 0
        End If 'RegKeyExists
    End If
End Sub 'RegDeleteKey
'-------------------------------------------------------------------------------

'Recursively delete a registry structure
Sub RegDeleteKeyEx(hDefKey, sSubKeyName) 
    Dim arrSubkeys
    Dim sSubkey

    On Error Resume Next
    Do While InStr(sSubKeyName, "\\")>0
        sSubKeyName = Replace(sSubKeyName, "\\", "\")
    Loop
    If Not Right(sSubKeyName, 1)="\" Then sSubKeyName=sSubKeyName & "\"
    oReg.EnumKey hDefKey, sSubKeyName, arrSubkeys 
    If IsArray(arrSubkeys) Then 
        For Each sSubkey In arrSubkeys 
            RegDeleteKeyEx hDefKey, sSubKeyName & sSubkey & "\"
        Next 
    End If 
    Log vbTab & vbTab & "Deleting key " & HiveString(hDefKey) & "\" & sSubKeyName
    If NOT fDetectOnly Then oReg.DeleteKey hDefKey, sSubKeyName 
End Sub 'RegDeleteKeyEx
'-------------------------------------------------------------------------------

'Return the alternate regkey location on 64bit environment
Function Wow64Key(hDefKey, sSubKeyName)
    Dim iPos

    On Error Resume Next
    Select Case hDefKey
        Case HKCU
            If Left(sSubKeyName, 17) = "Software\Classes\" Then
                Wow64Key = Left(sSubKeyName, 17) & "Wow6432Node\" & Right(sSubKeyName, Len(sSubKeyName)-17)
            Else
                iPos = InStr(sSubKeyName, "\")
                Wow64Key = Left(sSubKeyName, iPos) & "Wow6432Node\" & Right(sSubKeyName, Len(sSubKeyName)-iPos)
            End If
        
        Case HKLM
            If Left(sSubKeyName, 17) = "Software\Classes\" Then
                Wow64Key = Left(sSubKeyName, 17) & "Wow6432Node\" & Right(sSubKeyName, Len(sSubKeyName)-17)
            Else
                iPos = InStr(sSubKeyName, "\")
                Wow64Key = Left(sSubKeyName, iPos) & "Wow6432Node\" & Right(sSubKeyName, Len(sSubKeyName)-iPos)
            End If
        
        Case Else
            Wow64Key = "Wow6432Node\" & sSubKeyName
        
    End Select 'hDefKey
End Function 'Wow64Key
'-------------------------------------------------------------------------------

'File Helper Routines
'------------------------

'-------------------------------------------------------------------------------

'Function to compare two numbers of unspecified format
'Return values:
'Left file version is lower than right file version     -1
'Left file version is identical to right file version    0
'Left file version is higher than right file version     1
'Invalid comparison                                      2

Function CompareVersion(sFile1, sFile2, fAllowBlanks)

    Dim file1, file2
    Dim sDelimiter
    Dim iCnt, iAsc, iMax, iF1, iF2
    Dim fLEmpty, fREmpty

    CompareVersion = 0
    fLEmpty = False
    fREmpty = False
    
    'Ensure valid inputs values
    On Error Resume Next
    If IsEmpty(sFile1) Then fLEmpty = True
    If IsEmpty(sFile2) Then fREmpty = True
    If sFile1 = "" Then fLEmpty = True
    If sFile2 = "" Then fREmpty = True

    'Don't allow alpha characters
    If Not fLEmpty Then
        For iCnt = 1 To Len(sFile1)
            iAsc = Asc(UCase(Mid(sFile1, iCnt, 1)))
            If (iAsc>64) AND (iAsc<91) Then
                CompareVersion = 2
                Exit Function
            End If
        Next 'iCnt
    End If
    If Not fREmpty Then
        For iCnt = 1 To Len(sFile2)
            iAsc = Asc(UCase(Mid(sFile2, iCnt, 1)))
            If (iAsc>64) AND (iAsc<91) Then
                CompareVersion = 2
                Exit Function
            End If
        Next 'iCnt
    End If
    
    If fLEmpty AND (NOT fREmpty) Then
        If fAllowBlanks Then CompareVersion = -1 Else CompareVersion = 2
        Exit Function
    End If
    
    If (NOT fLEmpty) AND fREmpty Then
        If fAllowBlanks Then CompareVersion = 1 Else CompareVersion = 2
        Exit Function
    End If
    
    If fLEmpty AND fREmpty Then
        If fAllowBlanks Then CompareVersion = 0 Else CompareVersion = 2
        Exit Function
    End If
    
    'If Files are identical we're already done
    If sFile1 = sFile2 Then Exit Function

    'Split the VersionString
    file1 = Split(sFile1, Delimiter(sFile1))
    file2 = Split(sFile2, Delimiter(sFile2))

    'Ensure we get the lower count
    iMax = UBound(file1)
    CompareVersion = -1
    If iMax > UBound(file2) Then 
        iMax = UBound(file2)
        CompareVersion = 1
    End If

    'Compare the file versions
    For iCnt = 0 To iMax
        iF1 = CLng(file1(iCnt))
        iF2 = CLng(file2(iCnt))
        If iF1 > iF2 Then
            CompareVersion = 1
            Exit For
        ElseIf iF1 < iF2 Then
            CompareVersion = -1
            Exit For
        End If
    Next 'iCnt
End Function
'-------------------------------------------------------------------------------

'Use WI ProvideAssembly function to identify the path for an assembly.
'Returns the path to the file if the file exists.
'Returns an empty string if file does not exist

Function GetAssemblyPath(sLfn, sKeyPath, sDir)
    On Error Resume Next
    Dim sFile, sFolder, sExt, sRoot, sName
    Dim arrTmp
    
    'Defaults
    GetAssemblyPath=""
    sFile="" : sFolder="" : sExt="" : sRoot="" : sName=""
    

    'The componentpath should already point to the correct folder
    'except for components with a registry keypath element.
    'In that case tweak the directory folder to match
    If Left(sKeyPath, 1)="0" Then
        sFolder = sDir
        sFolder = oWShell.ExpandEnvironmentStrings("%SYSTEMROOT%")&Mid(sFolder, InStr(LCase(sFolder), "\winsxs\"))
        sFile = sLfn
    End If 'Left(sKeyPath, 1)="0"
    
    'Figure out the correct file reference
    If sFolder = "" Then sFolder = Left(sKeyPath, InStrRev(sKeyPath, "\"))
    sRoot = Left(sFolder, InStrRev(sFolder, "\", Len(sFolder)-1))
    arrTmp = Split(sFolder, "\")
    If IsArray(arrTmp) AND UBound(arrTmp)>0 Then sName = arrTmp(UBound(arrTmp)-1)
    If sFile = "" Then sFile = Right(sKeyPath, Len(sKeyPath)-InStrRev(sKeyPath, "\"))
    If oFso.FileExists(sFolder & sLfn) Then 
        sFile = sLfn
    Else
        'Handle .cat, .manifest and .policy files
        If InStr(sLfn, ".")>0 Then
            sExt = Mid(sLfn, InStrRev(sLfn, "."))
            Select Case LCase(sExt)
            Case ".cat"
                sFile = Left(sFile, InStrRev(sFile, ".")) & "cat"
                If Not oFso.FileExists(sFolder & sFile) Then
                    'Check Manifest folder
                    If oFso.FileExists(sRoot & "Manifests\" & sName & ".cat") Then
                        sFolder = sRoot & "Manifests\"
                        sFile = sName & ".cat"
                    Else
                        If oFso.FileExists(sRoot & "Policies\" & sName & ".cat") Then
                            sFolder = sRoot & "Policies\"
                            sFile = sName & ".cat"
                        End If
                    End If
                End If
            Case ".manifest"
                sFile = Left(sFile, InStrRev(sFile, ".")) & "manifest"
                If oFso.FileExists(sRoot & "Manifests\" & sName & ".manifest") Then
                    sFolder = sRoot & "Manifests\"
                    sFile = sName & ".manifest"
                End If
            Case ".policy"
                If iVersionNT < 600 Then
                    sFile = Left(sFile, InStrRev(sFile, ".")) & "policy"
                    If oFso.FileExists(sRoot & "Policies\" & sName & ".policy") Then
                        sFolder = sRoot & "Policies\"
                        sFile = sName & ".policy"
                    End If
                Else
                    sFile = Left(sFile, InStrRev(sFile, ".")) & "manifest"
                    If oFso.FileExists(sRoot & "Manifests\" & sName & ".manifest") Then
                        sFolder = sRoot & "Manifests\"
                        sFile = sName & ".manifest"
                    End If
                End If
            Case Else
            End Select
            
        End If 'InStr(sFile, ".")>0
    End If
    
    GetAssemblyPath = sFolder & sFile
    
End Function 'GetAssemblyPath
'-------------------------------------------------------------------------------

'Routine to check if it's required to extract .msp files first
Sub CheckPatchExtract

Const COL_FILEDESCRIPTION = 34
Const COL_FILEVERSION = 145

Dim File, location
Dim iMspCnt, iExeCnt

    For Each location in arrUpdateLocations
        If NOT location = sWiCacheDir Then
            iMspCnt = 0 : iExeCnt = 0
            For Each File in oFso.GetFolder(location).Files
                If LCase(Right(File.Name, 4))=".msp" Then iMspCnt=iMspCnt+1
                If LCase(Right(File.Name, 4))=".exe" Then iExeCnt=iExeCnt+1
            Next 'File
            If (iMspCnt=0) AND (iExeCnt>0) Then
            For Each File in oFso.GetFolder(location).Files
                If LCase(Right(File.Name, 4))=".exe" Then 
                    If InStr(LCase(GetDetailsOf(File, COL_FILEDESCRIPTION)), "(kb")>0 Then
                        ExtractPatch File
                    End If
                End If
            Next 'File
            End If
        End If 'sWiCacheDir
    Next 'location

End Sub 'CheckPatchExtract
'-------------------------------------------------------------------------------

Function GetDetailsOf(File, iColumn)
    Dim oFolder, oFolderItem
    
    set oFolder = oShellApp.NameSpace(File.ParentFolder.Path)

    If (NOT oFolder Is Nothing) Then
        Set oFolderItem = oFolder.ParseName(File.Name)
        If (NOT oFolderItem Is Nothing) Then
            GetDetailsOf = oFolder.GetDetailsOf(oFolderItem, iColumn)
        End If
    End If
End Function 'GetDetailsOf
'-------------------------------------------------------------------------------

Sub ExtractPatch (File)

Dim sCmd, sReturn

On Error Resume Next

    sCmd = chr(34) & File.Path & chr(34) & " /extract:" & chr(34) & File.ParentFolder.Path & chr(34) & " /quiet" 
    sReturn = oWShell.Run(sCmd, 1, True)
    sTmp = vbTab & "Extracting patch " & File.Name & " returned: " & sReturn & " " & ExtractorRetval(sReturn)
    Log sTmp
    If fCscript Then wscript.echo sTmp
End Sub
'-------------------------------------------------------------------------------

'Query Wmi to identify local hard disks.
'The result is stored in a global dic array

Sub FindLocalDisks(dicLocalDisks)
    Const DISK_LOCAL = 3
    Dim LogicalDisks, Disk

    On Error Resume Next

    Set LogicalDisks = oWmiLocal.ExecQuery("Select * from Win32_LogicalDisk")
    For Each Disk in LogicalDisks
        If Disk.DriveType = DISK_LOCAL Then dicLocalDisks.Add Disk.DeviceID, DISK_LOCAL
    Next 'Disk

End Sub 'FindLocalDisks
'-------------------------------------------------------------------------------

Sub CreateFolderStructure (sFolder)
    Dim fld
    Dim sCreate
    Dim arrFldStruct
    
    On Error Resume Next
    arrFldStruct = Split(sFolder, "\")
    sCreate = ""
    For Each fld in arrFldStruct
        If NOT sCreate = "" Then sCreate = sCreate & "\" & fld Else sCreate = fld
        If NOT oFso.FolderExists(sCreate) Then oFso.CreateFolder sCreate
    Next
End Sub 'CreateFolderStructure
'-------------------------------------------------------------------------------

Sub Log(sLog)
    LogStream.WriteLine sLog
End Sub 'Log
'-------------------------------------------------------------------------------

Function GetRandomMspName()

    Dim sRandom
    Dim iHigh, iLow
    
    On Error Resume Next

    iHigh = 268365550
    iLow = 1048576

    Randomize
    sRandom = sWICacheDir
    sRandom = sWICacheDir & LCase(Hex((iHigh-iLow + 1) * Rnd + iLow)) & ".msp"
    While oFso.FileExists(sRandom)
        Randomize
        sRandom = sWICacheDir & LCase(Hex((iHigh-iLow + 1) * Rnd + iLow)) & ".msp"
    Wend

    GetRandomMspName = sRandom

End Function 'GetRandomMspName
'-------------------------------------------------------------------------------


'String Helper Routines
'------------------------

'-------------------------------------------------------------------------------

Sub LogSummary(sProductCode, sLog)
    On Error Resume Next
    If dicSummary.Exists(sProductCode) Then
        dicSummary.Item(sProductCode) = dicSummary.Item(sProductCode) & sLog & vbCrLf
    Else
        dicSummary.Add sProductCode, vbCrLf & sLog & vbCrLf
    End If
End Sub 'Log
'-------------------------------------------------------------------------------

'Translate the Office build to the SP level
'Possible return values are RTM, SP1, SP2, SP3 or ""
Function GetSpLevel(sBuild)

Dim arrVersionString

GetSpLevel = ""

arrVersionString = Split(sBuild, ".")
If NOT IsArray(arrVersionString) Then Exit Function
If NOT UBound(arrVersionString) > 1 Then Exit Function 'Require "major.minor.build" format

Select Case arrVersionString(0) 'BuildMajor
Case "15"
    Select Case arrVersionString(2)
    Case "4420" : GetSpLevel = " - RTM"
    Case "4569" : GetSpLevel = " - SP1"
    End Select
Case "14"
    Select Case arrVersionString(2)
    Case "4763" : GetSpLevel = " - RTM"
    Case "6029" : GetSpLevel = " - SP1"
    Case "7015" : GetSpLevel = " - SP2"
    Case "5117", "5118", "5117" : GetSpLevel = " - Web V1"
    Case Else
    End Select
Case "12"
    Select Case arrVersionString(2)
    Case "4518" : GetSpLevel = " - RTM"
    Case "6021" : GetSpLevel = " - V3"
    Case "6213", "6214", "6215", "6219", "6230", "6237" : GetSpLevel = " - SP1"
    Case "6425", "6520" : GetSpLevel = " - SP2"
    Case "6514" : GetSpLevel = " - V4 with SP2"
    Case "6612" : GetSpLevel = " - SP3"
    Case Else
    End Select
Case "11"
    Select Case arrVersionString(2)
    Case "3216", "5510", "5614" : GetSpLevel = " - RTM"
    Case "4301", "6353", "6355", "6361", "6707" : GetSpLevel = " - SP1"
    Case "7969" : GetSpLevel = " - SP2"
    Case "8173" : GetSpLevel = " - SP3"
    Case Else
    End Select
Case "10"
    Select Case arrVersionString(2)
    Case "525", "2623", "2627", "2915" : GetSpLevel = " - RTM"
    Case "2514", "3416", "3506", "3520" : GetSpLevel = " - SP1"
    Case "4128", "4219", "4330", "5110" : GetSpLevel = " - SP2"
    Case "6308", "6612", "6626" : GetSpLevel = " - SP3"
    Case Else
    End Select
Case "9"
    Select Case arrVersionString(2)
    Case "2720" : GetSpLevel = " - RTM"
    Case "3821" : GetSpLevel = " - SR1"
    Case "4527" : GetSpLevel = " - SP2"
    Case "9327" : GetSpLevel = " - SP3"
    Case Else
    End Select
Case Else
End Select

End Function
'-------------------------------------------------------------------------------

Sub ComputerProperties
    Dim oOsItem
    Dim arrVersion
    Dim qOS
    On Error Resume Next
    
    sComputerName = oWShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
    
    'OS info from WMI Win32_OperatingSystem
    Set qOS = oWmiLocal.ExecQuery("Select * from Win32_OperatingSystem")
    For Each oOsItem in qOS 
        sOSinfo = "Operating System: " & oOsItem.Caption
        sOSinfo = sOSinfo & "," & "Service Pack: SP " & oOsItem.ServicePackMajorVersion
        sOSinfo = sOSinfo & "," & "Version: " & oOsItem.Version
        sOsVersion = oOsItem.Version
        sOSinfo = sOSinfo & "," & "Codepage: " & oOsItem.CodeSet
        sOSinfo = sOSinfo & "," & "Country Code: " & oOsItem.CountryCode
        sOSinfo = sOSinfo & "," & "Language: " & oOsItem.OSLanguage
    Next
    
    'Build the VersionNT number
    arrVersion = Split(sOsVersion, Delimiter(sOsVersion))
    iVersionNt = CInt(arrVersion(0))*100 + CInt(arrVersion(1))
    
End Sub 'ComputerProperties
'-------------------------------------------------------------------------------


'Sorts an array in descending order
Function BubbleSort(arrBubble)
    
    Dim sTmp
    
    Dim iCntOuter, iCntInner
    
    On Error Resume Next
    BubbleSort = arrBubble
    If NOT IsArray(arrBubble) Then Exit Function
    
    For iCntOuter = UBound(arrBubble)-1 To 0 Step -1
        'Inner sort loop
        For iCntInner = 0 To iCntOuter
            If arrBubble(iCntInner) < arrBubble(iCntInner+1) Then
                sTmp = arrBubble(iCntInner+1)
                arrBubble(iCntInner+1) = arrBubble(iCntInner)
                arrBubble(iCntInner) = sTmp
            End If
        Next 'iCntInner
    Next 'iCntOuter
    BubbleSort = arrBubble
End Function
'-------------------------------------------------------------------------------

'Returns the delimiter of a number string
Function Delimiter (sVersion)

    Dim iCnt, iAsc

    Delimiter = " "
    For iCnt = 1 To Len(sVersion)
        iAsc = Asc(Mid(sVersion, iCnt, 1))
        If Not (iASC >= 48 And iASC <= 57) Then 
            Delimiter = Mid(sVersion, iCnt, 1)
            Exit Function
        End If
    Next 'iCnt
End Function
'-------------------------------------------------------------------------------

'Returns an array with valid folder locations.
'Local folders first (except WICacheDir), network folders, WICacheDir last
Function EnsureLocation(sLocations)
    Dim sLocation, sLocalLocations, sNetworkLocations, DynLocation
    Dim fLocal
    Dim arrLocations
    Dim dicLocalDisks

    On Error Resume Next
    
    'Find local disk drives
    Set dicLocalDisks = CreateObject ("Scripting.Dictionary")
    FindLocalDisks dicLocalDisks

    sLocalLocations = "" : sNetworkLocations = ""
    sLocations = LCase(sLocations) & ";" & LCase(sScriptDir)
    sLocations = Replace(sLocations, ",", ";")
    arrLocations = RemoveDuplicates(Split(sLocations, ";"))
    sLocations = ""
    If NOT fDynSUpdateDiscovered Then DiscoverDynSUpdateFolders
    For Each Location in arrLocations
        sLocation = ""
        sLocation = LCase(Location)
        If NOT oFso.FolderExists (sLocation) Then sLocation = GetFullPathFromRelative (sLocation)
        If oFso.FolderExists (sLocation) Then
            'Ensure trailing '\'
            If NOT Right (sLocation, 1) = "\" Then sLocation = sLocation & "\"
            fLocal = dicLocalDisks.Exists (Left (sLocation, 2))
            If fLocal _
            Then sLocalLocations = sLocalLocations & ";" & sLocation _
            Else sNetworkLocations = sNetworkLocations & ";" & sLocation
        End If
        For Each DynLocation in dicDynCultFolders.Keys
            If oFso.FolderExists (sLocation & DynLocation) Then
                If NOT Right (DynLocation, 1) = "\" Then sLocation = sLocation & "\"
                If fLocal _
                Then sLocalLocations = LCase(sLocalLocations) & ";" & LCase(sLocation) & LCase(DynLocation) _
                Else sNetworkLocations = LCase(sNetworkLocations) & ";" & LCase(sLocation) & LCase(DynLocation)
            End If
        Next 'DynLocation
    Next 'Location
    sLocations = LCase(sLocalLocations) & LCase(sNetworkLocations) & ";" & LCase(sWICacheDir)
    EnsureLocation = RemoveDuplicates (Split (Mid (sLocations, 2), ";"))

End Function 'EnsureLocation

'-------------------------------------------------------------------------------
'   GetFullPathFromRelative
'
'   Expands a relative path syntax to the full path
'-------------------------------------------------------------------------------
Function GetFullPathFromRelative (sRelativePath)
    Dim sScriptDir

    sScriptDir = Left (wscript.ScriptFullName, InStrRev (wscript.ScriptFullName, "\"))
    ' ensure sRelativePath has no leading "\"
    If Left (sRelativePath, 1) = "\" Then sRelativePath = Mid (sRelativePath, 2)
    If oFso.FolderExists (oFso.GetAbsolutePathName (sScriptDir & sRelativePath)) _
    Then GetFullPathFromRelative = oFso.GetAbsolutePathName (sScriptDir & sRelativePath) _
    Else GetFullPathFromRelative = sRelativePath

End Function 'GetFullPathFromRelative
'-------------------------------------------------------------------------------

'Remove duplicate entries from a one dimensional array
Function RemoveDuplicates(Array)
    Dim Item
    Dim oDic
    
    On Error Resume Next
    Set oDic = CreateObject("Scripting.Dictionary")
    For Each Item in Array
        If Not oDic.Exists(Item) Then oDic.Add Item, Item
    Next 'Item
    RemoveDuplicates = oDic.Keys
End Function 'RemoveDuplicates
'-------------------------------------------------------------------------------

'Converts the GUID / ProductCode into the compressed format
Function GetCompressedGuid (sGuid)
    Dim sCompGUID
    Dim i
    
    On Error Resume Next
    sCompGUID = StrReverse(Mid(sGuid, 2, 8)) & _
                StrReverse(Mid(sGuid, 11, 4)) & _
                StrReverse(Mid(sGuid, 16, 4)) 
    For i = 21 To 24
	    If i Mod 2 Then
		    sCompGUID = sCompGUID & Mid(sGuid, (i + 1), 1)
	    Else
		    sCompGUID = sCompGUID & Mid(sGuid, (i - 1), 1)
	    End If
    Next
    For i = 26 To 37
	    If i Mod 2 Then
		    sCompGUID = sCompGUID & Mid(sGuid, (i - 1), 1)
	    Else
		    sCompGUID = sCompGUID & Mid(sGuid, (i + 1), 1)
	    End If
    Next
    GetCompressedGuid = sCompGUID
    
End Function
'-------------------------------------------------------------------------------

Function GetExpandedGuid (sGuid)

    Dim sExpandGuid
    Dim i
    
    On Error Resume Next
    sExpandGuid = "{" & StrReverse(Mid(sGuid, 1, 8)) & "-" & _
                        StrReverse(Mid(sGuid, 9, 4)) & "-" & _
                        StrReverse(Mid(sGuid, 13, 4))& "-"
    For i = 17 To 20
	    If i Mod 2 Then
		    sExpandGuid = sExpandGuid & mid(sGuid, (i + 1), 1)
	    Else
		    sExpandGuid = sExpandGuid & mid(sGuid, (i - 1), 1)
	    End If
    Next
    sExpandGuid = sExpandGuid & "-"
    For i = 21 To 32
	    If i Mod 2 Then
		    sExpandGuid = sExpandGuid & mid(sGuid, (i + 1), 1)
	    Else
		    sExpandGuid = sExpandGuid & mid(sGuid, (i - 1), 1)
	    End If
    Next
    sExpandGuid = sExpandGuid & "}"
    GetExpandedGuid = sExpandGuid
    
End Function
'-------------------------------------------------------------------------------

'Translation for msiexec.exe error codes
Function MsiexecRetVal(iRetVal)
    On Error Resume Next
    Select Case iRetVal
        Case 0 : MsiexecRetVal = "Success"
        Case 1259 : MsiexecRetVal = "APPHELP_BLOCK"
        Case 1601 : MsiexecRetVal = "INSTALL_SERVICE_FAILURE"
        Case 1602 : MsiexecRetVal = "INSTALL_USEREXIT"
        Case 1603 : MsiexecRetVal = "INSTALL_FAILURE"
        Case 1604 : MsiexecRetVal = "INSTALL_SUSPEND"
        Case 1605 : MsiexecRetVal = "UNKNOWN_PRODUCT"
        Case 1606 : MsiexecRetVal = "UNKNOWN_FEATURE"
        Case 1607 : MsiexecRetVal = "UNKNOWN_COMPONENT"
        Case 1608 : MsiexecRetVal = "UNKNOWN_PROPERTY"
        Case 1609 : MsiexecRetVal = "INVALID_HANDLE_STATE"
        Case 1610 : MsiexecRetVal = "BAD_CONFIGURATION"
        Case 1611 : MsiexecRetVal = "INDEX_ABSENT"
        Case 1612 : MsiexecRetVal = "INSTALL_SOURCE_ABSENT"
        Case 1613 : MsiexecRetVal = "INSTALL_PACKAGE_VERSION"
        Case 1614 : MsiexecRetVal = "PRODUCT_UNINSTALLED"
        Case 1615 : MsiexecRetVal = "BAD_QUERY_SYNTAX"
        Case 1616 : MsiexecRetVal = "INVALID_FIELD"
        Case 1618 : MsiexecRetVal = "INSTALL_ALREADY_RUNNING"
        Case 1619 : MsiexecRetVal = "INSTALL_PACKAGE_OPEN_FAILED"
        Case 1620 : MsiexecRetVal = "INSTALL_PACKAGE_INVALID"
        Case 1621 : MsiexecRetVal = "INSTALL_UI_FAILURE"
        Case 1622 : MsiexecRetVal = "INSTALL_LOG_FAILURE"
        Case 1623 : MsiexecRetVal = "INSTALL_LANGUAGE_UNSUPPORTED"
        Case 1624 : MsiexecRetVal = "INSTALL_TRANSFORM_FAILURE"
        Case 1625 : MsiexecRetVal = "INSTALL_PACKAGE_REJECTED"
        Case 1626 : MsiexecRetVal = "FUNCTION_NOT_CALLED"
        Case 1627 : MsiexecRetVal = "FUNCTION_FAILED"
        Case 1628 : MsiexecRetVal = "INVALID_TABLE"
        Case 1629 : MsiexecRetVal = "DATATYPE_MISMATCH"
        Case 1630 : MsiexecRetVal = "UNSUPPORTED_TYPE"
        Case 1631 : MsiexecRetVal = "CREATE_FAILED"
        Case 1632 : MsiexecRetVal = "INSTALL_TEMP_UNWRITABLE"
        Case 1633 : MsiexecRetVal = "INSTALL_PLATFORM_UNSUPPORTED"
        Case 1634 : MsiexecRetVal = "INSTALL_NOTUSED"
        Case 1635 : MsiexecRetVal = "PATCH_PACKAGE_OPEN_FAILED"
        Case 1636 : MsiexecRetVal = "PATCH_PACKAGE_INVALID"
        Case 1637 : MsiexecRetVal = "PATCH_PACKAGE_UNSUPPORTED"
        Case 1638 : MsiexecRetVal = "PRODUCT_VERSION"
        Case 1639 : MsiexecRetVal = "INVALID_COMMAND_LINE"
        Case 1640 : MsiexecRetVal = "INSTALL_REMOTE_DISALLOWED"
        Case 1641 : MsiexecRetVal = "SUCCESS_REBOOT_INITIATED"
        Case 1642 : MsiexecRetVal = "PATCH_TARGET_NOT_FOUND"
        Case 1643 : MsiexecRetVal = "PATCH_PACKAGE_REJECTED"
        Case 1644 : MsiexecRetVal = "INSTALL_TRANSFORM_REJECTED"
        Case 1645 : MsiexecRetVal = "INSTALL_REMOTE_PROHIBITED"
        Case 1646 : MsiexecRetVal = "PATCH_REMOVAL_UNSUPPORTED"
        Case 1647 : MsiexecRetVal = "UNKNOWN_PATCH"
        Case 1648 : MsiexecRetVal = "PATCH_NO_SEQUENCE"
        Case 1649 : MsiexecRetVal = "PATCH_REMOVAL_DISALLOWED"
        Case 1650 : MsiexecRetVal = "INVALID_PATCH_XML"
        Case 3010 : MsiexecRetVal = "SUCCESS_REBOOT_REQUIRED"
        Case Else : MsiexecRetVal = "Unknown Return Value"
    End Select
End Function 'MsiexecRetVal
'-------------------------------------------------------------------------------

'Error codes for 2007 Office update packages (aka Microsoft Self-Extractor)
Function ExtractorRetVal(iRetVal)

    On Error Resume Next
    Select Case iRetVal
        Case 0 : ExtractorRetVal = "Success"
        Case 17301 : ExtractorRetVal = "Error: General Detection error"
        Case 17302 : ExtractorRetVal = "Error: Applying patch"
        Case 17303 : ExtractorRetVal = "Error: Extracting file"
        Case 17021 : ExtractorRetVal = "Error: Creating temp folder"
        Case 17022 : ExtractorRetVal = "Success: Reboot flag set"
        Case 17023 : ExtractorRetVal = "Error: User cancelled installation"
        Case 17024 : ExtractorRetVal = "Error: Creating folder failed"
        Case 17025 : ExtractorRetVal = "Patch already installed"
        Case 17026 : ExtractorRetVal = "Patch already installed to admin installation"
        Case 17027 : ExtractorRetVal = "Installation source requires full file update"
        Case 17028 : ExtractorRetVal = "No product installed for contained patch"
        Case 17029 : ExtractorRetVal = "Patch failed to install"
        Case 17030 : ExtractorRetVal = "Detection: Invalid CIF format"
        Case 17031 : ExtractorRetVal = "Detection: Invalid baseline"
        Case 17034 : ExtractorRetVal = "Error: Required patch does not apply to the machine"
        Case Else  : ExtractorRetVal = "Unknown Return Value"
    End Select
End Function 'ExtractorRetVal
'-------------------------------------------------------------------------------

'Get Version Major from Office GUID
Function GetVersionMajor(sProductCode)
    Dim iVersionMajor
    On Error Resume Next

    iVersionMajor = 0
        If UCase(Right(sProductCode, 28)) = OFFICE_2000 Then iVersionMajor = 9
        If UCase(Right(sProductCode, 28)) = ORK_2000 Then iVersionMajor = 9
        If UCase(Right(sProductCode, 28)) = PRJ_2000 Then iVersionMajor = 9
        If UCase(Right(sProductCode, 28)) = VIS_2002 Then iVersionMajor = 10
        If UCase(Right(sProductCode, 28)) = OFFICE_2002 Then iVersionMajor = 10
        If UCase(Right(sProductCode, 28)) = OFFICE_2003 Then iVersionMajor = 11
        If UCase(Right(sProductCode, 28)) = WSS_2 Then iVersionMajor = 11
        If UCase(Right(sProductCode, 28)) = MOSS_2003 Then iVersionMajor = 11
        If UCase(Right(sProductCode, 28)) = PPS_2007 Then iVersionMajor = 12
        If UCase(Right(sProductCode, 17)) = OFFICEID OR UCase(Right(sProductCode, 17)) = OFFICEDBGID Then iVersionMajor = Mid(sProductCode, 4, 2)

    GetVersionMajor = iVersionMajor
End Function
'-------------------------------------------------------------------------------

'Get Office Family Version from GUID
Function GetOFamilyVer(sProductCode)
    Dim iOFamilyVer
    On Error Resume Next

    iOFamilyVer = 0
    If UCase(Right(sProductCode, 28)) = OFFICE_2000 Then iOFamilyVer = 2000
    If UCase(Right(sProductCode, 28)) = ORK_2000 Then iOFamilyVer = 2000
    If UCase(Right(sProductCode, 28)) = PRJ_2000 Then iOFamilyVer = 2000
    If UCase(Right(sProductCode, 28)) = VIS_2002 Then iOFamilyVer = 2002
    If UCase(Right(sProductCode, 28)) = OFFICE_2002 Then iOFamilyVer = 2002
    If UCase(Right(sProductCode, 28)) = OFFICE_2003 Then iOFamilyVer = 2003
    If UCase(Right(sProductCode, 28)) = WSS_2 Then iOFamilyVer = 2003
    If UCase(Right(sProductCode, 28)) = MOSS_2003 Then iOFamilyVer = 2003
    If UCase(Right(sProductCode, 28)) = PPS_2007 Then iOFamilyVer = 2007
    If UCase(Right(sProductCode, 17)) = OFFICEID OR UCase(Right(sProductCode, 17)) = OFFICEDBGID Then 
        Select Case Mid(sProductCode, 4, 2)
        Case 12
            iOFamilyVer = 2007
        Case 14
            iOFamilyVer = 2010
        Case 15
            iOFamilyVer = 2013
        End Select
    End If
    GetOFamilyVer = iOFamilyVer
End Function
'-------------------------------------------------------------------------------

'Convert the TargetVersionComparisonFilter to an Int
Function GetComparisonFilter(sTvCf, arrTargetVersion)
                    
        Dim i, iCnt
        Dim sValidate

        Select Case sTvCf
        Case "None"
            iCnt = -1
        Case "Major"
            iCnt = 0
        Case "MajorMinor"
            iCnt = 1
        Case "MajorMinorUpdate"
            iCnt = 2
        Case Else
            iCnt = -2
        End Select
        If iCnt > -1 Then
            For i = 0 To iCnt
                sValidate = sValidate & "." & arrTargetVersion(i)
            Next 'iCnt
            sValidate = Mid(sValidate, 2)
        Else
            sValidate = "None"
        End If
        GetComparisonFilter = sValidate
End Function 'GetComparisonFilterCnt
'-------------------------------------------------------------------------------

Function GetProductID (sProductCode, iVM)

Dim sReturn

    Dim sProdId
	
    If iVM = "" Then iVM = GetVersionMajor(sProductCode)
	If iVM < 12 Then sProdId = Mid(sProductCode, 4, 2) Else sProdId = Mid(sProductCode, 11, 4)
	Select Case iVM
    Case 14, 15
        Select Case sProdId
        
        Case "000F" : sReturn = "Office Mondo"
        Case "0010" : sReturn = "Web Folders (Rosebud)"
        Case "0011" : sReturn = "Office Professional Plus"
        Case "0012" : sReturn = "Office Standard"
        Case "0013" : sReturn = "Office Basic"
        Case "0014" : sReturn = "Office Professional"
        Case "0015" : sReturn = "Access"
        Case "0016" : sReturn = "Excel"
        Case "0017" : sReturn = "SharePoint Designer"
        Case "0018" : sReturn = "PowerPoint"
        Case "0019" : sReturn = "Publisher"
        Case "001A" : sReturn = "Outlook"
        Case "001B" : sReturn = "Word"
        Case "001C" : sReturn = "Access Runtime"
        Case "001E" : sReturn = "Language Pack"
        Case "001F" : sReturn = "Proof"
        Case "0020" : sReturn = "Office Compatibility Pack for Word, Excel, and PowerPoint 2007 File Formats"
        Case "0021" : sReturn = "Visual Studio Web Authoring Component (Office Visual Web Developer)"
        Case "0022" : sReturn = "Office Lite"
        Case "0023" : sReturn = "Language Pack Wizard"
        Case "0024" : sReturn = "Office Resource Kit"
        Case "0025" : sReturn = "Office Resource Kit Tools"
        'Case "0026" : sReturn = "Expression Web"
        Case "0027" : sReturn = "Project"
        Case "0027" : sReturn = "IME"
        Case "0029" : sReturn = "Excel (Home and Student)"
        Case "002A" : sReturn = "Office 64-bit Components"
        Case "002B" : sReturn = "Word (Home and Studen)"
        Case "002C" : sReturn = "Proofing"
        Case "002D" : sReturn = "Fonts"
        Case "002E" : sReturn = "Office Ultimate"
        Case "002F" : sReturn = "Office Home and Student"
        Case "0028" : sReturn = "Office IME"
        Case "0030" : sReturn = "Office Enterprise"
        'Case "0031" : sReturn = "Office Professional Hybrid"
        'Case "0032" : sReturn = "Office Personal"
        Case "0033" 
            If iVM = 14 Then sReturn = "Office PIPC1" Else sReturn = "Personal"
        Case "0034" : sReturn = "Office PIPC2"
        'Case "0035" : sReturn = "Office Professional Hybrid 2"
        Case "0036" : sReturn = "Office Resource Kit Docs"
        Case "0037" : sReturn = "PowerPoint (Home and Student)"
        Case "0038" : sReturn = "Outlook TimeZone"
        Case "0039" : sReturn = "InfoPath WSS Upgrade Tool"
        Case "003A" : sReturn = "Project Standard"
        Case "003B" : sReturn = "Project Professional"
        Case "003C" : sReturn = "Office Shared Services FE"
        Case "003D" : sReturn = "Office Single Image"
        Case "003E" : sReturn = "Office OFSMUI"
        'Case "003F" : sReturn = "Excel Viewer"
        Case "0040" : sReturn = "Office Watson Website"
        Case "0041" : sReturn = "Office Watson Live Crash"
        Case "0043" : sReturn = "Office 32bit Components"
        Case "0044" : sReturn = "InfoPath"
        Case "0045" : sReturn = "Expression Web"
        Case "0046" : sReturn = "Expression Web Language Pack"
        Case "0046" : sReturn = "Plr Excel Addin"
        Case "0048" : sReturn = "Outlook Hotmail Connector"
        Case "0049" : sReturn = "OneNote Language Interface Pack"
        Case "004A" : sReturn = "Proofing Tools Kit"
        Case "004B" : sReturn = "Office Client Proofing Tools Kit"
        Case "004C" : sReturn = "InfoPath Repair Utility"
        Case "004D" : sReturn = "Conferencing"
        Case "004E" : sReturn = "Outlook Social Connector"
        Case "004F" : sReturn = "Excel XLL SDK"
        Case "0050" : sReturn = "Visio SDK"
        Case "0051" : sReturn = "Visio Professional"
        Case "0052" : sReturn = "Visio Viewer"
        Case "0053" : sReturn = "Visio Standard"
        Case "0054" : sReturn = "Visio Shared MUI"
        Case "0055" : sReturn = "Visio MUI"
        Case "0056" : sReturn = "IGX"
        Case "0057" : sReturn = "Visio"
        Case "0058" : sReturn = "Visio Ultimate"
        Case "0060" : sReturn = "Click-to-Run"
        Case "0061" : sReturn = "Office Home and Student Click-to-Run"
        Case "0062" : sReturn = "Office Essentials Click-to-Run"
        Case "0063" : sReturn = "Office Speller Download"
        Case "0064" : sReturn = "Project Professional Click-to-Run"
        Case "0065" : sReturn = "Office Test Click-to-Run"
        Case "0066" : sReturn = "Office Starter Click-to-Run"
        Case "0067" : sReturn = "Access Click-to-Run"
        Case "0068" : sReturn = "Office Standard Click-to-Run"
        Case "0069" : sReturn = "Office Professional Click-to-Run"
        Case "006A" : sReturn = "Office MUI Click-to-Run"
        Case "006B" : sReturn = "Office Mondo Full Click-to-Run"
        Case "006C" : sReturn = "Office Mondo Click-to-Run"
        Case "006D" : sReturn = "Click-to-Run"
        Case "006E" : sReturn = "Office Shared MUI"
        Case "006F" : sReturn = "Office Shared WW"
        Case "0070" : sReturn = "OOBE"
        Case "0071" : sReturn = "Excel Common UI"
        Case "0072" : sReturn = "Word Common UI"
        Case "0073" : sReturn = "App-V Redist"
        Case "0074" : sReturn = "Office Starter"
        Case "0076" : sReturn = "InfoPath Form Filler"
        Case "007A" : sReturn = "Outlook Connector"
        Case "007C" : sReturn = "Outlook Social Connector Provider for FaceBook"
        Case "007D" : sReturn = "Outlook Social Connector Provider for Windows Live Messenger"
        Case "008A" : sReturn = "Office Recent Documents Gadget"
        Case "008B" : sReturn = "Office Small Business Basics"
        Case "0090" : sReturn = "DCF MUI"
        Case "00A1" : sReturn = "OneNote"
        Case "00A3" : sReturn = "OneNote Home Student"
        Case "00A4" : sReturn = "Office 2003 Web Components"
        Case "00A7" : sReturn = "Calendar Printing Assistant for Microsoft Office Outlook"
        Case "00A9" : sReturn = "InterConnect"
        Case "00AF" : sReturn = "PowerPoint Viewer"
        Case "00B0" : sReturn = "Save as PDF Add-in"
        Case "00B1" : sReturn = "Save as XPS Add-in"
        Case "00B2" : sReturn = "Save as PDF or XPS Add-in"
        Case "00B3" : sReturn = "Project Add-in for Outlook"
        Case "00B4" : sReturn = "Project MUI"
        Case "00B5" : sReturn = "Project MUI"
        Case "00B9" : sReturn = "Application Error Reporting"
        Case "00BA" : sReturn = "Groove"
        Case "00BC" : sReturn = "InterConnect Outlook"
        Case "00C1" : sReturn = "Office 32-bit Components"
        Case "00CA" : sReturn = "Office Small Business"
        Case "00E0" : sReturn = "Outlook Standalone NoProFeatures"
        Case "00E1" : sReturn = "Osm MUI"
        Case "00E2" : sReturn = "Osm XMUI"
        Case "00D0" : sReturn = "Access Source Code Control"
        Case "00D1"
            If iVM = 14 Then sReturn = "Access Connectivity Engine ACE" Else sReturn = "Access Database Engine"
        
        Case "0100" : sReturn = "Office MUI"
        Case "0101" : sReturn = "Office XMUI"
        Case "0103" : sReturn = "Office Proofing Tools Kit"
        Case "0114" : sReturn = "Groove Setup Metadata"
        Case "0115" : sReturn = "Office Shared Setup Metadata"
        Case "0116" : sReturn = "Office Shared Setup Metadata"
        Case "0117" : sReturn = "Access Setup Metadata"
        Case "0120" : sReturn = "Office Subscription"
        Case "0126" : sReturn = "Word 2010 KB 2428677"
        Case "012C" : sReturn = "Lync"
        Case "012B" : sReturn = "Lync MUI"
        Case "012D" : sReturn = "Lync Entry"
        Case "012E" : sReturn = "Lync Vdi"
        Case "011A" : sReturn = "Send A Smile"
        Case "011D" : sReturn = "Office Professional Plus Subscription"
        Case "011F" : sReturn = "Outlook Connector"
        Case "0138" : sReturn = "Office"
        
        Case "1014" : sReturn = "SharePoint Foundation Core"
        Case "1015" : sReturn = "SharePoint Foundation Lang Pack"
        Case "1017" : sReturn = "Groove Server Manager SKU"
        Case "1018" : sReturn = "Groove Server Relay SKU"
        Case "101F" : sReturn = "Office Server Proof"
        Case "1031" : sReturn = "Project Server Web Front End"
        Case "1032" : sReturn = "Project Server Application Server"
        Case "104B" : sReturn = "SharePoint Portal Server"
        Case "104C" : sReturn = "User Profiles (Srv)"
        Case "104E" : sReturn = "SharePoint Portal Language Pack"
        Case "107F" : sReturn = "Shared Components (Srv)"
        Case "1080" : sReturn = "Shared Coms (Srv) Language Pack"
        Case "1088" : sReturn = "Slide Library"
        Case "10B0" : sReturn = "Project Server Language Pack"
        Case "10D7" : sReturn = "InfoPath Forms Services"
        Case "10D8" : sReturn = "InfoPath Forms Services Language Pack"
        Case "10EA" : sReturn = "ULS Common Core Components"
        Case "10EB" : sReturn = "Office Document Lifecycle Application Server Components"
        Case "10EC" : sReturn = "Word Server"
        Case "10ED" : sReturn = "Word Server Language Pack"
        Case "10EE" : sReturn = "PerformancePoint Services"
        Case "10F0" : sReturn = "PerformancePoint Services Language Pack"
        Case "10F1" : sReturn = "Visio Services Language Pack"
        Case "10F3" : sReturn = "Visio Services Web Front End Components"
        Case "10F5" : sReturn = "Excel Services"
        Case "10F6" : sReturn = "Excel Services Components"
        Case "10F7" : sReturn = "Document Lifecycle Components"
        Case "10F8" : sReturn = "Excel Services Language Pack"
        Case "10FB" : sReturn = "Search Server"
        Case "10FC" : sReturn = "Search"
        Case "10FD" : sReturn = "Search Server Language Pack"
        Case "1103" : sReturn = "Document Lifecycle Components Language Pack"
        Case "1106" : sReturn = "Groove Server Manager"
        Case "1107" : sReturn = "Groove Server Manager Language Pack"
        Case "1104" : sReturn = "Slide Library Language Pack"
        Case "1105" : sReturn = "Office Primary Interop Assemblies"
        Case "110D" : sReturn = "SharePoint Server"
        Case "110F" : sReturn = "Project Server"
        Case "1109" : sReturn = "Groove Server Relay"
        Case "1110" : sReturn = "SharePoint Foundation (WSS)"
        Case "1112" : sReturn = "OMUI Language Pack (Srv)"
        Case "1113" : sReturn = "XMUI Language Pack (Srv)"
        Case "1115" : sReturn = "SharePoint Foundation (WSS) Lang Pack"
        Case "1119" : sReturn = "SharePoint Foundation (WSS) Lang Pack Core"
        Case "1121" : sReturn = "SharePoint Server SDK and ECM Starter Kit"
        Case "1122" : sReturn = "Windows SharePoint Services Developer Resources"
        Case "1123" : sReturn = "Access Services Server"
        Case "1124" : sReturn = "Access Services Language Pack"
        Case "1125" : sReturn = "Web Companion Web Front End Components"
        Case "1127" : sReturn = "Web Companion Components Language Pack"
        Case "112A" : sReturn = "Web Analytics Web Front End Components"
        Case "112D" : sReturn = "Office Web Apps Server"
        Case "1131" : sReturn = "Web Analytics Language Pack"
        Case "1138" : sReturn = "Excel Mobile Viewer Components"
        Case "113E" : sReturn = "Office Web Apps Excel Mobile Viewer Components"
        Case "113F" : sReturn = "Office Web Apps Shared Components"
        Case "1139" : sReturn = "FAST Search Server"
        Case "1140" : sReturn = "Office Web Apps Shared Components Language Pack"
        Case "1141" : sReturn = "Office Web Apps Proof"
        Case "1142" : sReturn = "Excel Web App Components"
        Case "1143" : sReturn = "Excel Web App Language Pack"
        Case "1144" : sReturn = "Office Web Apps Web Front End Components"
        Case "1145" : sReturn = "Office Web Apps Language Pack"
        
        Case "2000" : sReturn = "Microsoft Filter Pack"
        Case "2005" : sReturn = "File Validation Add-In"

        Case "3000" : sReturn = "Word App"
        Case "3001" : sReturn = "Excel App"
        Case "3002" : sReturn = "PowerPoint App"
        Case "3004" : sReturn = "Outlook App"
        Case "3005" : sReturn = "OneNote App"
        Case "3006" : sReturn = "Project App"
        Case "3007" : sReturn = "Publisher App"
        Case "3008" : sReturn = "Visio App"
        Case "3010" : sReturn = "ProjectCommon App"
        Case "300A" : sReturn = "InfoPath App"
        Case "300B" : sReturn = "Access App"
        Case "300C" : sReturn = "MondoOnly App"
        Case "300D" : sReturn = "OfficeShared App"
        Case "300E" : sReturn = "Spd App"
        Case "300F" : sReturn = "Groove App"
        Case Else
        
        End Select 'sProdId
    
    Case 12
        Select Case sProdId
        
        Case "0010" : sReturn = "Web Folders"
        Case "0011" : sReturn = "Office Professional Plus"
        Case "0012" : sReturn = "Office Standard"
        Case "0013" : sReturn = "Office Basic"
        Case "0014" : sReturn = "Office Professional"
        Case "0015" : sReturn = "Access"
        Case "0016" : sReturn = "Excel"
        Case "0017" : sReturn = "SharePoint Designer"
        Case "0018" : sReturn = "PowerPoint"
        Case "0019" : sReturn = "Publisher"
        Case "001A" : sReturn = "Outlook"
        Case "001B" : sReturn = "Word"
        Case "001C" : sReturn = "Access Runtime"
        Case "001F" : sReturn = "Proof"
        Case "0020" : sReturn = "Office Compatibility Pack for Word, Excel, and PowerPoint 2007 File Formats"
        Case "0021" : sReturn = "Visual Studio Web Authoring Component (Office Visual Web Developer 2007)"
        Case "0026" : sReturn = "Expression Web"
        Case "0029" : sReturn = "Excel"
        Case "002A" : sReturn = "Office 64-bit Components"
        Case "002B" : sReturn = "Word"
        Case "002C" : sReturn = "Proofing"
        Case "002E" : sReturn = "Office Ultimate"
        Case "002F" : sReturn = "Office Home and Student"
        Case "0028" : sReturn = "Office IME"
        Case "0030" : sReturn = "Office Enterprise"
        Case "0031" : sReturn = "Office Professional Hybrid"
        Case "0032" : sReturn = "Expression Web Language Pack"
        Case "0033" : sReturn = "Office Personal"
        Case "0035" : sReturn = "Office Professional Hybrid 2007"
        Case "0038" : sReturn = "Time Zone Data Update Tool for Outlook"
        Case "0037" : sReturn = "PowerPoint"
        Case "003A" : sReturn = "Project Standard"
        Case "003B" : sReturn = "Project Professional"
        Case "003F" : sReturn = "Excel Viewer"
        Case "0043" : sReturn = "Time Zone Data Update Engine for Outlook"
        Case "0044" : sReturn = "InfoPath"
        Case "0045" : sReturn = "Expression Web 2"
        Case "0046" : sReturn = "Expression Web Language Pack"
        Case "0051" : sReturn = "Visio Professional"
        Case "0052" : sReturn = "Visio Viewer"
        Case "0053" : sReturn = "Visio Standard"
        Case "0054" : sReturn = "Visio MUI"
        Case "0055" : sReturn = "Visio MUI"
        Case "0057" : sReturn = "Visio"
        Case "006E" : sReturn = "Office Shared"
        Case "008A" : sReturn = "Office Recent Documents Gadget"
        Case "00A1" : sReturn = "OneNote"
        Case "00A3" : sReturn = "OneNote Home Student"
        Case "00A4" : sReturn = "Office 2003 Web Components"
        Case "00A7" : sReturn = "Calendar Printing Assistant for Microsoft Office Outlook"
        Case "00A9" : sReturn = "InterConnect"
        Case "00AF" : sReturn = "PowerPoint Viewer"
        Case "00B0" : sReturn = "Save as PDF Add-in for 2007 Microsoft Office programs"
        Case "00B1" : sReturn = "Save as XPS Add-in for 2007 Microsoft Office programs"
        Case "00B2" : sReturn = "Save as PDF or XPS Add-in for 2007 Microsoft Office programs"
        Case "00B3" : sReturn = "Project Add-in for Outlook"
        Case "00B4" : sReturn = "Project MUI"
        Case "00B5" : sReturn = "Project MUI"
        Case "00B9" : sReturn = "Application Error Reporting"
        Case "00BA" : sReturn = "Groove"
        Case "00BC" : sReturn = "InterConnect Outlook"
        Case "00CA" : sReturn = "Office Small Business"
        Case "00E0" : sReturn = "Outlook"
        Case "00D1" : sReturn = "Access Connectivity Engine ACE"
        Case "0100" : sReturn = "Office MUI"
        Case "0101" : sReturn = "Office XMUI"
        Case "0103" : sReturn = "Office Proofing Tools Kit"
        Case "0114" : sReturn = "Groove Setup Metadata"
        Case "0115" : sReturn = "Office Shared Setup Metadata"
        Case "0116" : sReturn = "Office Shared Setup Metadata"
        Case "0117" : sReturn = "Access Setup Metadata"
        Case "011A" : sReturn = "Windows Live Web Folder Connector"
        Case "011F" : sReturn = "Outlook Connector"
        
        Case "1014" : sReturn = "Windows SharePoint Services 3.0 (STS)"
        Case "1015" : sReturn = "Windows SharePoint Services 3.0 Lang Pack"
        Case "1032" : sReturn = "Project Server Application Server"
        Case "104B" : sReturn = "Office SharePoint Portal"
        Case "104E" : sReturn = "Office SharePoint Portal Language Pack"
        Case "107F" : sReturn = "Office Shared Components (Srv)"
        Case "1080" : sReturn = "Office Shared Coms (Srv)"
        Case "1088" : sReturn = "Office Slide Library"
        Case "10D7" : sReturn = "InfoPath Forms Services"
        Case "10D8" : sReturn = "InfoPath Forms Services Language Pack"
        Case "10EB" : sReturn = "Office Document Lifecycle Application Server Components"
        Case "10F5" : sReturn = "Excel Services"
        Case "10F6" : sReturn = "Excel Services Web Front End Components"
        Case "10F7" : sReturn = "Office Document Lifecycle Components"
        Case "10F8" : sReturn = "Excel Services Language Pack"
        Case "10FB" : sReturn = "Search Front End"
        Case "10FC" : sReturn = "Search"
        Case "10FD" : sReturn = "Search Language Pack"
        Case "1103" : sReturn = "Office Document Lifecycle Components Language Pack"
        Case "1104" : sReturn = "Office Slide Library Language Pack"
        Case "1105" : sReturn = "Office Primary Interop Assemblies"
        Case "1106" : sReturn = "Groove Management Server"
        Case "1109" : sReturn = "Groove Server Relay"
        Case "110D" : sReturn = "Office SharePoint Server (MOSS)"
        Case "110F" : sReturn = "Project Server"
        Case "1110" : sReturn = "Windows SharePoint Services 3.0 (WSS)"
        Case "1121" : sReturn = "Office SharePoint Server 2007 SDK and ECM Starter Kit"
        Case "1122" : sReturn = "Windows SharePoint Services Developer Resources 1.2"
        Case Else
        
        End Select 'sProdId
    
    Case 11
        Select Case sProdId
        
        Case "11" : sReturn = "Office Professional Enterprise"
        Case "12" : sReturn = "Office Standard"
        Case "13" : sReturn = "Office Basic"
        Case "14" : sReturn = "Windows SharePoint Services 2.0"
        Case "15" : sReturn = "Access"
        Case "16" : sReturn = "Excel"
        Case "17" : sReturn = "FrontPage"
        Case "18" : sReturn = "PowerPoint"
        Case "19" : sReturn = "Publisher"
        Case "1A" : sReturn = "Outlook Professional"
        Case "1B" : sReturn = "Word"
        Case "1C" : sReturn = "Access Runtime"
        Case "1E" : sReturn = "Office MUI"
        Case "1F" : sReturn = "Office Proofing Tools Kit"
        Case "23" : sReturn = "Office MUI"
        Case "24" : sReturn = "Office Resource Kit (ORK)"
        Case "26" : sReturn = "Office XP Web Components"
        Case "2E" : sReturn = "Office Research Service SDK"
        Case "32" : sReturn = "Project Server"
        Case "33" : sReturn = "Office Personal Edition"
        Case "3A" : sReturn = "Project Standard" 
        Case "3B" : sReturn = "Project Professional"
        Case "3C" : sReturn = "Project MUI"
        Case "44" : sReturn = "InfoPath"
        Case "48" : sReturn = "InfoPath 2003 Toolkit for Visual Studio 2005"
        Case "49" : sReturn = "Office Primary Interop Assemblies"
        Case "51" : sReturn = "Visio Professional"
        Case "52" : sReturn = "Visio Viewer"
        Case "53" : sReturn = "Visio Standard"
        Case "55" : sReturn = "Visio for Enterprise Architects"
        Case "5E" : sReturn = "Visio MUI"
        Case "83" : sReturn = "Office HTML Viewer"
        Case "84" : sReturn = "Excel Viewer"
        Case "85" : sReturn = "Word Viewer"
        Case "92" : sReturn = "Windows SharePoint Services 2.0 English Template Pack"
        Case "93" : sReturn = "Office Web Parts and Components"
        Case "A1" : sReturn = "OneNote"
        Case "A4" : sReturn = "Office Web Components"
        Case "A5" : sReturn = "SharePoint Migration Tool"
        Case "A9" : sReturn = "InterConnect 2004"
        Case "AA" : sReturn = "PowerPoint 2003 Presentation Broadcast"
        Case "AB" : sReturn = "PowerPoint 2003 Template Pack 1"
        Case "AC" : sReturn = "PowerPoint 2003 Template Pack 2"
        Case "AD" : sReturn = "PowerPoint 2003 Template Pack 3"
        Case "AE" : sReturn = "Office Organization Chart 2.0"
        Case "CA" : sReturn = "Office Small Business Edition"
        Case "D0" : sReturn = "Access Developer Extensions"
        Case "DC" : sReturn = "Office Smart Document SDK"
        Case "E0" : sReturn = "Outlook Standard"
        Case "E3" : sReturn = "Office Professional Edition (with InfoPath)"
        Case "F7" : sReturn = "InfoPath 2003 Toolkit for Visual Studio .NET"
        Case "F8" : sReturn = "Office Remove Hidden Data Tool"
        Case "FD" : sReturn = "Outlook (distributed by MSN)"
        Case "FF" : sReturn = "Office Language Interface Pack"
        Case Else : sReturn = ""
        
        End Select 'ProdId
    
    Case 10
        Select Case sProdId
        
        Case "11" : sReturn = "Office Professional"
        Case "12" : sReturn = "Office Standard"
        Case "13" : sReturn = "Office Small Business"
        Case "14" : sReturn = "Office Web Server"
        Case "15" : sReturn = "Access"
        Case "16" : sReturn = "Excel"
        Case "17" : sReturn = "FrontPage"
        Case "18" : sReturn = "PowerPoint"
        Case "19" : sReturn = "Publisher"
        Case "1A" : sReturn = "Outlook"
        Case "1B" : sReturn = "Word"
        Case "1C" : sReturn = "Access Runtime"
        Case "1D" : sReturn = "Frontpage MUI"
        Case "1E" : sReturn = "Office MUI"
        Case "1F" : sReturn = "Office Proofing Tools Kit"
        Case "20" : sReturn = "System Files Update"
        Case "23" : sReturn = "Office MUI Wizard"
        Case "24" : sReturn = "Office Resource Kit (ORK)"
        Case "25" : sReturn = "Office Resource Kit (ORK) Web Download"
        Case "26" : sReturn = "Office XP Web Components"
        Case "27" : sReturn = "Project"
        Case "28" : sReturn = "Office Professional with FrontPage"
        Case "29" : sReturn = "Office Professional Subscription"
        Case "2A" : sReturn = "Office Small Business Subscription"
        Case "2B" : sReturn = "Publisher Deluxe Edition"
        Case "2F" : sReturn = "Office Standalone IME"
        Case "30" : sReturn = "Office Media Content "
        Case "32" : sReturn = "Project Web Server"
        Case "33" : sReturn = "Office PIPC1 - Pre Installed PC"
        Case "34" : sReturn = "Office PIPC2 - Pre Installed PC"
        Case "35" : sReturn = "Office Media Content Deluxe"
        Case "3A" : sReturn = "Project Standard" 
        Case "3B" : sReturn = "Project Professional"
        Case "3C" : sReturn = "Project MUI"
        Case "3D" : sReturn = "Office Standard for Students and Teachers"
        Case "51" : sReturn = "Visio Professional"
        Case "52" : sReturn = "Visio Viewer"
        Case "53" : sReturn = "Visio Standard"
        Case "54" : sReturn = "Visio Standard"
        Case "91" : sReturn = "Office Professional"
        Case "92" : sReturn = "Office Standard"
        Case "93" : sReturn = "Office Small Business"
        Case "94" : sReturn = "Office Web Server"
        Case "95" : sReturn = "Access"
        Case "96" : sReturn = "Excel"
        Case "97" : sReturn = "FrontPage"
        Case "98" : sReturn = "PowerPoint"
        Case "99" : sReturn = "Publisher"
        Case "9A" : sReturn = "Outlook"
        Case "9B" : sReturn = "Word"
        Case "9C" : sReturn = "Access Runtime"
        Case Else : sReturn = ""
        
        End Select 'ProdId
        
    Case 9
        Select Case CInt("&h" & sProdId)
        
        Case 0 : sReturn = "Office Premium CD1"
        Case 1 : sReturn = "Office Professional"
        Case 2 : sReturn = "Office Standard"
        Case 3 : sReturn = "Office Small Business"
        Case 4 : sReturn = "Office Premium CD2"
        Case 5 : sReturn = "Office CD2 SMALL"
        Case 6 : sReturn = "Office Personal"
        Case 7 : sReturn = "Word and Excel"
        Case 16 : sReturn = "Access"
        Case 17 : sReturn = "Excel"
        Case 18 : sReturn = "FrontPage"
        Case 19 : sReturn = "PowerPoint"
        Case 20 : sReturn = "Publisher"
        Case 21 : sReturn = "Office Server Extensions"
        Case 22 : sReturn = "Outlook"
        Case 23 : sReturn = "Word"
        Case 24 : sReturn = "Access Runtime"
        Case 25 : sReturn = "FrontPage Server Extensions"
        Case 26 : sReturn = "Publisher Standalone OEM"
        Case 27 : sReturn = "DMMWeb"
        Case 28 : sReturn = "FP WECCOM"
        Case 29 : sReturn = "Word"
        Case 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47 : sReturn = "Office MUI"
        Case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63 : sReturn = "Office Proofing Tools Kit"
        Case 64 : sReturn = "Publisher Trial"
        Case 65 : sReturn = "Publisher Trial Web"
        Case 66 : sReturn = "SBB"
        Case 67 : sReturn = "SBT"
        Case 68 : sReturn = "SBT CD2"
        Case 69 : sReturn = "SBTART"
        Case 70 : sReturn = "Office Web Components"
        Case 71 : sReturn = "VP Office CD2 with LVP"
        Case 72 : sReturn = "VP PUB with LVP"
        Case 73 : sReturn = "VP PUB with LVP OEM"
        Case 79 : sReturn = "Access 2000 SR-1 Run-Time Minimum"
        Case Else : sReturn = ""
        
        End Select 'sProdId
    Case Else : sReturn = ""
    End Select 'sVersion

    If UCase(Right(sProductCode, 17)) = OFFICEDBGID AND NOT sReturn = "" Then sReturn = sReturn & " Debug"
    GetProductID = sReturn
End Function
'-------------------------------------------------------------------------------

'Gets the release type ID from the ProductCode as integer and returns a string
Function GetReleaseType (iR)
    Dim sR
    
    Select Case iR
    'Disable Case 0 to avoid noise in the output
    Case 0 : sR = "Volume License"
    Case 1 : sR = "Retail"
    Case 2 : sR = "Trial"
    Case 5 : sR = "Free"
    Case Else : sR = ""
    End Select
    
    GetReleaseType = sR
End Function
'-------------------------------------------------------------------------------


'-------------------------------------------------------------------------------

'Get the culture info tag from LCID
Function GetCultureInfo (ByVal sLcid)

Dim sLang

    If Len(sLcid) = 38 Then
	' a ProductCode has been passed in
	' handle Office ProductCodes
	    If IsOfficeProduct(sLcid) Then
            Select Case GetVersionMajor(sLcid)
            Case 9, 10, 11
                sLcid = CInt("&h" & Mid(sLcid, 6, 4))
            Case Else
                sLcid = CInt("&h" & Mid(sLcid, 16, 4))
            End Select
        End If 'IsOfficeProduct
	End If

	Select Case UCase(Hex(CInt(sLcid)))
        Case "0" : sLang = "neutral"
        Case "7F" : sLang = "invariant"  'Invariant culture
        Case "36" : sLang = "af"	 ' Afrikaans
        Case "436" : sLang = "af-ZA"	 ' Afrikaans (South Africa)
        Case "1C" : sLang = "sq"	 ' Albanian
        Case "41C" : sLang = "sq-AL"	 ' Albanian (Albania)
        Case "1" : sLang = "ar"	 ' Arabic
        Case "1401" : sLang = "ar-DZ"	 ' Arabic (Algeria)
        Case "3C01" : sLang = "ar-BH"	 ' Arabic (Bahrain)
        Case "C01" : sLang = "ar-EG"	 ' Arabic (Egypt)
        Case "801" : sLang = "ar-IQ"	 ' Arabic (Iraq)
        Case "2C01" : sLang = "ar-JO"	 ' Arabic (Jordan)
        Case "3401" : sLang = "ar-KW"	 ' Arabic (Kuwait)
        Case "3001" : sLang = "ar-LB"	 ' Arabic (Lebanon)
        Case "1001" : sLang = "ar-LY"	 ' Arabic (Libya)
        Case "1801" : sLang = "ar-MA"	 ' Arabic (Morocco)
        Case "2001" : sLang = "ar-OM"	 ' Arabic (Oman)
        Case "4001" : sLang = "ar-QA"	 ' Arabic (Qatar)
        Case "401" : sLang = "ar-SA"	 ' Arabic (Saudi Arabia)
        Case "2801" : sLang = "ar-SY"	 ' Arabic (Syria)
        Case "1C01" : sLang = "ar-TN"	 ' Arabic (Tunisia)
        Case "3801" : sLang = "ar-AE"	 ' Arabic (U.A.E.)
        Case "2401" : sLang = "ar-YE"	 ' Arabic (Yemen)
        Case "2B" : sLang = "hy"	 ' Armenian
        Case "42B" : sLang = "hy-AM"	 ' Armenian (Armenia)
        Case "2C" : sLang = "az"	 ' Azeri
        Case "82C" : sLang = "az-Cyrl-AZ"	 ' Azeri (Azerbaijan, Cyrillic)
        Case "42C" : sLang = "az-Latn-AZ"	 ' Azeri (Azerbaijan, Latin)
        Case "2D" : sLang = "eu"	 ' Basque
        Case "42D" : sLang = "eu-ES"	 ' Basque (Basque)
        Case "23" : sLang = "be"	 ' Belarusian
        Case "423" : sLang = "be-BY"	 ' Belarusian (Belarus)
        Case "2" : sLang = "bg"	 ' Bulgarian
        Case "402" : sLang = "bg-BG"	 ' Bulgarian (Bulgaria)
        Case "3" : sLang = "ca"	 ' Catalan
        Case "403" : sLang = "ca-ES"	 ' Catalan (Catalan)
        Case "C04" : sLang = "zh-HK"	 ' Chinese (Hong Kong SAR, PRC)
        Case "1404" : sLang = "zh-MO"	 ' Chinese (Macao SAR)
        Case "804" : sLang = "zh-CN"	 ' Chinese (PRC)
        Case "4" : sLang = "zh-Hans"	 ' Chinese (Simplified)
        Case "1004" : sLang = "zh-SG"	 ' Chinese (Singapore)
        Case "404" : sLang = "zh-TW"	 ' Chinese (Taiwan)
        Case "7C04" : sLang = "zh-Hant"	 ' Chinese (Traditional)
        Case "1A" : sLang = "hr"	 ' Croatian
        Case "41A" : sLang = "hr-HR"	 ' Croatian (Croatia)
        Case "5" : sLang = "cs"	 ' Czech
        Case "405" : sLang = "cs-CZ"	 ' Czech (Czech Republic)
        Case "6" : sLang = "da"	 ' Danish
        Case "406" : sLang = "da-DK"	 ' Danish (Denmark)
        Case "65" : sLang = "dv"	 ' Divehi
        Case "465" : sLang = "dv-MV"	 ' Divehi (Maldives)
        Case "13" : sLang = "nl"	 ' Dutch
        Case "813" : sLang = "nl-BE"	 ' Dutch (Belgium)
        Case "413" : sLang = "nl-NL"	 ' Dutch (Netherlands)
        Case "9" : sLang = "en"	 ' English
        Case "C09" : sLang = "en-AU"	 ' English (Australia)
        Case "2809" : sLang = "en-BZ"	 ' English (Belize)
        Case "1009" : sLang = "en-CA"	 ' English (Canada)
        Case "2409" : sLang = "en-029"	 ' English (Caribbean)
        Case "1809" : sLang = "en-IE"	 ' English (Ireland)
        Case "2009" : sLang = "en-JM"	 ' English (Jamaica)
        Case "1409" : sLang = "en-NZ"	 ' English (New Zealand)
        Case "3409" : sLang = "en-PH"	 ' English (Philippines)
        Case "1C09" : sLang = "en-ZA"	 ' English (South Africa
        Case "2C09" : sLang = "en-TT"	 ' English (Trinidad and Tobago)
        Case "809" : sLang = "en-GB"	 ' English (United Kingdom)
        Case "409" : sLang = "en-US"	 ' English (United States)
        Case "3009" : sLang = "en-ZW"	 ' English (Zimbabwe)
        Case "25" : sLang = "et"	 ' Estonian
        Case "425" : sLang = "et-EE"	 ' Estonian (Estonia)
        Case "38" : sLang = "fo"	 ' Faroese
        Case "438" : sLang = "fo-FO"	 ' Faroese (Faroe Islands)
        Case "29" : sLang = "fa"	 ' Farsi
        Case "429" : sLang = "fa-IR"	 ' Farsi (Iran)
        Case "B" : sLang = "fi"	 ' Finnish
        Case "40B" : sLang = "fi-FI"	 ' Finnish (Finland)
        Case "C" : sLang = "fr"	 ' French
        Case "80C" : sLang = "fr-BE"	 ' French (Belgium)
        Case "C0C" : sLang = "fr-CA"	 ' French (Canada)
        Case "40C" : sLang = "fr-FR"	 ' French (France)
        Case "140C" : sLang = "fr-LU"	 ' French (Luxembourg)
        Case "180C" : sLang = "fr-MC"	 ' French (Monaco)
        Case "100C" : sLang = "fr-CH"	 ' French (Switzerland)
        Case "56" : sLang = "gl"	 ' Galician
        Case "456" : sLang = "gl-ES"	 ' Galician (Spain)
        Case "37" : sLang = "ka"	 ' Georgian
        Case "437" : sLang = "ka-GE"	 ' Georgian (Georgia)
        Case "7" : sLang = "de"	 ' German
        Case "C07" : sLang = "de-AT"	 ' German (Austria)
        Case "407" : sLang = "de-DE"	 ' German (Germany)
        Case "1407" : sLang = "de-LI"	 ' German (Liechtenstein)
        Case "1007" : sLang = "de-LU"	 ' German (Luxembourg)
        Case "807" : sLang = "de-CH"	 ' German (Switzerland)
        Case "8" : sLang = "el"	 ' Greek
        Case "408" : sLang = "el-GR"	 ' Greek (Greece)
        Case "47" : sLang = "gu"	 ' Gujarati
        Case "447" : sLang = "gu-IN"	 ' Gujarati (India)
        Case "D" : sLang = "he"	 ' Hebrew
        Case "40D" : sLang = "he-IL"	 ' Hebrew (Israel)
        Case "39" : sLang = "hi"	 ' Hindi
        Case "439" : sLang = "hi-IN"	 ' Hindi (India)
        Case "E" : sLang = "hu"	 ' Hungarian
        Case "40E" : sLang = "hu-HU"	 ' Hungarian (Hungary)
        Case "F" : sLang = "is"	 ' Icelandic
        Case "40F" : sLang = "is-IS"	 ' Icelandic (Iceland)
        Case "21" : sLang = "id"	 ' Indonesian
        Case "421" : sLang = "id-ID"	 ' Indonesian (Indonesia)
        Case "10" : sLang = "it"	 ' Italian
        Case "410" : sLang = "it-IT"	 ' Italian (Italy)
        Case "810" : sLang = "it-CH"	 ' Italian (Switzerland)
        Case "11" : sLang = "ja"	 ' Japanese
        Case "411" : sLang = "ja-JP"	 ' Japanese (Japan)
        Case "4B" : sLang = "kn"	 ' Kannada
        Case "44B" : sLang = "kn-IN"	 ' Kannada (India)
        Case "3F" : sLang = "kk"	 ' Kazakh
        Case "43F" : sLang = "kk-KZ"	 ' Kazakh (Kazakhstan)
        Case "57" : sLang = "kok"	 ' Konkani
        Case "457" : sLang = "kok-IN"	 ' Konkani (India)
        Case "12" : sLang = "ko"	 ' Korean
        Case "412" : sLang = "ko-KR"	 ' Korean (Korea)
        Case "40" : sLang = "ky"	 ' Kyrgyz
        Case "440" : sLang = "ky-KG"	 ' Kyrgyz (Kyrgyzstan)
        Case "26" : sLang = "lv"	 ' Latvian
        Case "426" : sLang = "lv-LV"	 ' Latvian (Latvia)
        Case "27" : sLang = "lt"	 ' Lithuanian
        Case "427" : sLang = "lt-LT"	 ' Lithuanian (Lithuania)
        Case "2F" : sLang = "mk"	 ' Macedonian
        Case "42F" : sLang = "mk-MK"	 ' Macedonian (Macedonia, FYROM)
        Case "3E" : sLang = "ms"	 ' Malay
        Case "83E" : sLang = "ms-BN"	 ' Malay (Brunei Darussalam)
        Case "43E" : sLang = "ms-MY"	 ' Malay (Malaysia)
        Case "4E" : sLang = "mr"	 ' Marathi
        Case "44E" : sLang = "mr-IN"	 ' Marathi (India)
        Case "50" : sLang = "mn"	 ' Mongolian
        Case "450" : sLang = "mn-MN"	 ' Mongolian (Mongolia)
        Case "14" : sLang = "no"	 ' Norwegian
        Case "414" : sLang = "nb-NO"	 ' Norwegian (Bokmål, Norway)
        Case "814" : sLang = "nn-NO"	 ' Norwegian (Nynorsk, Norway)
        Case "15" : sLang = "pl"	 ' Polish
        Case "415" : sLang = "pl-PL"	 ' Polish (Poland)
        Case "16" : sLang = "pt"	 ' Portuguese
        Case "416" : sLang = "pt-BR"	 ' Portuguese (Brazil)
        Case "816" : sLang = "pt-PT"	 ' Portuguese (Portugal)
        Case "46" : sLang = "pa"	 ' Punjabi
        Case "446" : sLang = "pa-IN"	 ' Punjabi (India)
        Case "18" : sLang = "ro"	 ' Romanian
        Case "418" : sLang = "ro-RO"	 ' Romanian (Romania)
        Case "19" : sLang = "ru"	 ' Russian
        Case "419" : sLang = "ru-RU"	 ' Russian (Russia)
        Case "4F" : sLang = "sa"	 ' Sanskrit
        Case "44F" : sLang = "sa-IN"	 ' Sanskrit (India)
        Case "C1A" : sLang = "sr-Cyrl-CS"	 ' Serbian (Serbia, Cyrillic)
        Case "81A" : sLang = "sr-Latn-CS"	 ' Serbian (Serbia, Latin)
        Case "1B" : sLang = "sk"	 ' Slovak
        Case "41B" : sLang = "sk-SK"	 ' Slovak (Slovakia)
        Case "24" : sLang = "sl"	 ' Slovenian
        Case "424" : sLang = "sl-SI"	 ' Slovenian (Slovenia)
        Case "A" : sLang = "es"	 ' Spanish
        Case "2C0A" : sLang = "es-AR"	 ' Spanish (Argentina)
        Case "400A" : sLang = "es-BO"	 ' Spanish (Bolivia)
        Case "340A" : sLang = "es-CL"	 ' Spanish (Chile)
        Case "240A" : sLang = "es-CO"	 ' Spanish (Colombia)
        Case "140A" : sLang = "es-CR"	 ' Spanish (Costa Rica)
        Case "1C0A" : sLang = "es-DO"	 ' Spanish (Dominican Republic)
        Case "300A" : sLang = "es-EC"	 ' Spanish (Ecuador)
        Case "440A" : sLang = "es-SV"	 ' Spanish (El Salvador)
        Case "100A" : sLang = "es-GT"	 ' Spanish (Guatemala)
        Case "480A" : sLang = "es-HN"	 ' Spanish (Honduras)
        Case "80A" : sLang = "es-MX"	 ' Spanish (Mexico)
        Case "4C0A" : sLang = "es-NI"	 ' Spanish (Nicaragua)
        Case "180A" : sLang = "es-PA"	 ' Spanish (Panama)
        Case "3C0A" : sLang = "es-PY"	 ' Spanish (Paraguay)
        Case "280A" : sLang = "es-PE"	 ' Spanish (Peru)
        Case "500A" : sLang = "es-PR"	 ' Spanish (Puerto Rico)
        Case "C0A" : sLang = "es-ES"	 ' Spanish (Spain)
        Case "380A" : sLang = "es-UY"	 ' Spanish (Uruguay)
        Case "200A" : sLang = "es-VE"	 ' Spanish (Venezuela)
        Case "41" : sLang = "sw"	 ' Swahili
        Case "441" : sLang = "sw-KE"	 ' Swahili (Kenya)
        Case "1D" : sLang = "sv"	 ' Swedish
        Case "81D" : sLang = "sv-FI"	 ' Swedish (Finland)
        Case "41D" : sLang = "sv-SE"	 ' Swedish (Sweden)
        Case "5A" : sLang = "syr"	 ' Syriac
        Case "45A" : sLang = "syr-SY"	 ' Syriac (Syria)
        Case "49" : sLang = "ta"	 ' Tamil
        Case "449" : sLang = "ta-IN"	 ' Tamil (India)
        Case "44" : sLang = "tt"	 ' Tatar
        Case "444" : sLang = "tt-RU"	 ' Tatar (Russia)
        Case "4A" : sLang = "te"	 ' Telugu
        Case "44A" : sLang = "te-IN"	 ' Telugu (India)
        Case "1E" : sLang = "th"	 ' Thai
        Case "41E" : sLang = "th-TH"	 ' Thai (Thailand)
        Case "1F" : sLang = "tr"	 ' Turkish
        Case "41F" : sLang = "tr-TR"	 ' Turkish (Turkey)
        Case "22" : sLang = "uk"	 ' Ukrainian
        Case "422" : sLang = "uk-UA"	 ' Ukrainian (Ukraine)
        Case "20" : sLang = "ur"	 ' Urdu
        Case "420" : sLang = "ur-PK"	 ' Urdu (Pakistan)
        Case "43" : sLang = "uz"	 ' Uzbek
        Case "843" : sLang = "uz-Cyrl-UZ"	 ' Uzbek (Uzbekistan, Cyrillic)
        Case "443" : sLang = "uz-Latn-UZ"	 ' Uzbek (Uzbekistan, Latin)
        Case "2A" : sLang = "vi"	 ' Vietnamese
        Case "42A" : sLang = "vi-VN"	 ' Vietnamese (Vietnam)
    Case Else : sLang = ""
    End Select
    GetCultureInfo = sLang
End Function
'-------------------------------------------------------------------------------


