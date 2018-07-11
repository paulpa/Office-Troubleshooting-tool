# Office-Troubleshooting-tool
=================================

Tool to fix office /SharePoint Patching issues 

Functionality of the Tool 
===========================
We can review the missing patches by running directly on the effected server box

We can review the logs from the other servers as well.

We have tools consolidated in a UI based which can make a life of the engineer easy

We can do the following 
=====================
1. Review the installation status 
2. We can easily find out if there are any missing .MSI and .MSP files from the windows installer folder 
3. We can get the KB /Hotfix number which is missing from the windows installer cache.
4. we can Fix the Patches 
  a. Repair Cache
  b. Apply Patch
  c. Build Installer cache from a good server source
  d. We can uninstall a Patch
  e. We can Reconcile Cache
  f. Clean the installer Cache
  
5. we can Enable and Disable Verbose logging for advance troubleshooting. 


How to Use the Tool
===================
1. Download the Release.zip on the computer on which you want to troubleshoot Patching issue
2. Extract the Release.zip on the server
3. Navigate to the \Release\ folder and double click on the OfficeSetupTroubleshooting.exe file which will open up the Application
4. In the "OfficeSetupTroubleshooting Application" Select the Data menu and select "check the Missing Patches and MSI" option
4. If there are any missing patches from the installer cache you will see the list of them with the KB /Hotfix numbers as below 

========================================================================================================================================
++++++++++++++++++++++++++++++++++++
++++++++++++Missing KBs+++++++++++++
++++++++++++++++++++++++++++++++++++
	Error: Local .msp package missing. Attempt failed to restore 'c:\windows\installer\9e798c5.msp' - '{5E642F23-B29D-4589-80CF-D25E8D90C8DA}' - 'Security Update for Microsoft Office 2013 (KB2726958) 64-Bit Edition

	Error: Local .msp package missing. Attempt failed to restore 'c:\windows\installer\9e798cb.msp' - '{44344941-19D2-4DB6-9B95-2ABC89263118}' - 'Security Update for Microsoft SharePoint Designer 2013 (KB2863836) 64-Bit Edition

	Error: Local .msp package missing. Attempt failed to restore 'c:\windows\installer\9e798ef.msp' - '{594F3F8D-DB6B-4D1A-ACA1-76B79FD564AA}' - 'Security Update for Microsoft Office 2013 (KB2768005) 64-Bit Edition

	Error: Local .msp package missing. Attempt failed to restore 'c:\windows\installer\a509140.msp' - '{5414D5E2-E2B0-4B9F-BE46-CDEAC37EB089}' - 'Security Update for Microsoft SharePoint Designer 2013 (KB2752096) 64-Bit Edition

	Error: Local .msp package missing. Attempt failed to restore 'c:\windows\installer\9e798ba.msp' - '{50B38E0B-474B-435D-97C2-153D5737152C}' - 'Security Update for Microsoft Office 2013 (KB2880463) 64-Bit Edition

++++++++++++++++++++++++++++++++++++
++++++++++Missing MSIs++++++++++++++
++++++++++++++++++++++++++++++++++++
Error:                          Product {90150000-00C1-0000-1000-0000000FF1CE} - Microsoft Office 32-bit Components 2013: Local cached .msi appears to be missing: C:\WINDOWS\Installer\38eb75.msi
Error:                          Product {90150000-00C1-0409-1000-0000000FF1CE} - Microsoft Office Shared 32-bit MUI (English) 2013: Local cached .msi appears to be missing: C:\WINDOWS\Installer\38eb56.msi
Error:                          Product {90150000-0115-0409-1000-0000000FF1CE} - Microsoft Office Shared Setup Metadata MUI (English) 2013: Local cached .msi appears to be missing: C:\WINDOWS\Installer\38eb51.msi
Error:                          Product {90150000-0017-0000-1000-0000000FF1CE} - Microsoft SharePoint Designer 2013: Local cached .msi appears to be missing: C:\WINDOWS\Installer\38ebc6.msi
Error:                          Product {90150000-0017-0409-1000-0000000FF1CE} - Microsoft SharePoint Designer MUI (English) 2013: Local cached .msi appears to be missing: C:\WINDOWS\Installer\38eb6f.msi
Error:                          Product {90150000-002C-0409-1000-0000000FF1CE} - Microsoft Office Proofing (English) 2013: Local cached .msi appears to be missing: C:\WINDOWS\Installer\38eb6a.msi

========================================================================================================================================
To the fix this issue we have click the "Fix the Patches" Tab
and we have the option to 

  a. Repair Cache
  b. Apply Patch
  c. Build Installer cache from a good server source
  d. We can uninstall a Patch
  e. We can Reconcile Cache
  f. Clean the installer Cache
  
5. we can Enable and Disable Verbose logging for advance troubleshooting. 

If you have any follow up questions or comments , Please write to Paulpa@microsoft.com
