- In a text editor modify the Hotfix_Installer\Hotfix_Installer.vdproj and adjust the following values as needed
 - "ProjectName"
 - "ProductName"
 - "Title"
 - "DefaultLocation"
 - "OutputFilename" (Debug and Release)
 - *** IMPORTANT ***
    - Set "ProductCode", "PackageCode" and "UpgradeCode" to a new GUID
      every time a new MSI patch is generated so that the installation of the patches won't
      fail with the message: "Another version of this product is already installed..."

- In Visual Studio, add any files/folders needed to the "Application Folder" and/or any other folders as needed

- After the MSI has been created, change its logo and application icon by following these steps:
 - Open the MSI in Orca (if necessary install it from S:\Development and QA\Non-MSDN\Orca\Orca.msi)
 - On the Tables panel on the left, select Binary
 - On the panel on the left double click [Binary Data] for DefBannerBitmap
 - Browse for the new logo in C:\Source\Hotfixes\Hotfix_Installer\Logos and icons\Advanced_Banner_Smaller.jpg
 - On the Tables panel on the left, select Icon
 - Right click on the empty panel on the right and select Add row
 - Enter "Data" (without the quotes) for Name
 - Click on Data (the row under "Name") and then browse for the icon file in C:\Source\Hotfixes\Hotfix_Installer\Logos and icons\OpenHR Fat clients icon.ico
 - Right click again and select Add row
 - Enter "Name" (without the quotes) for Name
 - Click on Data (the row under "Name") and then enter the name of the file selected above (OpenHR Fat clients icon.ico)
 - Accept all the prompts and save the MSI
 
- Last thing to do: Sign the MSI using "C:\Program Files (x86)\Windows Kits\10\bin\x86\signtool" sign /f "C:\Source\Web Apps\DMI.NET\OpenHR Installer\AdvancedPrivateCertificate.pfx" /p fowler /t http://timestamp.verisign.com/scripts/timstamp.dll /d "HOTFIX.msi" "HOTFIX.msi"
