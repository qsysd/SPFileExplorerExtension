Adds "Sharepoint" menu to right-click context menu in Windows File Explorer, with:
- Copy Link - returns link to the document in DocID format (Document ID)
- Copy Path - returns link to the document in path format
- Open in web browser - Opens the current location in Sharepoint's Web Explorer, using default browser.

Required packages to add from nuget console (Install-Package):
SharpShell
SharpShellTools
ServerRegistrationManager
Sharepoint.Client
Sharepoint.Client.Runtime

Testing the class library can be made using "ServerManager.exe" by installing  and registering the server (x64) and then trying the right-click copntext menu on objects in Sharepoint mapped drives.

The program assumes that J, K,and L are our mapped drives to Sharepoint document libraries, and "Sharepoint" menu is shown only on them. It can be changed to any other mapped drives, and any other URLs. (i.e. I've put contoso.com).

Tested with Windows 10 and Sharepoint 2019. Probably works with Sharepoint 2016 too, because of version of the Sharepoint.Client (version 15).
