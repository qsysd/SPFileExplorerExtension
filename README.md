This shell extension is emulating

Adds "Sharepoint" menu to right-click context menu in Windows File Explorer, with:

<img src="https://user-images.githubusercontent.com/56280244/67110242-8e0aa800-f1d2-11e9-8ebd-2d89f9964b3f.jpg" width="60%"></img> 

- Copy Link - returns link to the document in DocID format (Document ID) or link to the folder in Path format
- Copy Path - returns link to the document in Path format
- Open in web browser - Opens the location in Sharepoint Web Explorer, using default browser.

Required packages (Install-Package):
SharpShell
SharpShellTools
ServerRegistrationManager
Sharepoint.Client
Sharepoint.Client.Runtime

Testing the class library can be made using "ServerManager.exe" by installing and registering the server (x64) and then trying the right-click on objects in J,K,L drives.

The program assumes that J, K,and L are mapped drives to Sharepoint document libraries, and "Sharepoint" menu is shown only on them. It can be changed to any other mapped drives, and any other URLs. (i.e. I've put contoso.com).

Tested with Windows 10 and Sharepoint 2019. Probably works with Sharepoint 2016 too, because of version of the Sharepoint.Client (version 15).
