This shell extension is adding "Copy Link" functionality of Sharepoint web explorer 
<img src="https://user-images.githubusercontent.com/56280244/67110705-a0391600-f1d3-11e9-8e4c-25df4b0889d7.JPG" width="40%"></img> 

into Windows File explorer.

<img src="https://user-images.githubusercontent.com/56280244/67110242-8e0aa800-f1d2-11e9-8ebd-2d89f9964b3f.jpg" width="60%"></img> 

Works on files and folders.

Menu options descriptions are:

- Copy Link - returns link to the document in DocID format (Document ID) or link to the folder in Path format (emulates behavior of web explorer)
- Copy Path - returns link to the document in Path format
- Open location in web browser - Opens the location in Sharepoint Web Explorer, using default browser.

Required nu get packages (install them with Install-Package):
SharpShell
SharpShellTools
ServerRegistrationManager
Sharepoint.Client
Sharepoint.Client.Runtime

Testing can be done using "ServerManager.exe" by installing and registering the server (x64) and trying it on objects on J, K or L mapped drives.

It assumes that J, K,and L are Sharepoint mapped network drives and menu is shown only on them. This can be changed (very simply) to any other mapped drives, and any other URLs. (i.e. I've put contoso.com).

Tested with Windows 10 and Sharepoint 2019. Should also work with Sharepoint 2016, because of version 15 of Sharepoint.Client.

Further info on making the installation msi file with using srm.exe, can be found on this link:  https://www.codeproject.com/Articles/653780/NET-Shell-Extensions-Deploying-SharpShell-Servers
