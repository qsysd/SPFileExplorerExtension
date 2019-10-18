using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using SharpShell.Attributes;
using SharpShell.SharpContextMenu;
using SP = Microsoft.SharePoint.Client;

// <summary>
// SPFileExplorerExtension is an shell context menu extension,
// implemented with SharpShell. 
// </summary>
[ComVisible(true)]
[COMServerAssociation(AssociationType.AllFiles)]
[COMServerAssociation(AssociationType.Directory)]
public class SPFEExtension : SharpContextMenu
{
    //  let's create the menu strip.
    private ContextMenuStrip menu = new ContextMenuStrip();

    // <summary>
    // Determines whether the menu item can be shown for the selected item.
    // </summary>
    // <returns>
    //   <c>true</c> if item can be shown for the selected item for this instance.; 
    // otherwise, <c>false</c>.
    // </returns>
    protected override bool CanShowMenu()
    {
        //  We can show the item only for a single selection.
        //  Make it only show for J, K and L mapped drives.
        var MappedDriveCheck = SelectedItemPaths.First().ToString().Substring(0, 1);

        if (SelectedItemPaths.Count() == 1 && (MappedDriveCheck == "J" || MappedDriveCheck == "K" || MappedDriveCheck == "L"))
        // if (SelectedItemPaths.Count() == 1)
        {
            this.UpdateMenu();
            return true;
        }
        else
        {
           return false;
        }
    }

    // <summary>
    // Creates the context menu. This can be a single menu item or a tree of them.
    // Here we create the menu based on the type of item
    // </summary>
    // <returns>
    // The context menu for the shell context menu.
    // </returns>
    protected override ContextMenuStrip CreateMenu()
    {
        menu.Items.Clear();
        FileAttributes attr = File.GetAttributes(SelectedItemPaths.First());

        // check if the selected item is a directory
        if (attr.HasFlag(FileAttributes.Directory))
        {
            this.MenuDirectory();
        }
        else
        {
            this.MenuFiles();
        }

        // return the menu item
        return menu;
    }

    // <summary>
    // Updates the context menu. 
    // </summary>
    private void UpdateMenu()
    {
        // release all resources associated to existing menu
        menu.Dispose();
        menu = CreateMenu();
    }

    // <summary>
    // Creates the context menu when the selected item is a folder.
    // </summary>
    protected void MenuDirectory()
    {
        ToolStripMenuItem MainMenu;
        MainMenu = new ToolStripMenuItem
        {
            Text = "Sharepoint",
            Image = SPFileExplorerExtension.Properties.Resources.Sp_icon.ToBitmap()
        };

        ToolStripMenuItem SubMenu1;
        SubMenu1 = new ToolStripMenuItem
        {
            Text = "Copy link",
            Image = SPFileExplorerExtension.Properties.Resources.Sp_icon.ToBitmap()
        };
        SubMenu1.DropDownItems.Clear();
        SubMenu1.Click += (sender, args) => ShowItemPathLink();

        var SubMenu2 = new ToolStripMenuItem
        {
            Text = "Open in web browser",
            Image = SPFileExplorerExtension.Properties.Resources.Sp_icon.ToBitmap()
        };
        SubMenu2.DropDownItems.Clear();
        SubMenu2.Click += (sender, args) => OpenLocationInBrowser();

        // Let's attach the submenus to the main menu
        MainMenu.DropDownItems.Add(SubMenu1);
        MainMenu.DropDownItems.Add(SubMenu2);

        menu.Items.Clear();
        menu.Items.Add(MainMenu);
    }

    // <summary>
    // Creates the context menu when the selected item is of type "file".
    // </summary>
    protected void MenuFiles()
    {
        ToolStripMenuItem MainMenu;
        MainMenu = new ToolStripMenuItem
        {
            Text = "Sharepoint",
            Image = SPFileExplorerExtension.Properties.Resources.Sp_icon.ToBitmap()
        };

        ToolStripMenuItem SubMenu3;
        SubMenu3 = new ToolStripMenuItem
        {
            Text = "Copy link",
            Image = SPFileExplorerExtension.Properties.Resources.Sp_icon.ToBitmap()
        };
        SubMenu3.DropDownItems.Clear();
        SubMenu3.Click += (sender, args) => ShowItemPermaLink();

        var SubMenu4 = new ToolStripMenuItem
        {
            Text = "Copy path",
            Image = SPFileExplorerExtension.Properties.Resources.Sp_icon.ToBitmap()
        };
        SubMenu4.DropDownItems.Clear();
        SubMenu4.Click += (sender, args) => ShowItemPathLink();

        var SubMenu5 = new ToolStripMenuItem
        {
            Text = "Open location in web browser",
            Image = SPFileExplorerExtension.Properties.Resources.Sp_icon.ToBitmap()
        };
        SubMenu5.DropDownItems.Clear();
        SubMenu5.Click += (sender, args) => OpenLocationInBrowser();

        // Let's attach the submenus to the main menu
        MainMenu.DropDownItems.Add(SubMenu3);
        MainMenu.DropDownItems.Add(SubMenu4);
        MainMenu.DropDownItems.Add(SubMenu5);

        menu.Items.Clear();
        menu.Items.Add(MainMenu);
    }

    // <summary>
    // Shows name of selected files.
    // </summary>
    private void ShowItemPermaLink()
    {
        //  Builder for the output
        var builder = new StringBuilder();
        //
        builder.AppendLine(string.Format("{0}", Path.GetFullPath(SelectedItemPaths.First())));
        // make link
        ConvertPathToURL(ref builder);
        //
        GetDocIDURL(ref builder);
        //  Show the ouput.
        Clipboard.SetText(builder.ToString());
        string outputMessage = string.Format("{0}{1}{1}copied to clipboard", builder.ToString(), Environment.NewLine);
        MessageBox.Show(outputMessage);
    }
    private void ShowItemPathLink()
    {
        //  Builder for the output
        var builder = new StringBuilder();
        //
        FileAttributes attr = File.GetAttributes(SelectedItemPaths.First());
        //  check if selected item is a directory.
        if (attr.HasFlag(FileAttributes.Directory))
        {
            //  
            builder.AppendLine(string.Format("{0}", Path.GetFullPath(SelectedItemPaths.First())));
        }
        else
        {
            // 
            builder.AppendLine(string.Format("{0}?Web=1", Path.GetFullPath(SelectedItemPaths.First())));
        }
        // make link
        ConvertPathToURL(ref builder);
        //  Show the output
        Clipboard.SetText(builder.ToString());
        string outputMessage = string.Format("{0}{1}copied to clipboard", builder.ToString(), Environment.NewLine);
        MessageBox.Show(outputMessage);
    }

    /// <summary>
    /// Opens selected item's location in Sharepoint web document library
    /// </summary>
    private void OpenLocationInBrowser()
    {
        //  Builder for the output.
        var builder = new StringBuilder();
        FileAttributes attr = File.GetAttributes(SelectedItemPaths.First());
        //  check if selected item is a directory.
        if (attr.HasFlag(FileAttributes.Directory))
        {
            //  Show folder name.
            builder.AppendLine(string.Format("{0}", Path.GetFullPath(SelectedItemPaths.First())));
        }
        else
        {
            //  Show folder name.
            builder.AppendLine(string.Format("{0}", Path.GetFullPath(SelectedItemPaths.First()).Replace(Path.GetFileName(SelectedItemPaths.First()), "")));
        }
        // make link
        ConvertPathToURL(ref builder);
        // open in default web browser
        Process proc = new Process();
        proc.StartInfo.UseShellExecute = true;
        proc.StartInfo.FileName = builder.ToString();
        proc.Start();
    }



    private void GetDocIDURL(ref StringBuilder permaurl)
    {
        string ctxurl = permaurl.ToString(0, 24);
        string relurl = permaurl.Replace("https://sp.contoso.com", "").ToString();

        // Starting with ClientContext, the constructor requires a URL to the 
        // server running SharePoint. 
        SP.ClientContext context = new SP.ClientContext(ctxurl);
        // The SharePoint web at the URL.
        SP.Web web = context.Web;
        // Load           
        context.Load(web);
        // Execute query. 
        context.ExecuteQuery();
        //
        SP.File ObjFile = web.GetFileByServerRelativeUrl(relurl);
        SP.ListItem item = ObjFile.ListItemAllFields;
        //
        context.Load(ObjFile);
        context.Load(item);
        context.ExecuteQuery();
        //
        //string fileName = item.FieldValues["FileLeafRef"].ToString();
        //string fileType = System.IO.Path.GetExtension(fileName);
        //Guid uniqueId = new Guid(item.FieldValues["UniqueId"].ToString());
        var furl = item.FieldValues["_dlc_DocIdUrl"] as SP.FieldUrlValue;
        permaurl.Clear();
        permaurl.Append(furl.Url);
    }

    private void ConvertPathToURL(ref StringBuilder input)
    {
        input.Replace("J:\\", "https://sp.contoso.com/j/documents/");
        input.Replace("K:\\", "https://sp.contoso.com/k/documents/");
        input.Replace("L:\\", "https://sp.contoso.com/l/documents/");
        input.Replace("\\", "/");
        input.Replace(" ", "%20");
    }
}

