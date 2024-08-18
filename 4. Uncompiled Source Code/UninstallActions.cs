using System;
using System.ComponentModel;
using System.Configuration.Install;
using System.IO;

[RunInstaller(true)]
public class CustomActions : Installer
{
    public override void Uninstall(System.Collections.IDictionary savedState)
    {
        base.Uninstall(savedState);

        string genAppDataFolder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        string appData = Path.Combine(genAppDataFolder, "GES");

        if (Directory.Exists(appData))
        {
            Directory.Delete(appData, true); // true means it will delete subdirectories and files in the directory as well
        }
    }
}
