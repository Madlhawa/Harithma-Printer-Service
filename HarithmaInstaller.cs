using System.Collections;
using System.ComponentModel;
using System.Configuration.Install;
using System.IO;

[RunInstaller(true)]
public class HarithmaInstaller : Installer
{
    public HarithmaInstaller()
    {
        this.Committed += new InstallEventHandler(OnCommitted);
    }

    private void OnCommitted(object sender, InstallEventArgs e)
    {
        // Delay for 3 seconds (3000 milliseconds)
        System.Threading.Thread.Sleep(3000);

        // Path to the application to be launched
        string targetDir = Context.Parameters["targetdir"];
        string appPath = Path.Combine(targetDir, "Printer Service.exe");

        if (File.Exists(appPath))
        {
            // Launch the application
            System.Diagnostics.Process.Start(appPath);
        }
    }

    public override void Install(IDictionary savedState)
    {
        base.Install(savedState);
    }

    public override void Commit(IDictionary savedState)
    {
        base.Commit(savedState);
    }

    public override void Rollback(IDictionary savedState)
    {
        base.Rollback(savedState);
    }

    public override void Uninstall(IDictionary savedState)
    {
        base.Uninstall(savedState);
    }
}
