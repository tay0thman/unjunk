using System;
using System.IO;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Threading;
using System.Windows.Forms;

internal static class UnjunkLauncher
{
    [STAThread]
    private static void Main()
    {
        try
        {
            string script = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "unjunk.ps1");
            if (!File.Exists(script)) throw new FileNotFoundException("Keep Unjunk.exe beside unjunk.ps1 and the Unjunk modules.", script);
            InitialSessionState session = InitialSessionState.CreateDefault();
            // This applies only to this embedded session, never machine/user policy.
            session.ExecutionPolicy = Microsoft.PowerShell.ExecutionPolicy.Bypass;
            using (Runspace runspace = RunspaceFactory.CreateRunspace(session))
            {
                runspace.ApartmentState = ApartmentState.STA;
                runspace.ThreadOptions = PSThreadOptions.UseCurrentThread;
                runspace.Open();
                using (PowerShell shell = PowerShell.Create())
                {
                    shell.Runspace = runspace;
                    shell.AddCommand(script).AddParameter("Gui");
                    shell.Invoke();
                    if (shell.HadErrors)
                    {
                        string details = String.Join(Environment.NewLine, shell.Streams.Error);
                        MessageBox.Show(details, "Unjunk", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }
        catch (Exception error)
        {
            MessageBox.Show(error.Message, "Unjunk could not start", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
