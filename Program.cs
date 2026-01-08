using System;
using System.IO;
using System.Reflection;
using System.Management.Automation;
using System.Management.Automation.Runspaces;

namespace BMAT_CC_Host
{
    internal class Program
    {
        static int Main(string[] args)
        {
            try
            {
                string psHome = AppContext.BaseDirectory;
                Environment.SetEnvironmentVariable("PSHOME", psHome);

                // Load embedded PowerShell script
                var assembly = Assembly.GetExecutingAssembly();

                // IMPORTANT:
                // Embedded resource name is usually: <RootNamespace>.<FileName>
                const string resourceName = "BMAT_CC_Host.BMAT_CC.ps1";

                using Stream? stream = assembly.GetManifestResourceStream(resourceName);
                if (stream == null)
                    throw new InvalidOperationException($"Embedded script not found: {resourceName}");

                using StreamReader reader = new StreamReader(stream);
                string scriptContents = reader.ReadToEnd();

                // Create PowerShell runspace WITHOUT snap-ins (required for single-file apps)
                InitialSessionState iss = InitialSessionState.CreateDefault2();

                using Runspace runspace = RunspaceFactory.CreateRunspace(iss);
                runspace.Open();

                using PowerShell ps = PowerShell.Create();
                ps.Runspace = runspace;

                // Pass command-line args to the script (optional)
                ps.AddScript(scriptContents);
                ps.AddParameter("Args", args);

                ps.Invoke();

                // Report PowerShell errors cleanly
                if (ps.HadErrors)
                {
                    Console.Error.WriteLine("PowerShell execution failed:");
                    foreach (var error in ps.Streams.Error)
                        Console.Error.WriteLine(error.ToString());

                    return 1;
                }

                return 0;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("FATAL ERROR:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
                Console.WriteLine();
                Console.WriteLine("Press any key to exit...");
                Console.ReadKey();
                return -1;
            }
        }
    }
}
