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
                string baseDir = AppContext.BaseDirectory;
                string psHome = Path.Combine(baseDir, "Bin");

                if (!Directory.Exists(psHome))
                    throw new DirectoryNotFoundException($"PSHOME not found: {psHome}");

                Environment.SetEnvironmentVariable("PSHOME", psHome);
                Environment.SetEnvironmentVariable("PSModulePath", Path.Combine(psHome, "Modules"));

                Console.WriteLine("BaseDirectory: " + baseDir);
                Console.WriteLine("PSHOME: " + psHome);

                Assembly asm = Assembly.GetExecutingAssembly();
                using Stream scriptStream =
                    asm.GetManifestResourceStream("BMAT_CC_Host.BMAT_CC.ps1")
                    ?? throw new InvalidOperationException("Embedded script not found");

                using StreamReader reader = new StreamReader(scriptStream);
                string script = reader.ReadToEnd();

                InitialSessionState iss = InitialSessionState.Create();

                // Explicitly import built-in modules
                iss.ImportPSModule(new[]
                {
                    "Microsoft.PowerShell.Management",
                    "Microsoft.PowerShell.Utility",
                    "Microsoft.PowerShell.Security",
                    "Microsoft.PowerShell.Archive"
                });

                using Runspace runspace = RunspaceFactory.CreateRunspace(iss);
                runspace.Open();

                // Create PowerShell instance and assign runspace
                using PowerShell ps = PowerShell.Create();
                ps.Runspace = runspace;

                // Add script and pass command-line arguments
                ps.AddScript(script);
                ps.AddParameter("Args", args);

                // Invoke the script
                ps.Invoke();

                // Check for errors
                if (ps.HadErrors)
                {
                    Console.Error.WriteLine("PowerShell execution failed:");
                    foreach (var error in ps.Streams.Error)
                        Console.Error.WriteLine(error.ToString());
                    Console.WriteLine();
                    Console.WriteLine("Press any key to exit...");
                    Console.ReadKey();
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
