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
                //string baseDir = AppContext.BaseDirectory;
                string baseDir = AppContext.BaseDirectory.TrimEnd(Path.DirectorySeparatorChar);
                string psHome = Path.Combine(baseDir, "Bin");

                if (!Directory.Exists(psHome))
                    throw new DirectoryNotFoundException($"PowerShell home not found: {psHome}");

                Environment.SetEnvironmentVariable("PSHOME", psHome);
                Environment.SetEnvironmentVariable(
                    "PSModulePath",
                    Path.Combine(psHome, "Modules"),
                    EnvironmentVariableTarget.Process
                );

                Console.WriteLine("BaseDirectory: " + baseDir);
                Console.WriteLine("PSHOME: " + psHome);

                // Load embedded script
                Assembly asm = Assembly.GetExecutingAssembly();
                const string scriptResource = "BMAT_CC_Host.BMAT_CC.ps1";

                using Stream scriptStream =
                    asm.GetManifestResourceStream(scriptResource)
                    ?? throw new InvalidOperationException($"Embedded script not found: {scriptResource}");

                using StreamReader reader = new StreamReader(scriptStream);
                string script = reader.ReadToEnd();

                InitialSessionState iss = InitialSessionState.CreateDefault2();
                using Runspace runspace = RunspaceFactory.CreateRunspace(iss);
                runspace.Open();

                // Create PowerShell instance and assign runspace
                using PowerShell ps = PowerShell.Create();
                ps.Runspace = runspace;
                ps.AddScript(script, useLocalScope: true);
                ps.AddParameter("Args", args);
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
