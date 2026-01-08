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
                // 1️⃣ Determine exe folder
                string exeDir = AppContext.BaseDirectory;
                string binPath = Path.Combine(exeDir, "Bin");

                // 2️⃣ Create Bin folder if it doesn't exist
                Directory.CreateDirectory(binPath);

                // 3️⃣ Extract embedded PowerShell runtime files (DLLs, Modules)
                ExtractEmbeddedResources(binPath);

                // 4️⃣ Set PSHOME to Bin folder
                Environment.SetEnvironmentVariable("PSHOME", binPath);

                Console.WriteLine("BaseDirectory: " + exeDir);
                Console.WriteLine("PSHOME: " + Environment.GetEnvironmentVariable("PSHOME"));

                // 5️⃣ Load embedded PowerShell script
                string scriptContents = LoadEmbeddedScript("BMAT_CC_Host.BMAT_CC.ps1");

                // 6️⃣ Create a runspace with full language support
                InitialSessionState iss = InitialSessionState.CreateDefault2();
                using Runspace runspace = RunspaceFactory.CreateRunspace(iss);
                runspace.Open();

                // 7️⃣ Create PowerShell instance and assign runspace
                using PowerShell ps = PowerShell.Create();
                ps.Runspace = runspace;

                // 8️⃣ Add script and pass command-line arguments
                ps.AddScript(scriptContents);
                ps.AddParameter("Args", args);

                // 9️⃣ Invoke the script
                ps.Invoke();

                // 10️⃣ Check for errors
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

        /// <summary>
        /// Loads the embedded .ps1 script from the assembly resources.
        /// </summary>
        private static string LoadEmbeddedScript(string resourceName)
        {
            var assembly = Assembly.GetExecutingAssembly();
            using Stream? stream = assembly.GetManifestResourceStream(resourceName);
            if (stream == null)
                throw new InvalidOperationException($"Embedded script not found: {resourceName}");

            using StreamReader reader = new StreamReader(stream);
            return reader.ReadToEnd();
        }

        /// <summary>
        /// Extracts embedded PowerShell runtime files and modules to the target directory.
        /// Resource names must start with "BMAT_CC_Host.PowerShell7" (adjust as needed).
        /// </summary>
        private static void ExtractEmbeddedResources(string targetDir)
        {
            var assembly = Assembly.GetExecutingAssembly();

            foreach (string resourceName in assembly.GetManifestResourceNames())
            {
                // Only extract PowerShell runtime resources
                if (!resourceName.StartsWith("BMAT_CC_Host.PowerShell7"))
                    continue;

                using Stream? stream = assembly.GetManifestResourceStream(resourceName);
                if (stream == null) continue;

                // Remove namespace prefix and convert to relative path
                string relativePath = resourceName.Substring("BMAT_CC_Host.PowerShell7.".Length).Replace('.', Path.DirectorySeparatorChar);

                string fullPath = Path.Combine(targetDir, relativePath);

                // Ensure directory exists
                Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);

                // Write the file
                using (var fs = new FileStream(fullPath, FileMode.Create, FileAccess.Write))
                {
                    stream.CopyTo(fs);
                }
            }
        }
    }
}
