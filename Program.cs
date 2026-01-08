using System;
using System.IO;
using System.Reflection;
using System.Management.Automation;

namespace BMAT_CC_Host
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // This fix tells PowerShell that "Home" is wherever this EXE is running
                string appPath = AppContext.BaseDirectory;
                Environment.SetEnvironmentVariable("PSHOME", appPath);

                var assembly = Assembly.GetExecutingAssembly();
                // Replace 'BMAT_CC' with your actual Namespace/Project name if needed
                var resourceName = "BMAT_CC.BMAT_CC.ps1";

                using (Stream stream = assembly.GetManifestResourceStream(resourceName))
                {
                    if (stream == null) throw new Exception($"Could not find embedded script: {resourceName}");

                    using (StreamReader reader = new StreamReader(stream))
                    {
                        string scriptContents = reader.ReadToEnd();
                        using (PowerShell ps = PowerShell.Create())
                        {
                            ps.AddScript(scriptContents);
                            ps.Invoke();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("FATAL ERROR: " + ex.Message);
                Console.WriteLine(ex.StackTrace);
                Console.WriteLine("Press any key to exit...");
                Console.ReadKey();
            }
        }
    }
}