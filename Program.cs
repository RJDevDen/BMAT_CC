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
            var assembly = Assembly.GetExecutingAssembly();
            // Note: Replace 'BMAT_CC' with your actual Project Name if different
            var resourceName = "BMAT_CC.BMAT_CC.ps1";

            using (Stream stream = assembly.GetManifestResourceStream(resourceName))
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
}