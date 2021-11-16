using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ImportExcel2Access
{
    [RunInstaller(true)]
    public partial class Installer : System.Configuration.Install.Installer
    {
        public Installer()
        {
            InitializeComponent();
        }

        public override void Uninstall(IDictionary savedState)
        {
            DeleteSetting();

            Process application = null;
            try
            {
                foreach (var process in Process.GetProcesses())
                {
                    if (!process.ProcessName.ToLower().Contains("creatinginstaller"))
                    {
                        continue;
                    }

                    application = process;

                    break;
                }

                if (application != null && application.Responding)
                {
                    application.Kill();
                    base.Uninstall(savedState);
                }
            }
            catch
            {

            }

            base.Uninstall(savedState);
        }

        private void DeleteSetting()
        {
            string pathtodelete = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string ignorFile = Assembly.GetExecutingAssembly().GetName().ToString();
            List<string> listFile = Directory.GetFiles(pathtodelete).ToList();
            string[] subDirectory = Directory.GetDirectories(pathtodelete);

            // Delete sub folder
            foreach (var folder in subDirectory)
            {
                try
                {
                    Directory.Delete(folder, true);
                }
                catch
                {

                }
            }

            try
            {
                foreach (string file in listFile)
                {
                    string fileName = Path.GetFileName(file);
                    if (ignorFile.Contains(fileName))
                        continue;
                    File.Delete(file);
                }
            }
            catch
            {

            }     
        }
    }
}
