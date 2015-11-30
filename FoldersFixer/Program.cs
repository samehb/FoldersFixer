using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Diagnostics;
using System.Text.RegularExpressions;
using Microsoft.Win32;
using System.IO;

namespace FoldersFixer
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            if (args.Length == 0) // If there are no arguments provided simply register.
                Register("folder", "FolderFixer", "Fix Folders Names", "\"" + Application.ExecutablePath + "\" \"%L\"");
            else
            {
                if (args[0] == "-unreg")
                    Unregister("folder", "FolderFixer");
                else
                    FixFoldersName(args[0]);
            }
        }

        // Shell Extension - Register
        static void Register(string mainkey, string subkey, string commandtext, string commandpath)
        {
            string RPath = mainkey + @"\shell\" + subkey;
            string CRPath = RPath + @"\command";
            using (RegistryKey key = Registry.ClassesRoot.CreateSubKey(RPath))
            {
                key.SetValue(null, commandtext);
                using (RegistryKey key2 = Registry.ClassesRoot.CreateSubKey(CRPath))
                {
                    key2.SetValue(null, commandpath);
                }
            }
        }

        // Shell Extension - Unregister
        static void Unregister(string mainkey, string subkey)
        {
            string RPath = mainkey + @"\shell\" + subkey;
            Registry.ClassesRoot.DeleteSubKeyTree(RPath);
        }

        // This method is getting a listing subdirectories within a directory providing both short and long names.
        static string ProcessDirs(string basedir)
        {
            string output = "";
            try
            {
                Process Process = new Process();
                ProcessStartInfo StartInfo = new ProcessStartInfo();
                StartInfo.FileName = "cmd.exe";
                StartInfo.Arguments = "/c dir /ad /x";
                StartInfo.UseShellExecute = false;
                StartInfo.RedirectStandardOutput = true;
                StartInfo.CreateNoWindow = true;
                StartInfo.WorkingDirectory = basedir;
                StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                Process.StartInfo = StartInfo;
                Process.Start();
                output = Process.StandardOutput.ReadToEnd();
                Process.WaitForExit();
                output = output.Replace("\r\n", "\n");
            }
            catch
            { }

            return output;
        }

        // This method is for renaming the folders having trailing dots in their names.
        static void FixFoldersName(string basedir)
        {
            string BaseDir = basedir;
            string[] Directories = ProcessDirs(BaseDir).Split('\n');
            if (Directories.Length < 10)
                return;
            foreach (string line in Directories)
            {
                Regex r = new Regex(@"<DIR>\s{10}(\S+)\s+(.*)\.$");
                Match match = r.Match(line);
                if (match.Success)
                {
                    string ShortPath = BaseDir + "\\" + match.Groups[1].Value;
                    string LongPath = BaseDir + "\\" + match.Groups[2].Value;
                    Directory.Move(ShortPath, LongPath + "_renamed");// Two statements are required for renaming because of the limitation of Directory.Move (not detecting the trailing dot in folder name).
                    Directory.Move(LongPath + "_renamed", LongPath);//
                }
            }
        }
    }
}
