using Microsoft.Extensions.FileSystemGlobbing.Abstractions;
using Microsoft.Extensions.FileSystemGlobbing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace WindowsService1
{
    public partial class Service1 : ServiceBase
    {
        private readonly string _localFolderPath = @"C:\Users\natal\OneDrive\Desktop\localfolder";
        private readonly string _fileSharePath = @"\\DESKTOP-BHSHK2E\fileshare";
        private readonly string _logFolderPath = @"C:\Users\natal\OneDrive\Desktop\logfolder";

        private readonly FileSystemWatcher _localWatcher = new FileSystemWatcher();
        private readonly FileSystemWatcher _fileShareWatcher = new FileSystemWatcher();
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} Service was started");
            WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} Setting up local folder watcher");
            SetupWatcher(_localWatcher, _localFolderPath);
            WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} Setting up file share folder watcher");
            SetupWatcher(_fileShareWatcher, _fileSharePath);
        }

        protected override void OnStop()
        {
            WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} Service was stopped");
            _localWatcher.Dispose();
            _fileShareWatcher.Dispose();
        }

        private void SetupWatcher(FileSystemWatcher watcher, string path)
        {
            watcher.Path = path;
            watcher.IncludeSubdirectories = true;
            watcher.EnableRaisingEvents = true;

            watcher.Created += OnChanged;
        }

        private void OnChanged(object sender, FileSystemEventArgs e)
        {
            WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} Syncing files started");
            SyncFiles();
        }

        private void SyncFiles()
        {
            var matcher = new Matcher();
            matcher.AddInclude("**/*.txt");

            var localDirectoryInfo = new DirectoryInfoWrapper(new DirectoryInfo(_localFolderPath));
            var fileShareDirectoryInfo = new DirectoryInfoWrapper(new DirectoryInfo(_fileSharePath));

            // Sync from local to file share
            var localFiles = matcher.Execute(localDirectoryInfo);
            WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} Found {localFiles.Files.Count()} new files in local folder");

            foreach (var localFile in localFiles.Files)
            {
                var fileSharePath = Path.Combine(fileShareDirectoryInfo.FullName, localFile.Path);
                var fileShareFileInfo = new FileInfo(fileSharePath);

                if (!fileShareFileInfo.Exists)
                {
                    // Copy the file from local to file share
                    WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} Copying {new FileInfo(localFile.Path).Name} to fileshare");
                    File.Copy(localFile.Path, fileSharePath, true);
                    WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} File {new FileInfo(localFile.Path).Name} successfully copied to fileshare");
                }
            }

            // Sync from file share to local
            var fileShareFiles = matcher.Execute(fileShareDirectoryInfo);
            WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} Found {fileShareFiles.Files.Count()} new files in file share folder");

            foreach (var fileShareFile in fileShareFiles.Files)
            {
                var localPath = Path.Combine(localDirectoryInfo.FullName, fileShareFile.Path);
                var localFileInfo = new FileInfo(localPath);

                if (!localFileInfo.Exists)
                {
                    // Copy the file from file share to local
                    WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} Copying {new FileInfo(fileShareFile.Path).Name} to local folder");
                    File.Copy(fileShareFile.Path, localPath, true);
                    WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} File {new FileInfo(fileShareFile.Path).Name} successfully copied to local folder");
                }
            }

            WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} Sync completed");
        }

        private void WriteToLogs(string msg)
        {
            string path = _logFolderPath + "\\Logs";

            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            string filePath = _logFolderPath + "\\Logs\\ServiceLog" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";

            if (!File.Exists(filePath))
            {
                using (StreamWriter sw = File.CreateText(filePath))
                {
                    sw.WriteLine(msg);
                }
            }

            else
            {
                using (StreamWriter sw = File.AppendText(filePath))
                {
                    sw.WriteLine(msg);
                }
            }

        }
    }
}
