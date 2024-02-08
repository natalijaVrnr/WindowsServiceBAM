using Microsoft.Extensions.FileSystemGlobbing.Abstractions;
using Microsoft.Extensions.FileSystemGlobbing;
using System;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Security.Permissions;
using System.Net;

namespace WindowsService1
{
    public partial class Service1 : ServiceBase
    {
        private string _localFolderPath = "";
        private string _fileSharePath = "";
        private string _logFolderPath = "";
        private string _username = "";
        private string _password = "";

        private readonly FileSystemWatcher _localWatcher = new FileSystemWatcher();
        private readonly FileSystemWatcher _fileShareWatcher = new FileSystemWatcher();

        private static object logFileLock = new object();

        private NetworkCredential credentials;
        //private NetworkCredential credentials = new NetworkCredential($@"DESKTOP-BHSHK2E\n.kundzina@gmail.com", "AmberHeard15");
        public Service1()
        {
            InitializeComponent();
        }

        [PermissionSet(SecurityAction.Demand, Name = "FullTrust")]
        protected override void OnStart(string[] args)
        {
            try 
            {
                WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} Service was started");

                string[] imagePathArgs = Environment.GetCommandLineArgs();

                _localFolderPath = $@"{imagePathArgs[2]}";
                _fileSharePath = $@"{imagePathArgs[4]}";
                _logFolderPath = $@"{imagePathArgs[6]}";
                _username = $@"{imagePathArgs[8]}";
                _password = $@"{imagePathArgs[10]}";

                credentials = new NetworkCredential(_username, _password);

                WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} _localFolderPath: {_localFolderPath}");
                WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} _fileSharePath: {_fileSharePath}");
                WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} _logFolderPath: {_logFolderPath}");

                WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} Setting up local folder watcher");

                _localWatcher.Path = _localFolderPath;
                _localWatcher.IncludeSubdirectories = true;
                _localWatcher.EnableRaisingEvents = true;
                _localWatcher.Created += OnChangedLocal;

                WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} Setting up file share folder watcher");

                using (new NetworkConnection(_fileSharePath, credentials))
                {
                    _fileShareWatcher.Path = _fileSharePath;
                    _fileShareWatcher.IncludeSubdirectories = true;
                    _fileShareWatcher.EnableRaisingEvents = true;
                    _fileShareWatcher.Created += OnChangedFileshare;
                }

                SyncFilesFromLocal();
                SyncFilesFromFileshare();
            }
            catch (Exception ex)
            {
                WriteToLogs(ex.ToString());
            }
        }

        protected override void OnStop()
        {
            try
            {
                WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} Service was stopped");
                _localWatcher.Dispose();
                _fileShareWatcher.Dispose();
            }
            catch (Exception ex)
            {
                WriteToLogs(ex.ToString());
            }
        }

        private void OnChangedLocal(object sender, FileSystemEventArgs e)
        {
            try
            {
                SyncFilesFromLocal();
            }
            catch (Exception ex) 
            { 
                WriteToLogs(ex.ToString());
            }
        }

        private void OnChangedFileshare(object sender, FileSystemEventArgs e)
        {
            try
            {
                SyncFilesFromFileshare();
            }
            catch (Exception ex)
            {
                WriteToLogs(ex.ToString());
            }
        }

        private void SyncFilesFromLocal()
        {
            using (new NetworkConnection(_fileSharePath, credentials))
            {
                WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} Sync from local folder started");

                var matcher = new Matcher();
                matcher.AddInclude("**/*.txt");

                var localDirectoryInfo = new DirectoryInfoWrapper(new DirectoryInfo(_localFolderPath));
                var fileShareDirectoryInfo = new DirectoryInfoWrapper(new DirectoryInfo(_fileSharePath));

                // Sync from local to file share
                var localFiles = matcher.Execute(localDirectoryInfo);
                WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} Found {localFiles.Files.Count()} files in local folder");

                foreach (var localFile in localFiles.Files)
                {
                    var fileSharePath = Path.Combine(fileShareDirectoryInfo.FullName, localFile.Path);
                    var fileShareFileInfo = new FileInfo(fileSharePath);

                    if (!fileShareFileInfo.Exists)
                    {
                        var localPath = Path.Combine(_localFolderPath, localFile.Path);

                        // Copy the file from local to file share
                        WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} Copying {new FileInfo(localFile.Path).Name} to fileshare folder");

                        using (var sourceStream = new FileStream(localPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                        using (var destinationStream = new FileStream(fileSharePath, FileMode.CreateNew))
                        {
                            sourceStream.CopyTo(destinationStream);
                        }

                        //File.Copy(localPath, fileSharePath, true);
                        WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} File {new FileInfo(localFile.Path).Name} successfully copied to fileshare folder");
                    }
                }

                WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} Sync from local folder completed");
            }    
        }

        private void SyncFilesFromFileshare()
        {
            using (new NetworkConnection(_fileSharePath, credentials))
            {
                WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} Sync from fileshare folder started");

                var matcher = new Matcher();
                matcher.AddInclude("**/*.txt");

                var localDirectoryInfo = new DirectoryInfoWrapper(new DirectoryInfo(_localFolderPath));
                var fileShareDirectoryInfo = new DirectoryInfoWrapper(new DirectoryInfo(_fileSharePath));

                // Sync from file share to local
                var fileShareFiles = matcher.Execute(fileShareDirectoryInfo);
                WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} Found {fileShareFiles.Files.Count()} files in fileshare folder");

                foreach (var fileShareFile in fileShareFiles.Files)
                {
                    var localPath = Path.Combine(localDirectoryInfo.FullName, fileShareFile.Path);
                    var localFileInfo = new FileInfo(localPath);

                    if (!localFileInfo.Exists)
                    {
                        var fileSharePath = Path.Combine(_fileSharePath, fileShareFile.Path);

                        // Copy the file from file share to local
                        WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} Copying {new FileInfo(fileShareFile.Path).Name} to local folder");

                        using (var sourceStream = new FileStream(fileSharePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                        using (var destinationStream = new FileStream(localPath, FileMode.CreateNew))
                        {
                            sourceStream.CopyTo(destinationStream);
                        }

                        //File.Copy(fileSharePath, localPath, true);
                        WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} File {new FileInfo(fileShareFile.Path).Name} successfully copied to local folder");
                    }
                }

                WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} Sync from fileshare folder completed");
            }    
        }

        private void WriteToLogs(string msg)
        {
            string path = _logFolderPath + "\\Logs";

            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            string filePath = _logFolderPath + "\\Logs\\ServiceLog" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";

            lock (logFileLock)
            {
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
}
