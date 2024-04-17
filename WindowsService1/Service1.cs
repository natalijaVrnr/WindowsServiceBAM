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
        private string _fileShareUsername = "";
        private string _fileSharePassword = "";
        private string _localFolderUsername = "";
        private string _localFolderPassword = "";

        private readonly FileSystemWatcher _localWatcher = new FileSystemWatcher();
        private readonly FileSystemWatcher _fileShareWatcher = new FileSystemWatcher();

        private static object _logFileLock = new object();

        private NetworkCredential _fileShareCredentials;
        private NetworkCredential _localFolderCredentials;
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

                _localFolderUsername = $@"{imagePathArgs[8]}";
                _localFolderPassword = $@"{imagePathArgs[10]}";

                _fileShareUsername = $@"{imagePathArgs[12]}";
                _fileSharePassword = $@"{imagePathArgs[14]}";


                _localFolderCredentials = new NetworkCredential(_localFolderUsername, _localFolderPassword);
                _fileShareCredentials = new NetworkCredential(_fileShareUsername, _fileSharePassword);

                WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} _localFolderPath: {_localFolderPath}");
                WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} _fileSharePath: {_fileSharePath}");
                WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} _logFolderPath: {_logFolderPath}");
                WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} _localFolderUsername: {_localFolderUsername}");
                WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} _localFolderPassword: {_localFolderPassword}");
                WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} _fileShareUsername: {_fileShareUsername}");
                WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} _fileSharePassword: {_fileSharePassword}");

                WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} Setting up local folder watcher");

                using (new NetworkConnection(_localFolderPath, _localFolderCredentials))
                {
                    _localWatcher.Path = _localFolderPath;
                    _localWatcher.IncludeSubdirectories = true;
                    _localWatcher.EnableRaisingEvents = true;
                    _localWatcher.Created += OnChangedLocal;
                }

                WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} Setting up file share folder watcher");

                using (new NetworkConnection(_fileSharePath, _fileShareCredentials))
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
                WriteToLogs(ex.ToString());
            }
        }

        private void SyncFilesFromLocal()
        {
            using (new NetworkConnection(_localFolderPath, _localFolderCredentials))
            {
                using (new NetworkConnection(_fileSharePath, _fileShareCredentials))
                {
                    WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} Sync from local folder triggered");

                    var matcher = new Matcher();
                    matcher.AddInclude("**/vfpCmds.tmp");

                    var localDirectoryInfo = new DirectoryInfoWrapper(new DirectoryInfo(_localFolderPath));
                    var fileShareDirectoryInfo = new DirectoryInfoWrapper(new DirectoryInfo(_fileSharePath));

                    // Sync from local to file share
                    var localFiles = matcher.Execute(localDirectoryInfo);

                    foreach (var localFile in localFiles.Files)
                    {
                        var fileSharePath = Path.Combine(fileShareDirectoryInfo.FullName, localFile.Path);
                        var fileShareFileInfo = new FileInfo(fileSharePath);

                        WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} Found vfpCmds.tmp in local folder");

                        if (!fileShareFileInfo.Exists)
                        {
                            var localPath = Path.Combine(_localFolderPath, localFile.Path);

                            bool success = false;
                            int retries = 3;
                            int delayMilliseconds = 1000;

                            while (!success && retries > 0)
                            {
                                try
                                {
                                    // Copy the file from local to fileshare
                                    WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} Copying {new FileInfo(localFile.Path).Name} to fileshare folder");

                                    using (var sourceStream = new FileStream(localPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                                    using (var destinationStream = new FileStream(fileSharePath, FileMode.CreateNew))
                                    {
                                        sourceStream.CopyTo(destinationStream);
                                    }

                                    //File.Copy(localPath, fileSharePath, true);
                                    WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} File {new FileInfo(localFile.Path).Name} successfully copied to fileshare folder");

                                    //Delete from fileshare
                                    File.Delete(localPath);

                                    WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} File {new FileInfo(localFile.Path).Name} successfully deleted from local folder");

                                    success = true;
                                }
                                catch (Exception ex)
                                {
                                    // Handle file access errors
                                    WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} Error copying file {ex.Message} to fileshare folder. Retry {4 - retries}");

                                    // Wait for a while before retrying
                                    System.Threading.Thread.Sleep(delayMilliseconds);

                                    retries--;
                                }
                            }

                            if (!success)
                            {
                                WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} Failed to copy file {new FileInfo(localFile.Path).Name} to fileshare folder");
                            }

                        }
                    }

                    WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} Sync from local folder completed");
                }
            } 
        }

        private void SyncFilesFromFileshare()
        {
            using (new NetworkConnection(_fileSharePath, _fileShareCredentials))
            {
                using (new NetworkConnection(_localFolderPath, _localFolderCredentials))
                {
                    WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} Sync from fileshare folder triggered");

                    var matcher = new Matcher();
                    matcher.AddInclude("**/vfpResult.tmp");

                    var localDirectoryInfo = new DirectoryInfoWrapper(new DirectoryInfo(_localFolderPath));
                    var fileShareDirectoryInfo = new DirectoryInfoWrapper(new DirectoryInfo(_fileSharePath));

                    // Sync from file share to local
                    var fileShareFiles = matcher.Execute(fileShareDirectoryInfo);

                    foreach (var fileShareFile in fileShareFiles.Files)
                    {
                        var localPath = Path.Combine(localDirectoryInfo.FullName, fileShareFile.Path);
                        var localFileInfo = new FileInfo(localPath);

                        WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} Found vfpResult.tmp in fileshare folder");

                        if (!localFileInfo.Exists)
                        {
                            var fileSharePath = Path.Combine(_fileSharePath, fileShareFile.Path);

                            bool success = false;
                            int retries = 3;
                            int delayMilliseconds = 1000;

                            while (!success && retries > 0)
                            {
                                try
                                {
                                    // Copy the file from file share to local
                                    WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} Copying {new FileInfo(fileShareFile.Path).Name} to local folder");

                                    using (var sourceStream = new FileStream(fileSharePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                                    using (var destinationStream = new FileStream(localPath, FileMode.CreateNew))
                                    {
                                        sourceStream.CopyTo(destinationStream);
                                    }

                                    //File.Copy(fileSharePath, localPath, true);
                                    WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} File {new FileInfo(fileShareFile.Path).Name} successfully copied to local folder");

                                    //Delete from fileshare
                                    File.Delete(fileSharePath);

                                    success = true;
                                }

                                catch (Exception ex)
                                {
                                    // Handle file access errors
                                    WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} Error copying file {ex.Message} to local folder. Retry {4 - retries}");

                                    // Wait for a while before retrying
                                    System.Threading.Thread.Sleep(delayMilliseconds);

                                    retries--;
                                }
                            }

                            if (!success)
                            {
                                WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} Failed to copy file {new FileInfo(fileShareFile.Path).Name} to local folder");
                            }
                        }
                    }

                    WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} Sync from fileshare folder completed");
                }
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

            lock (_logFileLock)
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
