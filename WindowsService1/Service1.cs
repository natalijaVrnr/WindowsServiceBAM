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
        public Service1()
        {
            InitializeComponent();
        }

        // executed when service is being started
        [PermissionSet(SecurityAction.Demand, Name = "FullTrust")]
        protected override void OnStart(string[] args)
        {
            try 
            {
                WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} Service was started");

                string[] imagePathArgs = Environment.GetCommandLineArgs();

                // gettings all variable values from imagepath variable (the one edited in regedit)
                _localFolderPath = $@"{imagePathArgs[2]}";
                _fileSharePath = $@"{imagePathArgs[4]}";
                _logFolderPath = $@"{imagePathArgs[6]}";

                _localFolderUsername = $@"{imagePathArgs[8]}";
                _localFolderPassword = $@"{imagePathArgs[10]}";

                _fileShareUsername = $@"{imagePathArgs[12]}";
                _fileSharePassword = $@"{imagePathArgs[14]}";

                // establishing network connection credentials
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

                // setting file watchers for local folder and for fileshare
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

                // initial sync - there maybe errors from fromlocal sync if fileshare manages to transfer some files and localsync will be called already before local folder sync method is called, but that is okay
                // we need to call both methods to make sure the folders are synced

                SyncFilesFromFileshare();
                SyncFilesFromLocal();
            }
            catch (Exception ex)
            {
                WriteToLogs(ex.ToString());
            }
        }

        // executed when service is being stopped
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

        // sync from local to file share
        private void SyncFilesFromLocal()
        {
            using (new NetworkConnection(_localFolderPath, _localFolderCredentials))
            {
                using (new NetworkConnection(_fileSharePath, _fileShareCredentials))
                {
                    WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} Sync from local folder triggered");

                    // setting matcher for detecting file
                    var matcher = new Matcher();
                    matcher.AddInclude("**/vfpCmds.tmp");

                    var localDirectoryInfo = new DirectoryInfoWrapper(new DirectoryInfo(_localFolderPath));
                    var fileShareDirectoryInfo = new DirectoryInfoWrapper(new DirectoryInfo(_fileSharePath));

                    // fetching all files that match the matcher
                    var localFiles = matcher.Execute(localDirectoryInfo);

                    foreach (var localFile in localFiles.Files)
                    {
                        // constructing the same path in fileshare folder to check if file already exists there
                        var fileSharePath = Path.Combine(fileShareDirectoryInfo.FullName, localFile.Path);
                        var fileShareFileInfo = new FileInfo(fileSharePath);

                        WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} Found vfpCmds.tmp in local folder");

                        if (!fileShareFileInfo.Exists)
                        {
                            var localPath = Path.Combine(_localFolderPath, localFile.Path);

                            bool success = false;
                            int retries = 3;
                            int delayMilliseconds = 1000;

                            // 3 retries with 1 second delay between each, after that consider as failed
                            while (!success && retries > 0)
                            {
                                try
                                {
                                    // copy the file from local to fileshare
                                    WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} Copying {new FileInfo(localFile.Path).Name} to fileshare folder");

                                    using (var sourceStream = new FileStream(localPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                                    using (var destinationStream = new FileStream(fileSharePath, FileMode.CreateNew))
                                    {
                                        sourceStream.CopyTo(destinationStream);
                                    }

                                    WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} File {new FileInfo(localFile.Path).Name} successfully copied to fileshare folder");

                                    // delete the file from fileshare
                                    File.Delete(localPath);

                                    WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} File {new FileInfo(localFile.Path).Name} successfully deleted from local folder");

                                    success = true;
                                }
                                // handle file access errors
                                catch (Exception ex)
                                {
                                    WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} Error copying file {ex.Message} to fileshare folder. Retry {4 - retries}");

                                    // wait for a while before retrying
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

                    var matcherCmds = new Matcher();
                    matcherCmds.AddInclude("**/vfpCmds.tmp");

                    var localDirectoryInfo = new DirectoryInfoWrapper(new DirectoryInfo(_localFolderPath));
                    var fileShareDirectoryInfo = new DirectoryInfoWrapper(new DirectoryInfo(_fileSharePath));

                    // Sync from file share to local
                    var fileShareFiles = matcher.Execute(fileShareDirectoryInfo);
                    var cmdFiles = matcherCmds.Execute(fileShareDirectoryInfo);

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

                    // delete cmd file if there is any and after result file has been copied to local folder
                    if (fileShareFiles.Files.Count() > 0)
                    {
                        foreach (var cmdFile in cmdFiles.Files)
                        {
                            WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} Deleting file {new FileInfo(cmdFile.Path).Name} from fileshare folder");
                            var fileSharePath = Path.Combine(_fileSharePath, cmdFile.Path);
                            File.Delete(fileSharePath);
                        }

                        WriteToLogs($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss.fff")} Sync from fileshare folder completed");
                    }
                    
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
