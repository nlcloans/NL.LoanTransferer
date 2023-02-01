using WinSCP;
using SshNet;
using System.Diagnostics;
using System.Text;
//using Renci.SshNet;

namespace System.IO
{
    public static class ExtendedMethod
    {
        //private readonly Settings _settings;


        public static void Main(String[] args)
        {
            int count = 0;
            int intervalCount = 0;
            StringBuilder sb = new StringBuilder();

            try
            {
                SessionOptions sessionOptions = new SessionOptions
                {
                    Protocol = Protocol.Sftp,
                    HostName = "",
                    UserName = "",
                    Password = "",
                    Timeout = TimeSpan.FromMinutes(5),
                    GiveUpSecurityAndAcceptAnySshHostKey = true
                };

                string remoteDirectory = "/";
                DirectoryInfo closing = new DirectoryInfo(@"C:\Closing");
                DirectoryInfo credit = new DirectoryInfo(@"C:\Credit");

                using (Session session = new Session())
                {
                    session.Open(sessionOptions);
                    Console.WriteLine("Connected to SFTP");
                    sb.Append("Connected to SFTP");

                    TransferOptions transferOptions = new TransferOptions();
                    transferOptions.TransferMode = TransferMode.Binary;
                    TransferOperationResult transferResult;
                    bool next = true;

                    while(next)
                    {
                        FileInfo[] closingInfo = closing.GetFiles();
                        if (closingInfo.Length > 0)
                        {
                            Console.WriteLine("Watching Closing");
                            sb.AppendLine("Watching Closing");

                            foreach (FileInfo f in closingInfo)
                            {
                                var loanNumber = Path.GetFileName(f.FullName);
                                var fileAccess = FileLocked(f);

                                //Closing
                                if (!f.Name.Contains("_1") && f.Extension == ".pdf" && f.Length > 5000 && fileAccess == true)
                                {
                                    Console.WriteLine(loanNumber + " File Name");
                                    sb.AppendLine(loanNumber + " File Name");
                                    File.Move(f.FullName, f.FullName.Replace(".pdf", "_1.pdf"));
                                    Console.WriteLine(loanNumber + " File Renamed");
                                    sb.AppendLine(loanNumber + " File Renamed");
                                    transferResult = session.PutFiles(f.FullName.Replace(".pdf", "_1.pdf"), remoteDirectory, false, transferOptions);
                                    Console.WriteLine(loanNumber + " Checking for Upload");
                                    sb.AppendLine(loanNumber + " Checking for Upload");
                                    transferResult.Check();
                                    Console.WriteLine(loanNumber + " Loan Uploaded");
                                    sb.AppendLine(loanNumber + " Loan Uploaded");
                                    File.Delete(f.FullName.Replace(".pdf", "_1.pdf"));
                                    Console.WriteLine(loanNumber + " Deleted");
                                    sb.AppendLine(loanNumber + " Deleted");
                                    Console.WriteLine("=======================");
                                    sb.AppendLine("=======================");
                                    count++;
                                    Console.WriteLine(count + " Loans Processed");
                                    sb.AppendLine(count + " Loans Processed");
                                    intervalCount++;
                                    File.AppendAllText(@"C:\LoanLog\log" + DateTime.Now.Day + ".txt", sb.ToString());
                                    sb.Clear();
                                }

                                if (intervalCount == 30)
                                {
                                    GarbageCollection(sb);
                                    intervalCount = 0;
                                }
                            }
                        }

                        FileInfo[] creditInfo = credit.GetFiles();
                        if (creditInfo.Length > 0)
                        {
                            Console.WriteLine("Watching Credit");
                            sb.AppendLine("Watching Credit");

                            foreach (FileInfo f in creditInfo)
                            {
                                var loanNumber = Path.GetFileName(f.FullName);
                                var fileAccess = FileLocked(f);

                                //Credit
                                if (!f.Name.Contains("_2") && f.Extension == ".pdf" && f.Length > 5000 && fileAccess == true)
                                {
                                    Console.WriteLine(loanNumber + " File Name");
                                    sb.AppendLine(loanNumber + " File Name");
                                    File.Move(f.FullName, f.FullName.Replace(".pdf", "_2.pdf"));
                                    Console.WriteLine(loanNumber + " File Renamed");
                                    sb.AppendLine(loanNumber + " File Renamed");
                                    transferResult = session.PutFiles(f.FullName.Replace(".pdf", "_2.pdf"), remoteDirectory, false, transferOptions);
                                    Console.WriteLine(loanNumber + " Checking for Upload");
                                    sb.AppendLine(loanNumber + " Checking for Upload");
                                    transferResult.Check();
                                    Console.WriteLine(loanNumber + " Loan Uploaded");
                                    sb.AppendLine(loanNumber + " Loan Uploaded");
                                    File.Delete(f.FullName.Replace(".pdf", "_2.pdf"));
                                    Console.WriteLine(loanNumber + " Deleted");
                                    sb.AppendLine(loanNumber + " Deleted");
                                    Console.WriteLine("=======================");
                                    sb.AppendLine("=======================");
                                    count++;
                                    Console.WriteLine(count + " Loans Processed");
                                    sb.AppendLine(count + " Loans Processed");
                                    intervalCount++;
                                    File.AppendAllText(@"C:\LoanLog\log" + DateTime.Now.Day + ".txt", sb.ToString());
                                    sb.Clear();
                                }

                                if(intervalCount == 10)
                                {
                                    GarbageCollection(sb);
                                    intervalCount = 0;
                                }
                            }
                        }

                        ProcessCheck(sb);

                        Console.WriteLine("Waiting for files from Encompass");
                        sb.AppendLine("Waiting for files from Encompass");
                        Thread.Sleep(15000);
                    }

                    Console.WriteLine("Total uploaded: " + count);
                    sb.AppendLine("Total uploaded: " + count);
                    File.AppendAllText(@"C:\LoanLog\log" + DateTime.Now + ".txt", sb.ToString());
                    session.Close();
                    sb.Clear();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Total uploaded before exception: " + count);
                Console.WriteLine("Error: {0}", e);

                sb.AppendLine(e.ToString());

                File.AppendAllText(@"C:\LoanLog\logCrash" + DateTime.Now + ".txt", sb.ToString());
            }
        }
        public static void ProcessCheck(StringBuilder sb)
        {
            try
            {
                Process[] processes = Process.GetProcesses();

                foreach (Process pr in processes)
                {
                    if (pr.ProcessName.ToLower().Contains("encompass"))
                        if (pr.MainWindowTitle.ToLower().Contains("error"))
                            pr.CloseMainWindow();
                }
            }
            catch (Exception e)
            {
                sb.AppendLine(e.Message);
                Console.WriteLine(e.Message);
            }
        }

        public static void GarbageCollection(StringBuilder sb)
        {
            DirectoryInfo user = new DirectoryInfo(@"C:\Users");
            var usersFolder = user.GetDirectories().Where(f => !f.FullName.Contains("Public")).OrderByDescending(f => f.LastWriteTime).FirstOrDefault();
            DirectoryInfo tempFolder = new DirectoryInfo(usersFolder + @"\AppData\Local\Temp\Encompass\Temp");
            var newestFolder = tempFolder.GetDirectories().Where(f => !f.FullName.Contains("Assemblies")).OrderByDescending(f => f.CreationTime).FirstOrDefault();
            DirectoryInfo eFolder = new DirectoryInfo(newestFolder + @"\eFolder");
            DirectoryInfo outputPdf = new DirectoryInfo(newestFolder + @"\OutputPdf");
            int count = 0;

            Console.WriteLine("Starting Credit Garbage Cleanup");
            sb.AppendLine("Starting Credit Garbage Cleanup");
            foreach (FileInfo file in eFolder.EnumerateFiles())
            {
                if (file.LastAccessTime < DateTime.Now.AddMinutes(-10))
                {
                    file.Delete();
                    count++;
                }
            }
            Console.WriteLine("eFolder Done");
            sb.AppendLine("eFolder Done");

            foreach (FileInfo file in outputPdf.EnumerateFiles())
            {
                if (file.LastAccessTime < DateTime.Now.AddMinutes(-10))
                {
                    file.Delete();
                    count++;
                }
            }
            Console.WriteLine("OutputPdf Done");
            sb.AppendLine("OutputPdf Done");
            Console.WriteLine(count + " files removed");
            sb.AppendLine(count + " files removed");
        }

        public static bool FileLocked(FileInfo file)
        {
            try
            {
                using (FileStream stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None))
                {
                    stream.Close();
                }
            }
            catch (IOException)
            {
                return false;
            }

            return true;
        }
    }
}