using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security.Permissions;
using System.Security.Principal;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ExcelEngine;
using ExcelEngine.NPOI;
using Security;

namespace Demo
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            try
            {
                TestExcel();
                //TestSecurity();
                Console.ReadLine();
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        private static void TestSecurity()
        {
            //AppDomain.CurrentDomain.SetPrincipalPolicy(PrincipalPolicy.WindowsPrincipal);
            var enc = Cryptography.Md5Encrypt("kha2015");
            Console.WriteLine(enc);
            Console.WriteLine(Cryptography.Md5Decrypt(enc));

            ValidateLogin("kha", "password253");
        }

        private static void TestExcel()
        {
            //Write();
            //Read();
            ThreadsRead();
        }

        private static void Read()
        {
            //var f1 = @"D:\Documents\Tutorials\Microsoft\Common\Excel Import Export\Lib\npoi-master\examples\hssf\ImportXlsToDataTable\xls\Book1.xls";
            //var f2 = @"C:\Users\kha\Desktop\BigFile.xlsx";
            //var f3 = @"C:\Users\kha\Desktop\sample.xls";
            //var f4 = @"C:\Users\kha\Desktop\Blackfire-Stock.xls";
            //var f = @"C:\Users\kha\Desktop\test.xlsx";
            var f = @"C:\Users\kha\Desktop\Classeur1.xlsx";
            var start = DateTime.Now;
            Console.WriteLine("Start reading : " + DateTime.Now);

            //var reader = new Reader();
            //reader.InitializeWorkbook(f5);
            //var list = reader.ConvertToList().Reverse().ToList();
            //foreach (var val in list)
            //{
            //    Console.WriteLine(val);
            //}
            //Console.WriteLine("Total : " + list.Count);
            //Console.WriteLine("Reading finished : " + (DateTime.Now - start));

            var reader = new Reader(f, XlsxReaderType.Epplus);
            var list = reader.ConvertToList();
            Console.WriteLine("file imported: " + list.Count() + " rows at " + (DateTime.Now - start) + ".\n");
        }

        private static void Write()
        {
            var start = DateTime.Now;
            Console.WriteLine("Start Write : " + DateTime.Now);
            var f4 = @"C:\Users\kha\Desktop\test.xls";
            Writer.Instance.CreateWorkBook(f4);
            Console.WriteLine("CreateFile finished: " + (DateTime.Now - start));
        }

        private static void ValidateLogin(string login, string password)
        {

            Console.WriteLine("Thread.CurrentPrincipal.Identity.Name = " + Thread.CurrentPrincipal.Identity.Name);

            SecurityManager.SetCurrentPrincipal(new CustomIdentity(login, "Custom") { Id = 1, FirstName = "khamis", LastName = "hajji" },
                                                 new[] {"Admin", "Dev"});


            Console.WriteLine("User Info :\n");
            Console.WriteLine("Thread.CurrentPrincipal.Identity.Name = " + Thread.CurrentPrincipal.Identity.Name);
            Console.WriteLine("SecurityManager.CurrentIdentity.Name = " + SecurityManager.CurrentIdentity.Name);
            Console.WriteLine("SecurityManager.CurrentIdentity.Id = " + SecurityManager.CurrentIdentity.Id);
            Console.WriteLine("SecurityManager.CurrentIdentity.FirstName = " + SecurityManager.CurrentIdentity.FirstName);
            Console.WriteLine("SecurityManager.CurrentIdentity.LastName = " + SecurityManager.CurrentIdentity.LastName);
        }

        private static void ThreadsRead()
        {
            //var f1 = @"D:\Documents\Tutorials\Microsoft\Common\Excel Import Export\Lib\npoi-master\examples\hssf\ImportXlsToDataTable\xls\Book1.xls";
            //var f2 = @"C:\Users\kha\Desktop\BigFile.xlsx";
            //var f3 = @"C:\Users\kha\Desktop\sample.xls";
            //var f4 = @"C:\Users\kha\Desktop\Blackfire-Stock.xls";
            var f = @"C:\Users\kha\Desktop\test.xlsx";
            //var start = DateTime.Now;
            Console.WriteLine("Start reading : " + DateTime.Now);
            var listThs = new List<Thread>();
            for (var i = 1; i < 2; i++)
            {
                listThs.Add(new Thread(() =>
                {
                    var start = DateTime.Now;
                    var reader = new Reader(f.Replace("test", "Import\\test1"), XlsxReaderType.Npoi);
                    var list = reader.ConvertToList();
                    Console.WriteLine(string.Format("{0} file imported: {1} rows at {2}.\n",Thread.CurrentThread.Name, list.Count(), (DateTime.Now - start)));
                }) { Name = "Thread " + i });
            }
            foreach (var thread in listThs)
            {
                thread.Start();
            }
            listThs[0].Join();
        }
    }
}
