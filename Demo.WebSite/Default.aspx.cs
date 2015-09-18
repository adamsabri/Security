using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using ExcelEngine;

namespace Demo.WebSite
{
    public partial class _Default : Page
    {
        private string _path;
        private string _fileNamePath;

        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void ImportBtn_Click(object sender, EventArgs e)
        {
            MsgLabel.Text = string.Empty;
            _path = UploadFile();
           var th1 =  new Thread(() =>
               {
                   var start = DateTime.Now;
                   var reader = new Reader(_path.Replace("test", "test1"));
                   var list = reader.ConvertToList();
                   MsgLabel.Text += "Thread file imported: " + list.Count() + " rows at " + (DateTime.Now - start) + ".\n";
                   Debug.WriteLine("thread 1: " + MsgLabel.Text);
               });
            th1.Start();
         
            th1.Join();
        }

        private void Import()
        {
            
        }


        protected string UploadFile()
        {
                var path = HttpContext.Current.Server.MapPath("~/Uploads/" + FileUpload.FileName);
                FileUpload.PostedFile.SaveAs(path);
            return path;
        }
    }
}