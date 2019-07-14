using System;
using System.Windows.Forms;
using System.Security.Permissions;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.IO;
using System.Threading;

namespace Max2Babylon
{
    [PermissionSet(SecurityAction.Demand, Name = "FullTrust")]
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    public class CYExporterForm : Form
    {
        private WebBrowser webBrowser;
        private BabylonExporter exporter;
        private bool ing = false;
        private String expParams = null;
        private String ouputfile = null;


        //private String cache = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "cy", "cy.cache");
        private String cache = System.IO.Path.Combine(Application.LocalUserAppDataPath, "cy.cache");
        public CYExporterForm(BabylonExportActionItem babylonExportAction)
        {

            this.Text = "模型导出";
            this.Size = new System.Drawing.Size(600, 600);
            this.MinimumSize = new System.Drawing.Size(600, 600);
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;

            webBrowser = new WebBrowser();
            this.Controls.Add(webBrowser);

            //webBrowser.Navigate("https://yuojiu.com/p3dsmax.html?t=" + new DateTime().Millisecond);
            //webBrowser.Navigate("https://test.jiuyuo.com/p3dsmax.html?t=" + new DateTime().Millisecond);
            webBrowser.Navigate("http://192.168.1.111:8080/p3dsmax.html?t=" + new DateTime().Millisecond);
            //webBrowser.Navigate("http://192.168.31.31:8080/max.html?t=" + new DateTime().Millisecond);
            webBrowser.Dock = DockStyle.Fill;

            this.Load += new EventHandler(Form_Load);

            this.Activated += new System.EventHandler(this.ExporterForm_Activated);
            this.Deactivate += new System.EventHandler(this.ExporterForm_Deactivate);
        }


        private void ExporterForm_Activated(object sender, EventArgs e)
        {
            Loader.Global.DisableAccelerators();
        }

        private void ExporterForm_Deactivate(object sender, EventArgs e)
        {
            Loader.Global.EnableAccelerators();
        }

        void Form_Load(object sender, EventArgs e)
        {
            webBrowser.AllowWebBrowserDrop = false;
            webBrowser.IsWebBrowserContextMenuEnabled = false;
            webBrowser.WebBrowserShortcutsEnabled = false;
            webBrowser.ObjectForScripting = this;

        }

        public void Export(String message)
        {
            if (ing)
            {
                ShowMessage("error", "进行中,请勿重复点击");
                return;
            }
            ing = true;
            ouputfile = System.IO.Path.GetTempFileName().Replace(".tmp", ".cy");
            ouputfile = Path.GetDirectoryName(ouputfile);
            ouputfile = Path.Combine(ouputfile, "model.cy");

            expParams = message;
            Thread t = new Thread(new ThreadStart(DoExport));
            t.IsBackground = true;
            t.Start();

        }

        public String WriteCache(String data) {
            try {
                System.IO.File.WriteAllText(cache, data);
                return "ok," + cache;
            } catch (Exception e) {
                return "fail," + e.Message;
            }
        }

        public String ReadCache() {

            try {
                return System.IO.File.ReadAllText(cache);
            } catch (Exception e) {
                return "";
            }
        }

        private delegate void ShowMessageDelegate(string type, string message);
        private void ShowMessage(string type, string message)
        {
            if (this.InvokeRequired)
            {
                ShowMessageDelegate showMessageDelegate = ShowMessage;
                this.Invoke(showMessageDelegate, new object[] { type, message });
            }
            else
            {
                webBrowser.Document.InvokeScript("$cy", new String[] { type, message });
            }

        }

        private void DoExport()
        {
            bool success = false;
            String err = null;
            ExportParameters exportParameters = null;
            try
            {
                StringReader sr = new StringReader(expParams);
                JsonSerializer serializer = new JsonSerializer();
                exportParameters = (ExportParameters)serializer.Deserialize(new JsonTextReader(sr), typeof(ExportParameters));

                ShowMessage("warning", "临时文件" + ouputfile);
                exportParameters.outputPath = ouputfile;

                ShowMessage("log", "开始...");

                exporter = new BabylonExporter();

                exporter.OnImportProgressChanged += progress =>
                {
                    ShowMessage("progress", progress.ToString());
                };

                exporter.OnWarning += (warning, rank) =>
                {

                    ShowMessage("warning", warning);
                };

                exporter.OnError += (error, rank) =>
                {
                    ShowMessage("error", error);
                };

                exporter.OnMessage += (message, color, rank, emphasis) =>
                {
                    ShowMessage("log", message);
                };

                exporter.callerForm = this;

                exporter.Export(exportParameters);
                ShowMessage("log", "开始上传");


                //上传
                System.Net.WebClient client = new System.Net.WebClient();

                client.UploadProgressChanged += (sender, evt) => {
                    ShowMessage("progress", evt.ProgressPercentage + "");
                };

                byte[] response = client.UploadFile(exportParameters.url, "POST", ouputfile);
                client.Dispose();

                ShowMessage("log", "上传结束");
                string uploadResp = System.Text.Encoding.UTF8.GetString(response);
                Resp resp = JsonConvert.DeserializeObject< Resp>(uploadResp);
                if (resp.success) {
                    success = true;
                    ShowMessage("log", "上传成功");
                }
                else {
                    ShowMessage("log", "上传失败");
                    success = false;
                    err = resp.msg;
                }
            }
            catch (Exception ex)
            {
                err = ex.Message;
            }


            try {
                ShowMessage("done", success ? "1" : "0");
                if (err != null)
                {
                    ShowMessage("error", err);
                }

            } catch (Exception eee) {
                ShowMessage("error", eee.Message);
            }


            BringToFront();

            //try
            //{
            //    HtmlElement fileInput = webBrowser.Document.GetElementById("file");
            //    //ShowMessage("warning", "找到文件" + fileInput.Id);
            //    //fileInput.Focus();
            //    //ShowMessage("warning", "Focus");
            //    //SendKeys.SendWait(exportParameters.outputPath);

            //}
            //catch (Exception ee)
            //{
            //    ShowMessage("error", ee.Message);
            //}


            ing = false;
        }
    }

    public class Resp {
        public bool success;
        public string msg;
    }
}
