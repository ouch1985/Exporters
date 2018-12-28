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
        private SaveFileDialog saveFileDialog;
        private String expParams = null;

        public CYExporterForm(BabylonExportActionItem babylonExportAction)
        {
            this.Text = "模型导出";
            this.Size = new System.Drawing.Size(Screen.PrimaryScreen.Bounds.Width / 4, Screen.PrimaryScreen.Bounds.Height / 4);
            this.MinimumSize = new System.Drawing.Size(Screen.PrimaryScreen.Bounds.Width/4, Screen.PrimaryScreen.Bounds.Height/4);
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.saveFileDialog.DefaultExt = "glb";
            this.saveFileDialog.Filter = "glb files|*.glb";

            webBrowser = new WebBrowser();
            this.Controls.Add(webBrowser);

            webBrowser.Navigate("http://192.168.31.31:8080/max.html?t=" + new DateTime().Millisecond);
            webBrowser.Dock = DockStyle.Fill;

            this.Load += new EventHandler(Form_Load);


        }

        void Form_Load(object sender, EventArgs e)
        {
            webBrowser.AllowWebBrowserDrop = false;
            webBrowser.IsWebBrowserContextMenuEnabled = false;
            webBrowser.WebBrowserShortcutsEnabled = false;
            webBrowser.ObjectForScripting = this;
        }

        public void SelectFile() {
            if (saveFileDialog.ShowDialog(this) == DialogResult.OK)
            {
                webBrowser.Document.InvokeScript("updateFile", new String[] { saveFileDialog.FileName });
            }
        }

        public void Export(String message)
        {
            if (ing)
            {
                webBrowser.Document.InvokeScript("done", new String[] { "0", "进行中,请勿重复点击"});
                return;
            }
            ing = true;

            expParams = message;
            Thread t = new Thread(new ThreadStart(DoExport));
            t.IsBackground = true;
            t.Start();

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
                webBrowser.Document.InvokeScript(type, new String[] { message });
            }
        }

        private void DoExport()
        {


            bool success = false;
            String err = null;
            try
            {
                StringReader sr = new StringReader(expParams);
                JsonSerializer serializer = new JsonSerializer();
                ExportParameters exportParameters = (ExportParameters)serializer.Deserialize(new JsonTextReader(sr), typeof(ExportParameters));

                //webBrowser.Document.InvokeScript("log", new String[] { "开始..." });
                ShowMessage("log", "开始...");

                exporter = new BabylonExporter();

                exporter.OnImportProgressChanged += progress =>
                {
                    //webBrowser.Document.InvokeScript("progress", new String[] { progress.ToString() });
                    ShowMessage("progress", progress.ToString());
                };

                //ExportParameters exportParameters = new ExportParameters
                //{
                //    outputPath = message,
                //    outputFormat = "glb",
                //    exportTangents = false
                //};

                exporter.callerForm = this;

                exporter.Export(exportParameters);
                success = true;
            }
            catch (Exception ex)
            {
                err = ex.Message;
            }


            try {
                //webBrowser.Document.InvokeScript("done", new String[] { success ? "1" : "0", err});
                ShowMessage("done", success ? "1" : "0");
                if(err != null)
                {
                    ShowMessage("log", err);
                }

            } catch (Exception e) { 
            }


            BringToFront();
            ing = false;
        }
    }
}
