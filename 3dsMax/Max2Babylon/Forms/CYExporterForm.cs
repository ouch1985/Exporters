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
            this.Size = new System.Drawing.Size(600, 600);
            this.MinimumSize = new System.Drawing.Size(600, 600);
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.saveFileDialog.DefaultExt = "cy";
            this.saveFileDialog.Filter = "cy files|*.cy";

            webBrowser = new WebBrowser();
            this.Controls.Add(webBrowser);

            //webBrowser.Navigate("http://129.204.129.112:9999/max.html?t=" + new DateTime().Millisecond);
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

        public void SelectFile(String format) {
            if (saveFileDialog.ShowDialog(this) == DialogResult.OK)

            {
                ShowMessage("updateFile", saveFileDialog.FileName);
            }
        }

        public void Export(String message)
        {
            if (ing)
            {
                ShowMessage("error", "进行中,请勿重复点击");
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
                webBrowser.Document.InvokeScript("$cy", new String[] { type, message });
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
                success = true;
            }
            catch (Exception ex)
            {
                err = ex.Message;
            }


            try {
                ShowMessage("done", success ? "1" : "0");
                if(err != null)
                {
                    ShowMessage("error", err);
                }

            } catch (Exception e) { 
            }


            BringToFront();
            ing = false;
        }
    }
}
