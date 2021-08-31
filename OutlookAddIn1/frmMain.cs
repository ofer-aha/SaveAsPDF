using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.Word;
using System;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using word = Microsoft.Office.Interop.Word;
using System.Linq;
using System.Collections.Generic;

namespace SaveAsPDF1

{
    public partial class frmMain : Form
    {
        //private MailItemRibbon demoRibbon;

        public frmMain()
        {
            InitializeComponent();
        }

        //public frmMain(MailItemRibbon demoRibbon)
        //{
        //    this.demoRibbon = demoRibbon;
        //}

        private void frmMain_Load(object sender, EventArgs e)
        {
            PGlobals.ROOTDRIVE = @"J:\";

            MailItem mailItem = null;
            dynamic windowType = Globals.ThisAddIn.Application.ActiveWindow();
            if (windowType is Outlook.Explorer)
            {
                // frmMain Explorer
                Outlook.Explorer explorer = windowType as Outlook.Explorer;
                mailItem = explorer.Selection[1] as Outlook.MailItem;
                
            }
            else if (windowType is Outlook.Inspector)
            {
                // Read or Compose
                Outlook.Inspector inspector = windowType as Outlook.Inspector;
                mailItem = inspector.CurrentItem as Outlook.MailItem;
                
            }
           
            if (mailItem is Outlook.MailItem)
            {
                lblSubject.Text = mailItem.Subject;
            }
            else
            {
                MessageBox.Show("Not a Mail Item");
                this.Close();
            }
            
         
        }

        

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
           
            if (dialog.ShowDialog() != DialogResult.OK) { return; }

            this.treeView1.Nodes.Add(TraverseDirectory(dialog.SelectedPath));

        }

        private void treeView1_BeforeExpand(object sender, TreeViewCancelEventArgs e)
        {
            TreeNode tn = e.Node.Nodes[0];
            if (tn.Text == "...")
            {
                e.Node.Nodes.AddRange(getFolderNodes(((DirectoryInfo)e.Node.Tag)
                      .FullName, true).ToArray());
                if (tn.Text == "...") tn.Parent.Nodes.Remove(tn);
            }
        }

        static List<TreeNode> getFolderNodes(string dir, bool expanded)
        {
            var dirs = Directory.GetDirectories(dir).ToArray();
            var nodes = new List<TreeNode>();
            foreach (string d in dirs)
            {
                DirectoryInfo di = new DirectoryInfo(d);
                TreeNode tn = new TreeNode(di.Name);
                tn.Tag = di;
                int subCount = 0;
                try { subCount = Directory.GetDirectories(d).Count(); }
                catch { /* ignore accessdenied */  }
                if (subCount > 0) tn.Nodes.Add("...");
                if (expanded) tn.Expand();   //  **
                nodes.Add(tn);
            }
            return nodes;
        }
        private TreeNode TraverseDirectory(string path)
        {
            TreeNode result = new TreeNode(path);
            if (Directory.Exists(path))
            {
                
                foreach (string subdirectory in Directory.GetDirectories(path))
                {
                    result.Nodes.Add(TraverseDirectory(subdirectory));
                }

                return result;
            }
            else
            {
                result.Nodes.Clear();
                return result;
            }
        }

        /// <summary>
        /// Building the project's path
        /// </summary>
        /// <param name="ProjectID"></param>
        /// <returns></returns>
        private static string ProjectPath(string ProjectID)
        {
            string sProjectPath = "";
            
            if (ProjectID != null)
            {
                if (!ProjectID.Contains("-"))
                {
                    //simple pojectid XXX or XXXX
                    sProjectPath =  PGlobals.ROOTDRIVE + (ProjectID.Length == 3 ? "0" + ProjectID : ProjectID).Substring(0,2) + "\\" + ProjectID + "\\"; 
                }
                else
                {
                    //more complicated project id: XXX-X or XXX-XX or XXX-X-XX well you catch my point....
                    string[] split = ProjectID.Split(new char[] { '-' }); //break project id to parts
                    
                    string sRootFolder = PGlobals.ROOTDRIVE + (split[0].Length == 3 ? "0" + split[0] : split[0]).Substring(0, 2) + "\\"; // J:\XX\
                                      //      if exist   J:\XX\XXXX-XX   then    J:\XX\XXXX-X\       else           J:\XX\XXXX\XXXX-X
                    sProjectPath = Directory.Exists(sRootFolder + ProjectID + "\\") ? sRootFolder + ProjectID + "\\" : sRootFolder + split[0] + "\\" + ProjectID + "\\";

                    if (split.Length > 2) //projectID looks like that XXXX-XX-XX 
                     //                                 if exist   J:\XX\XXXX-XX\XXXX-XX-X\                               then       
                    { 
                        sProjectPath = Directory.Exists(sRootFolder + split[0] + "-" + split[1] + "\\" + ProjectID + "\\") ?
                                                         sRootFolder + split[0] + "-" + split[1] + "\\" + ProjectID + "\\" :                     // J:\XX\XXXX-XX\XXXX-XX\  
                                                         sRootFolder + split[0] + "\\" + split[0] + "-" + split[1] + "\\" + ProjectID + "\\";    //else J:\XX\XXXX\XXXX-X\XXXX-XX\
                    }
                }
            }
            return sProjectPath;
        }

        

        private void txtProjectID_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Space)
            {
                string sPath = ProjectPath(txtProjectID.Text);
                if (Directory.Exists(sPath))
                {
                    lblFolder.Text = sPath;

                    treeView1.Nodes.Clear();
                    Cursor.Current = Cursors.WaitCursor; //show the user we are soning somthing..... "the compuer is thinking" 
                    treeView1.Nodes.Add(TraverseDirectory(lblFolder.Text));
                    //this.treeView1.Nodes.AddRange(getFolderNodes(lblFolder.Text, false).ToArray());
                    Cursor.Current = Cursors.Default; 
                }
                else
                {
                    lblFolder.Text = "";
                    this.treeView1.Nodes.Clear();
                    txtProjectID.Clear();
                }

            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {

            Outlook.MailItem mailItem = null;
            dynamic windowType = Globals.ThisAddIn.Application.ActiveWindow();
            if (windowType is Outlook.Explorer)
            {
                // Main Explorer
                Explorer explorer = windowType as Outlook.Explorer;
                mailItem = explorer.Selection[1] as Outlook.MailItem;
            }
            else if (windowType is Outlook.Inspector)
            {
                // Read or Compose
                Inspector inspector = windowType as Outlook.Inspector;
                mailItem = inspector.CurrentItem as Outlook.MailItem;
            }

            //cretae uniq timestamp for uniqe fimenames

            string TimeStamp = DateTime.Now.ToString("yyyyMMddHHmmssffff", CultureInfo.CurrentUICulture);

            string tFilename = @Path.GetTempPath() + TimeStamp + ".mht";

            mailItem.SaveAs(@tFilename, OlSaveAsType.olMHTML);
            


            word.Application oWord = new word.Application();
            word.Document oDOC = oWord.Documents.Open(@tFilename, true);

            object misValue = System.Reflection.Missing.Value;

            string oFileName = @lblFolder.Text + TimeStamp + "_" + SafeFileName(mailItem.Subject) + ".pdf";

            ConvertPDF(oDOC, misValue, @oFileName);

        }

        private static void ConvertPDF(Document oDOC, object misValue, string oFileName)
        {
            oDOC.ExportAsFixedFormat(@oFileName,
                            word.WdExportFormat.wdExportFormatPDF,
                            false,                  //open after export
                            word.WdExportOptimizeFor.wdExportOptimizeForPrint,
                            word.WdExportRange.wdExportAllDocument, (int)misValue, (int)misValue, WdExportItem.wdExportDocumentWithMarkup, false,
                             true, WdExportCreateBookmarks.wdExportCreateWordBookmarks, true, true, false, misValue);
        }

        private static string SafeFileName(string inTXT)
        {
            string pattern = @"[\\/:*?""<>|]";
            //Regex rg = new Regex(pattern);
            
            return Regex.Replace(inTXT, pattern, "").Trim();
        }

        private void button2_Click(object sender, EventArgs e)
        {
           frmSettings  frm = new frmSettings();
            frm.Show();
        }
    }
}
