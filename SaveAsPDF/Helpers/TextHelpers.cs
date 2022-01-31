using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SaveAsPDF.Helpers
{

    public static class TextHelpers
    {
        /// <summary>
        /// convert the file size to "humen" readle figurs 
        /// </summary>
        /// <param name="byteCount"></param>
        /// <returns>String XXKB/MB/GB</returns>
        public static String BytesToString(this int byteCount)
        {
            string[] suf = { "B", "KB", "MB", "GB", "TB", "PB", "EB" }; //Longs run out around EB
            if (byteCount == 0)
            {
                return "0" + suf[0];
            }

            long bytes = Math.Abs(byteCount);
            int place = Convert.ToInt32(Math.Floor(Math.Log(bytes, 1024)));
            double num = Math.Round(bytes / Math.Pow(1024, place), 1);
            return (Math.Sign(byteCount) * num).ToString() + suf[place];
        }

        /// <summary>
        /// validate the email address format only if text.length>0
        /// </summary>
        /// <param name="s"></param>
        /// <returns>Default: True </returns>
        public static bool ValidMail(this string s)
        {
            bool output = true;
            if (s.Length > 0)
            {
                var regex = new Regex(@"^([0-9a-zA-Z]([\+\-_\.][0-9a-zA-Z]+)*)+@(([0-9a-zA-Z][-\w]*[0-9a-zA-Z]*\.)+[a-zA-Z0-9]{2,17})$");
                output = regex.IsMatch(s);
            }
            return output;
        }

        public static string GetBetween(string strSource, string strStart, string strEnd)
        {
            if (strSource.Contains(strStart) && strSource.Contains(strEnd))
            {
                int Start, End;
                Start = strSource.IndexOf(strStart, 0) + strStart.Length;
                End = strSource.IndexOf(strEnd, Start);
                return strSource.Substring(Start, End - Start);
            }
            return "";

        }

        public static void EnableContextMenu(this RichTextBox rtb)
        {
            if (rtb.ContextMenuStrip == null)
            {
                // Create a ContextMenuStrip without icons
                ContextMenuStrip cms = new ContextMenuStrip();
                cms.ShowImageMargin = false;

                // 3. Add the Cut option (cuts the selected text inside the richtextbox)
                ToolStripMenuItem tsmiCut = new ToolStripMenuItem("חתוך");
                tsmiCut.Click += (sender, e) => rtb.Cut();
                cms.Items.Add(tsmiCut);

                // 4. Add the Copy option (copies the selected text inside the richtextbox)
                ToolStripMenuItem tsmiCopy = new ToolStripMenuItem("העתק");
                tsmiCopy.Click += (sender, e) => rtb.Copy();
                cms.Items.Add(tsmiCopy);

                // 5. Add the Paste option (adds the text from the clipboard into the richtextbox)
                ToolStripMenuItem tsmiPaste = new ToolStripMenuItem("הדבק");
                tsmiPaste.Click += (sender, e) => rtb.Paste();
                cms.Items.Add(tsmiPaste);

                // 6. Add the Delete Option (remove the selected text in the richtextbox)
                ToolStripMenuItem tsmiDelete = new ToolStripMenuItem("מחק");
                tsmiDelete.Click += (sender, e) => rtb.SelectedText = "";
                cms.Items.Add(tsmiDelete);

                // Add a Separator
                cms.Items.Add(new ToolStripSeparator());

                // 7. Add the Select All Option (selects all the text inside the richtextbox)
                ToolStripMenuItem tsmiSelectAll = new ToolStripMenuItem("בחר הכל");
                tsmiSelectAll.Click += (sender, e) => rtb.SelectAll();
                cms.Items.Add(tsmiSelectAll);

                // Add a Separator
                cms.Items.Add(new ToolStripSeparator());

                // 1. Add the Undo option
                ToolStripMenuItem tsmiUndo = new ToolStripMenuItem("בטל");
                tsmiUndo.Click += (sender, e) => rtb.Undo();
                cms.Items.Add(tsmiUndo);

                // 2. Add the Redo option
                ToolStripMenuItem tsmiRedo = new ToolStripMenuItem("בצע שוב");
                tsmiRedo.Click += (sender, e) => rtb.Redo();
                cms.Items.Add(tsmiRedo);




                // When opening the menu, check if the condition is fulfilled 
                // in order to enable the action
                cms.Opening += (sender, e) =>
                {
                    tsmiUndo.Enabled = !rtb.ReadOnly && rtb.CanUndo;
                    tsmiRedo.Enabled = !rtb.ReadOnly && rtb.CanRedo;
                    tsmiCut.Enabled = !rtb.ReadOnly && rtb.SelectionLength > 0;
                    tsmiCopy.Enabled = rtb.SelectionLength > 0;
                    tsmiPaste.Enabled = !rtb.ReadOnly && Clipboard.ContainsText();
                    tsmiDelete.Enabled = !rtb.ReadOnly && rtb.SelectionLength > 0;
                    tsmiSelectAll.Enabled = rtb.TextLength > 0 && rtb.SelectionLength < rtb.TextLength;
                };

                rtb.ContextMenuStrip = cms;
            }
        }
    }
}
