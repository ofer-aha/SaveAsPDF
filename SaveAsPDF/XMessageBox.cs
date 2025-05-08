using System.Windows.Forms;



namespace SaveAsPDF
{
    public enum XMessageBoxButtons { OK, OKCancel, YesNo, YesNoCancel, RetryCancel }
    public enum XMessageBoxIcon { None, Information, Warning, Error, Question }
    public enum XMessageAlignment { Left, Center, Right }
    public enum XMessageLanguage { Hebrew, English }

    public static class XMessageBox
    {
        public static DialogResult Show(string text, string caption, XMessageBoxButtons buttons, XMessageBoxIcon icon, XMessageAlignment alignment = XMessageAlignment.Right, XMessageLanguage language = XMessageLanguage.Hebrew)
        {
            using (var msgBox = new XMessageBoxForm(text, caption, buttons, icon, alignment, language))
            {
                return msgBox.ShowDialog();
            }
        }
    }
}


// Usage example:
// var result = XMessageBox.Show("האם אתה בטוח שברצונך לצאת?", "אישור יציאה", XMessageBoxButtons.YesNo, XMessageBoxIcon.Question);






