using System.Windows.Forms;



namespace SaveAsPDF
{
    public enum CustomMessageBoxButtons { OK, OKCancel, YesNo, YesNoCancel, RetryCancel }
    public enum CustomMessageBoxIcon { None, Information, Warning, Error, Question }
    public enum CustomMessageAlignment { Left, Center, Right }
    public enum CustomMessageLanguage { Hebrew, English }

    public static class CustomMessageBox
    {
        public static DialogResult Show(string text, string caption, CustomMessageBoxButtons buttons, CustomMessageBoxIcon icon, CustomMessageAlignment alignment = CustomMessageAlignment.Right, CustomMessageLanguage language = CustomMessageLanguage.Hebrew)
        {
            using (var msgBox = new CustomMessageBoxForm(text, caption, buttons, icon, alignment, language))
            {
                return msgBox.ShowDialog();
            }
        }
    }
}


// Usage example:
// var result = CustomMessageBox.Show("האם אתה בטוח שברצונך לצאת?", "אישור יציאה", CustomMessageBoxButtons.YesNo, CustomMessageBoxIcon.Question);






