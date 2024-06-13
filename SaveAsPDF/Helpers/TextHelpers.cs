using System;
using System.Text.RegularExpressions;

namespace SaveAsPDF.Helpers
{
    public static class TextHelpers
    {
        /// <summary>
        /// convert the file size to "human" readable figures 
        /// </summary>
        /// <param name="byteCount"></param>
        /// <returns>String XXKB/MB/GB</returns>
        public static String BytesToString(this int byteCount)
        {
            string[] suf = { " B", " KB", " MB", " GB", " TB", " PB", " EB" }; //Longs run out around EB
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
        /// <summary>
        ///  extract a substring between two specified strings
        /// </summary>
        /// <param name="strSource"></param>
        /// <param name="strStart"></param>
        /// <param name="strEnd"></param>
        /// <returns></returns>
        public static string GetBetween(string strSource, string strStart, string strEnd)
        {
            int startIndex = strSource.IndexOf(strStart);
            if (startIndex == -1)
                return ""; // Start string not found

            int endIndex = strSource.IndexOf(strEnd, startIndex + strStart.Length);
            if (endIndex == -1)
                return ""; // End string not found

            return strSource.Substring(startIndex + strStart.Length, endIndex - startIndex - strStart.Length);
        }
    }
}
