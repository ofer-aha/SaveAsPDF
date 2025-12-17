namespace SaveAsPDF.Helpers
{
    public static class StringExtensions
    {
        // existing SafeProjectID() etc...

        public static string SafeExtractProjectId(this string subject)
        {
            if (string.IsNullOrWhiteSpace(subject))
                return string.Empty;

            // Simple strategy: first whitespace‑separated token that looks like a project id
            var parts = subject.Split(new[] { ' ', '\t', '-', '_', ':', ';', ',' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var part in parts)
            {
                if (part.SafeProjectID())
                    return part;
            }

            return string.Empty;
        }
    }
}