namespace WordGit
{
    internal static class RibbonIcons
    {
        // The WPS icon gallery draws these at the same size and in the same theme as the rest of the ribbon.
        internal static string ForWps(string markup)
        {
            return markup
                .Replace("imageMso=\"ViewOutlineView\"", "imageMso=\"FileProperties\"")
                .Replace("imageMso=\"FilePrintPreview\"", "imageMso=\"FileOpen\"")
                .Replace("imageMso=\"FileNew\"", "imageMso=\"FileSaveAs\"")
                .Replace("imageMso=\"WindowSwitchWindowsMenuWord\"", "imageMso=\"Refresh\"")
                .Replace("imageMso=\"ReviewAcceptChange\"", "imageMso=\"HyperlinkInsert\"")
                .Replace("imageMso=\"ReviewRejectChange\"", "imageMso=\"Delete\"")
                .Replace("imageMso=\"ReviewCompareMenu\"", "imageMso=\"Copy\"");
        }
    }
}
