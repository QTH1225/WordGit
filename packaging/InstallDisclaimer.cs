using System;
using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace WordGitSetup
{
    internal sealed class InstallDisclaimer : Form
    {
        internal const string DocumentPath = "docs/INSTALL-DISCLAIMER.md";

        internal InstallDisclaimer(string payloadRoot)
        {
            var lines = File.ReadAllLines(Path.Combine(payloadRoot, DocumentPath));
            var title = lines[0].Substring("### ".Length);
            Text = "WordGit 安装";
            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Font;
            StartPosition = FormStartPosition.CenterParent;
            ShowInTaskbar = false;
            MinimizeBox = MaximizeBox = false;
            ClientSize = new Size(690, 480);
            MinimumSize = new Size(500, 370);

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Padding = new Padding(18)
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(new Label
            {
                Text = title, AutoSize = true, Margin = new Padding(0, 0, 0, 12),
                Font = new Font(Font.FontFamily, 16, FontStyle.Bold)
            }, 0, 0);

            var body = new RichTextBox
            {
                AccessibleName = "免责声明正文", Dock = DockStyle.Fill, Margin = Padding.Empty,
                Font = this.Font, ReadOnly = true, BackColor = SystemColors.Window, DetectUrls = false,
                ScrollBars = RichTextBoxScrollBars.Vertical
            };
            FormatBody(body, lines);
            layout.Controls.Add(body, 0, 1);

            var footer = new TableLayoutPanel
            {
                AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Dock = DockStyle.Fill,
                ColumnCount = 4, RowCount = 1, Margin = new Padding(0, 12, 0, 0)
            };
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            var consent = new CheckBox
            {
                Text = "我同意", AutoSize = true, Anchor = AnchorStyles.Left, Margin = Padding.Empty
            };
            var proceed = new Button
            {
                Text = "继续安装", AutoSize = true, Enabled = false, Padding = new Padding(8, 3, 8, 3)
            };
            var cancel = new Button
            {
                Text = "取消安装", AutoSize = true, DialogResult = DialogResult.Cancel, Padding = new Padding(8, 3, 8, 3)
            };
            consent.CheckedChanged += (sender, args) => proceed.Enabled = consent.Checked;
            proceed.Click += (sender, args) => { if (consent.Checked) DialogResult = DialogResult.OK; };
            footer.Controls.Add(consent, 0, 0);
            footer.Controls.Add(proceed, 2, 0);
            footer.Controls.Add(cancel, 3, 0);
            layout.Controls.Add(footer, 0, 2);

            Controls.Add(layout);
            AcceptButton = proceed;
            CancelButton = cancel;
            ActiveControl = body;
        }

        private static void FormatBody(RichTextBox body, string[] lines)
        {
            using (var bold = new Font(body.Font, FontStyle.Bold))
            {
                for (var index = 1; index < lines.Length; index++)
                {
                    var line = lines[index].TrimEnd('\\');
                    var heading = line.StartsWith("#### ", StringComparison.Ordinal);
                    if (heading) line = line.Substring("#### ".Length);
                    if (line.StartsWith("- ", StringComparison.Ordinal)) line = "• " + line.Substring(2);
                    foreach (var span in Regex.Split(line, @"(\*\*[^*]+\*\*)"))
                    {
                        var emphasized = span.StartsWith("**", StringComparison.Ordinal);
                        body.SelectionFont = heading || emphasized ? bold : body.Font;
                        body.AppendText((emphasized ? span.Substring(2, span.Length - 4) : span).Replace("`", ""));
                    }
                    body.AppendText(Environment.NewLine);
                }
            }
            body.Select(0, 0);
            body.ScrollToCaret();
        }
    }
}
