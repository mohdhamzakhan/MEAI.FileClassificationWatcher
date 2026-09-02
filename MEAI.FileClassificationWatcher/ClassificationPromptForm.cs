using System;
using System.Drawing;
using System.Windows.Forms;

namespace MEAI.FileClassificationWatcher
{
    // Modal dialog asking the user to pick a classification level. Reused for:
    //  - a brand new document/workbook/presentation (mandatory, no Cancel)
    //  - an existing file opened with no classification yet (mandatory, no Cancel)
    //  - the confirm-before-close prompt (Cancel aborts the close instead)
    public class ClassificationPromptForm : Form
    {
        private readonly RadioButton _rbTopSecret = new() { Text = "Top Secret", AutoSize = true };
        private readonly RadioButton _rbSecret = new() { Text = "Secret", AutoSize = true };
        private readonly RadioButton _rbConfidential = new() { Text = "Confidential", AutoSize = true };
        private readonly RadioButton _rbPublic = new() { Text = "Public", AutoSize = true };
        private readonly Button _btnOk = new() { Text = "Confirm", DialogResult = DialogResult.OK, AutoSize = true };
        private readonly Button _btnCancel;

        public ClassificationLevel? SelectedLevel { get; private set; }

        public ClassificationPromptForm(string documentName, string headline, ClassificationLevel? current = null, bool allowCancel = true)
        {
            Text = "Document Classification";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MinimizeBox = false;
            MaximizeBox = false;
            ControlBox = allowCancel;
            TopMost = true;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            Padding = new Padding(0, 0, 20, 20);

            const int contentWidth = 320;

            // AutoSize = true + a MaximumSize cap makes the label wrap and grow DOWNWARD
            // to fit however many lines headline + filename end up being. A fixed
            // Size(320, 50) with AutoSize = false silently clips whatever text doesn't
            // fit — which was always the filename, since it's the last line.
            var lbl = new Label
            {
                Text = $"{headline}\n\n\"{documentName}\"",
                AutoSize = true,
                MaximumSize = new Size(contentWidth, 0),
                Location = new Point(20, 15)
            };
            Controls.Add(lbl); // must be parented before reading lbl.Height — AutoSize measures on add

            // Position everything below the label relative to its actual measured height,
            // instead of hardcoded Y values, so this keeps working no matter how long the
            // filename or headline is.
            int y = lbl.Bottom + 15;
            _rbTopSecret.Location = new Point(30, y);
            _rbSecret.Location = new Point(30, y + 25);
            _rbConfidential.Location = new Point(30, y + 50);
            _rbPublic.Location = new Point(30, y + 75);

            foreach (var rb in new[] { _rbTopSecret, _rbSecret, _rbConfidential, _rbPublic })
                Controls.Add(rb);

            // Pre-select the current level if we know it; otherwise default to
            // Confidential rather than silently defaulting to Public.
            if (current == ClassificationLevel.TopSecret) _rbTopSecret.Checked = true;
            else if (current == ClassificationLevel.Secret) _rbSecret.Checked = true;
            else if (current == ClassificationLevel.Public) _rbPublic.Checked = true;
            else _rbConfidential.Checked = true;

            int buttonY = _rbPublic.Bottom + 20;
            _btnOk.Location = new Point(130, buttonY);
            _btnOk.Click += (_, _) =>
            {
                SelectedLevel =
                    _rbTopSecret.Checked ? ClassificationLevel.TopSecret :
                    _rbSecret.Checked ? ClassificationLevel.Secret :
                    _rbConfidential.Checked ? ClassificationLevel.Confidential :
                    ClassificationLevel.Public;
            };
            Controls.Add(_btnOk);

            _btnCancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, AutoSize = true };
            _btnCancel.Location = new Point(_btnOk.Right + 10, buttonY); // relative to OK's width, no fixed overlap risk
            _btnCancel.Enabled = allowCancel;
            _btnCancel.Visible = allowCancel;
            Controls.Add(_btnCancel);

            AcceptButton = _btnOk;
            if (allowCancel) CancelButton = _btnCancel;

            // ControlBox = false removes the X button, but doesn't reliably stop every way
            // of force-closing a window (Alt+F4, right-click the taskbar entry > Close).
            // When this prompt is meant to be mandatory (no Cancel), block any close that
            // didn't come from clicking Confirm, so SelectedLevel can never end up null here.
            FormClosing += (_, e) =>
            {
                if (!allowCancel && DialogResult != DialogResult.OK)
                    e.Cancel = true;
            };
        }
    }
}