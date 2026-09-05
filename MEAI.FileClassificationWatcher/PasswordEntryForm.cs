using System;
using System.Drawing;
using System.Windows.Forms;

namespace MEAI.FileClassificationWatcher
{
    // Collects a password from the user to protect a Secret/TopSecret document. Unlike
    // ClassificationPromptForm, this is deliberately NOT reused for anything else — a
    // password field needs masking, confirmation-match validation, and inline error
    // display that a simple radio-button prompt doesn't.
    public class PasswordEntryForm : Form
    {
        private readonly TextBox _txtPassword = new() { UseSystemPasswordChar = true, Width = 220 };
        private readonly TextBox _txtConfirm = new() { UseSystemPasswordChar = true, Width = 220 };
        private readonly Label _lblError = new() { ForeColor = Color.Firebrick, AutoSize = true, Visible = false };
        private readonly Button _btnOk = new() { Text = "Set Password", DialogResult = DialogResult.None, AutoSize = true };
        private readonly Button _btnCancel;

        public string? EnteredPassword { get; private set; }

        public PasswordEntryForm(string documentName, string levelDisplayName, bool allowCancel)
        {
            Text = "Set Document Password";
            Icon = BrandIcon.Create();
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MinimizeBox = false;
            MaximizeBox = false;
            ControlBox = allowCancel; // hides the window's X button too when mandatory —
                                      // just disabling the Cancel button still left Alt+F4
                                      // and the titlebar close as a way to bypass this.
            TopMost = true;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            Padding = new Padding(0, 0, 20, 20);

            var lbl = new Label
            {
                Text = $"\"{documentName}\" is classified {levelDisplayName}.\nEnter a password required to open this file:",
                AutoSize = true,
                MaximumSize = new Size(320, 0),
                Location = new Point(20, 15)
            };
            Controls.Add(lbl);

            int y = lbl.Bottom + 15;
            var lblPw = new Label { Text = "Password:", AutoSize = true, Location = new Point(20, y + 4) };
            _txtPassword.Location = new Point(110, y);
            Controls.Add(lblPw);
            Controls.Add(_txtPassword);

            y = _txtPassword.Bottom + 10;
            var lblConfirm = new Label { Text = "Confirm:", AutoSize = true, Location = new Point(20, y + 4) };
            _txtConfirm.Location = new Point(110, y);
            Controls.Add(lblConfirm);
            Controls.Add(_txtConfirm);

            y = _txtConfirm.Bottom + 8;
            _lblError.Location = new Point(20, y);
            _lblError.MaximumSize = new Size(310, 0);
            Controls.Add(_lblError);

            y = _lblError.Bottom + 12;
            _btnOk.Location = new Point(20, y);
            _btnOk.Click += (_, _) => Validate();
            Controls.Add(_btnOk);

            _btnCancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, AutoSize = true };
            _btnCancel.Location = new Point(_btnOk.Right + 10, y);
            _btnCancel.Enabled = allowCancel;
            _btnCancel.Visible = allowCancel;
            Controls.Add(_btnCancel);

            AcceptButton = _btnOk;
            if (allowCancel) CancelButton = _btnCancel;

            // Same reasoning as ClassificationPromptForm: ControlBox = false removes the X
            // button, but doesn't stop Alt+F4 or a taskbar "Close" from force-closing the
            // window. When mandatory, block any close that didn't come from Validate()
            // succeeding, so EnteredPassword can't end up null here either.
            FormClosing += (_, e) =>
            {
                if (!allowCancel && DialogResult != DialogResult.OK)
                    e.Cancel = true;
            };
        }

        private void Validate()
        {
            if (string.IsNullOrEmpty(_txtPassword.Text))
            {
                ShowError("Password can't be empty.");
                return;
            }
            if (_txtPassword.Text != _txtConfirm.Text)
            {
                ShowError("Passwords don't match.");
                return;
            }

            EnteredPassword = _txtPassword.Text;
            DialogResult = DialogResult.OK;
            Close();
        }

        private void ShowError(string message)
        {
            _lblError.Text = message;
            _lblError.Visible = true;
        }
    }
}