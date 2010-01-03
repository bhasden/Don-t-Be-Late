using System;
using System.Reflection;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace RiotousLabs.DontBeLate
{
    // Derived from http://blogs.msdn.com/coding4fun/archive/2006/10/31/908472.aspx
    public partial class ThisAddIn
    {
        #region Private Properties
        private Office.CommandBarButton ToolbarButton
        {
            get;
            set;
        }
        #endregion

        #region Application Events
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Toolbar button code derived from http://blogs.msdn.com/dancre/archive/2004/03/21/93712.aspx
            var l_CommandBars = this.Application.ActiveExplorer().CommandBars;

            // Create button
            this.ToolbarButton = (Office.CommandBarButton)l_CommandBars["Standard"].Controls.Add(1, Missing.Value, Missing.Value, Missing.Value, true);
            this.ToolbarButton.Caption = "DontBeLate";
            this.ToolbarButton.Style = Office.MsoButtonStyle.msoButtonIcon;
            this.ToolbarButton.Picture = ConvertImage.Convert(Properties.Resources.Icon);
            this.ToolbarButton.Tag = "Don't Be Late Settings";
            this.ToolbarButton.Visible = true;

            this.ToolbarButton.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(Settings_Click);
            this.Application.Reminder += new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_ReminderEventHandler(Reminder);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
        #endregion

        private void Settings_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            var l_Form = new SettingsDialog();
            l_Form.EmailTextBox.Text = Properties.Settings.Default.Email;
            l_Form.EnabledCheckBox.Checked = Properties.Settings.Default.Enabled;

            if (l_Form.ShowDialog() == DialogResult.OK)
            {
                Properties.Settings.Default.Email = l_Form.EmailTextBox.Text;
                Properties.Settings.Default.Enabled = l_Form.EnabledCheckBox.Checked;

                Properties.Settings.Default.Save();
            }
        }

        private void Reminder(object Item)
        {
            if (Properties.Settings.Default.Enabled)
            {
                string l_Message;

                if (Item is Outlook.MailItem)
                {
                    var l_Mail = (Outlook.MailItem)Item;
                    l_Message = String.Format("Mail subject: {0}\nReminder time: {1:MMM d, yyyy @ h:mm tt}", l_Mail.Subject, l_Mail.ReminderTime);
                }
                else if (Item is Outlook.AppointmentItem)
                {
                    var l_Appt = (Outlook.AppointmentItem)Item;

                    if (string.IsNullOrEmpty(l_Appt.Location))
                    {
                        l_Message = String.Format("Appointment subject: {0}\nStart time: {1:MMM d, yyyy @ h:mm tt}", l_Appt.Subject, l_Appt.Start);
                    }
                    else
                    {
                        l_Message = String.Format("Appointment subject: {0}\nLocation: {1}\nStart time: {2:MMM d, yyyy @ h:mm tt}", l_Appt.Subject, l_Appt.Location, l_Appt.Start);
                    }
                }
                else if (Item is Outlook.TaskItem)
                {
                    var l_Task = (Outlook.TaskItem)Item;
                    l_Message = String.Format("Task subject: {0}\nReminder time: {1:MMM d, yyyy @ h:mm tt}", l_Task.Subject, l_Task.ReminderTime);
                }
                else
                {
                    // Unsupported item
                    return;
                }

                // Mail sending derived from http://support.microsoft.com/kb/310263
                try
                {
                    var l_ReminderMail = (Outlook.MailItem)this.Application.CreateItem(Outlook.OlItemType.olMailItem);

                    l_ReminderMail.Recipients.Add(Properties.Settings.Default.Email);
                    l_ReminderMail.Subject = "Don't Be Late";
                    l_ReminderMail.Body = l_Message;

                    // Can't seem to find a way to remove the ambiguity warning
                    l_ReminderMail.Send();
                }
                catch (Exception Ex)
                {
                    MessageBox.Show(string.Format("Could not send Don't Be Late email.\n\nReason:\n{0}", Ex));
                }
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
