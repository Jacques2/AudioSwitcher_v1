using System;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Windows.Forms;
using FortyOne.AudioSwitcher.Helpers;
using FortyOne.AudioSwitcher.Properties;

namespace FortyOne.AudioSwitcher
{
    public sealed partial class ExceptionDisplayForm : Form
    {
        private readonly Exception exception;

        public ExceptionDisplayForm()
        {
            InitializeComponent();
        }

        public ExceptionDisplayForm(string title, Exception ex)
        {
            InitializeComponent();

            exception = ex;
            Text = title;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Close();
        }

        public bool IsUserAdministrator()
        {
            //bool value to hold our return value
            bool isAdmin;
            try
            {
                //get the currently logged in user
                var user = WindowsIdentity.GetCurrent();
                var principal = new WindowsPrincipal(user);
                isAdmin = principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
            catch (UnauthorizedAccessException)
            {
                isAdmin = false;
            }
            catch (Exception)
            {
                isAdmin = false;
            }
            return isAdmin;
        }

        private void btnReport_Click(object sender, EventArgs e)
        {

        }

        private void ExceptionDisplayForm_Load(object sender, EventArgs e)
        {
            if (exception != null)
                txtError.Text = exception.ToString();
        }
    }
}