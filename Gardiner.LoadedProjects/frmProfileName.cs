using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DavidGardiner.Gardiner_LoadedProjects
{
    public partial class frmProfileName : Form
    {
        public frmProfileName()
        {
            InitializeComponent();
        }

        public string ProfileName { get; set; }


        private void btnOK_Click( object sender, EventArgs e )
        {
            ProfileName = txtName.Text;

            DialogResult = DialogResult.OK;

            Close();
        }

        private void btnCancel_Click( object sender, EventArgs e )
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private void frmProfileName_Load( object sender, EventArgs e )
        {
            txtName.Text = ProfileName;
        }
    }
}
