using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace Gardiner.LoadedProjects
{
    public partial class frmProfiles : Form
    {
        public frmProfiles()
        {
            InitializeComponent();
            SelectedProfile = null;
        }

        public Settings Settings { get; set; }

        public IList<string> Unloaded { get; set; }
        public Profile SelectedProfile { get; private set; }

        private void btnCancel_Click( object sender, EventArgs e )
        {
            DialogResult = DialogResult.Cancel;

            Close();
        }

        private void frmProfiles_Load( object sender, EventArgs e )
        {
            lstProfiles.BeginUpdate();
            lstProfiles.Items.Clear();
            object[] profiles = Settings.Profiles.ToArray();
            lstProfiles.Items.AddRange( profiles );
            lstProfiles.EndUpdate();
        }

        private void btnSave_Click( object sender, EventArgs e )
        {
            // save current loaded projects

            var profile = new Profile();
            profile.UnloadedProjects.Clear();
            foreach ( var proj in Unloaded )
            {
                profile.UnloadedProjects.Add( proj );
            }

            bool saveProfile = false;
            using ( var frm = new frmProfileName {ProfileName = DateTime.Now.ToString()} )
            {
                if (frm.ShowDialog(this) == DialogResult.OK)
                    saveProfile = true;

                profile.Name = frm.ProfileName;
            }

            if (saveProfile)
                {
                    Settings.Profiles.Add( profile );

                    DialogResult = DialogResult.OK;
                    Close();
                }
        }

        private void btnDelete_Click( object sender, EventArgs e )
        {
            if ( lstProfiles.SelectedItem != null &&
                 MessageBox.Show( "Are you sure?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation ) ==
                 DialogResult.Yes )
            {
                Settings.Profiles.Remove( (Profile) lstProfiles.SelectedItem );
            }
        }


        private void btnLoad_Click( object sender, EventArgs e )
        {
            SelectedProfile = (Profile) lstProfiles.SelectedItem;
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}