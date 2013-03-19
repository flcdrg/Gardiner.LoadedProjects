﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace DavidGardiner.Gardiner_LoadedProjects
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
            lstProfiles.Items.Clear();
            lstProfiles.BeginUpdate();
            lstProfiles.Items.AddRange( Settings.Profiles.ToArray() );
            lstProfiles.EndUpdate();
        }

        private void btnSave_Click( object sender, EventArgs e )
        {
            // save current loaded projects

            var profile = new Profile();
            profile.UnloadedProjects.AddRange( Unloaded );
            profile.Name = DateTime.Now.ToString();
            Settings.Profiles.Add( profile );
            DialogResult = DialogResult.OK;
            Close();
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
            this.Close();
        }
    }
}