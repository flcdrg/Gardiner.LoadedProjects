using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Xml.Serialization;
using DavidGardiner.Gardiner_LoadedProjects.Annotations;

namespace DavidGardiner.Gardiner_LoadedProjects
{
    [Serializable]
    [XmlRoot("Settings")]
    public class Settings : INotifyPropertyChanged
    {
        private ObservableCollection<Profile> _profiles;
        public ObservableCollection<Profile> Profiles
        {
            get { return _profiles; }
            private set
            {
                if ( Equals( value, _profiles ) )
                    return;
                _profiles = value;
                OnPropertyChanged();
            }
        }

        public Settings()
        {
            Profiles = new ObservableCollection<Profile>();

            Profiles.CollectionChanged += ( sender, args ) => OnPropertyChanged( "Profiles" );
        }

        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged( [CallerMemberName] string propertyName = null )
        {
            var handler = PropertyChanged;
            if ( handler != null )
                handler( this, new PropertyChangedEventArgs( propertyName ) );
        }
    }
}