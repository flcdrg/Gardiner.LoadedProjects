using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Xml.Serialization;
using DavidGardiner.Gardiner_LoadedProjects.Annotations;

namespace DavidGardiner.Gardiner_LoadedProjects
{
    [Serializable]
    public class Profile : INotifyPropertyChanged
    {
        private string _name;
        private ObservableCollection<string> _unloadedProjects;

        public Profile()
        {
            UnloadedProjects = new ObservableCollection<string>();
            UnloadedProjects.CollectionChanged += ( sender, args ) => OnPropertyChanged( "UnloadedProjects" );
        }

        /// <summary>
        /// Array of project paths, relative to solution
        /// </summary>
        [XmlArrayItem("Project")]
        public ObservableCollection<string> UnloadedProjects
        {
            get { return _unloadedProjects; }
            private set
            {
                if ( Equals( value, _unloadedProjects ) )
                    return;
                _unloadedProjects = value;
                OnPropertyChanged();
            }
        }

        [XmlAttribute("Name")]
        public string Name
        {
            get { return _name; }
            set
            {
                if ( value == _name )
                    return;
                _name = value;
                OnPropertyChanged();
            }
        }

        public override string ToString()
        {
            return Name;
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