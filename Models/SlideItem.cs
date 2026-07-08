using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace TocBuilder_dotnet_framework.Models
{
    public class SlideItem : INotifyPropertyChanged 
    {
        private bool _isSelected = true; 
        private int _number; 
        private byte[] _thumbnail;
        private bool _isBackground;

        public bool IsSelected
        {
            get => _isSelected;
            set { _isSelected = value; OnPropertyChanged(); }
        }

        public int Number { get => _number; set { _number = value; OnPropertyChanged(); } }

        public byte[] Thumbnail 
        { 
            get => _thumbnail; 
            set { _thumbnail = value; OnPropertyChanged(); } 
        }

        public bool IsBackground
        {
            get => _isBackground;
            set
            {
                if (_isBackground != value)
                {
                    _isBackground = value;
                    OnPropertyChanged();
                }
            }
        }

        public event PropertyChangedEventHandler PropertyChanged; 

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null) 
        { 
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName)); 
        }
    }
}
