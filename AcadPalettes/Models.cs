using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AcadPalettes
{
    public class ViewModelBase : System.ComponentModel.INotifyPropertyChanged
    {
        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(propertyName));
        }
        public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;
    }

    public class DWG : ViewModelBase
    {
        private string _Path;
        public string Path
        {
            get { return this._Path; }
            set
            {
                if (this._Path != value)
                {
                    this._Path = value;
                    this.OnPropertyChanged("Path");
                }
            }
        }

        private bool _Check;
        public bool Check
        {
            get { return this._Check; }
            set
            {
                if (this._Check != value)
                {
                    this._Check = value;
                    this.OnPropertyChanged("Check");
                }
            }
        }
    }
}
