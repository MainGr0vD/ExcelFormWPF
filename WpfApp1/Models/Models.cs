using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Drawing;
using System.Diagnostics;
using System.ComponentModel;

namespace WpfApp1.Models
{
    public class PersonDataModel : INotifyPropertyChanged
    {
        private string _namePerson;
        private string _addressPerson;
        private string _creditPerson;
        private DateTime? _datePerson;
        private string _startPerson;
        private string _finishPerson;
        private string _stringPerson;

        public string stringPerson
        {
            get { return _stringPerson; }
            set
            {
                if (_stringPerson != value)
                {
                    _stringPerson = value;
                    OnPropertyChanged(nameof(stringPerson));
                }
            }
        }

        public string namePerson
        {
            get { return _namePerson; }
            set
            {
                if (_namePerson != value)
                {
                    _namePerson = value;
                    OnPropertyChanged(nameof(namePerson));
                }
            }
        }

        public string addressPerson
        {
            get { return _addressPerson; }
            set
            {
                if (_addressPerson != value)
                {
                    _addressPerson = value;
                    OnPropertyChanged(nameof(addressPerson));
                }
            }
        }

        public string creditPerson
        {
            get { return _creditPerson; }
            set
            {
                if (_creditPerson != value)
                {
                    _creditPerson = value;
                    OnPropertyChanged(nameof(creditPerson));
                }
            }
        }

        public DateTime? datePerson
        {
            get { return _datePerson; }
            set
            {
                if (_datePerson != value)
                {
                    _datePerson = value;
                    OnPropertyChanged(nameof(datePerson));
                }
            }
        }

        public string finishPerson
        {
            get { return _finishPerson; }
            set
            {
                if (_finishPerson != value)
                {
                    _finishPerson = value;
                    OnPropertyChanged(nameof(finishPerson));
                }
            }
        }

        public string startPerson
        {
            get { return _startPerson; }
            set
            {
                if (_startPerson != value)
                {
                    _startPerson = value;
                    OnPropertyChanged(nameof(startPerson));
                }
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
