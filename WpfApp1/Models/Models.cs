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


        public string NamePerson
        {
            get { return _namePerson; }
            set
            {
                if (_namePerson != value)
                {
                    _namePerson = value;
                    OnPropertyChanged(nameof(NamePerson));
                }
            }
        }

        public string AddressPerson
        {
            get { return _addressPerson; }
            set
            {
                if (_addressPerson != value)
                {
                    _addressPerson = value;
                    OnPropertyChanged(nameof(AddressPerson));
                }
            }
        }

        public string CreditPerson
        {
            get { return _creditPerson; }
            set
            {
                if (_creditPerson != value)
                {
                    _creditPerson = value;
                    OnPropertyChanged(nameof(CreditPerson));
                }
            }
        }

        public DateTime? DatePerson
        {
            get { return _datePerson; }
            set
            {
                if (_datePerson != value)
                {
                    _datePerson = value;
                    OnPropertyChanged(nameof(DatePerson));
                }
            }
        }

        public string FinishPerson
        {
            get { return _finishPerson; }
            set
            {
                if (_finishPerson != value)
                {
                    _finishPerson = value;
                    OnPropertyChanged(nameof(FinishPerson));
                }
            }
        }

        public string StartPerson
        {
            get { return _startPerson; }
            set
            {
                if (_startPerson != value)
                {
                    _startPerson = value;
                    OnPropertyChanged(nameof(StartPerson));
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
