﻿using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WpfApp1.Models;

namespace WpfApp1.ViewModels
{
    public class PersonData : INotifyPropertyChanged
    {

        private PersonDataModel _person;

        public PersonDataModel Person
        {
            get { return _person; }
            set
            {
                if (_person != value)
                {
                    _person = value;
                    OnPropertyChanged(nameof(Person));
                }
            }
        }

        public ICommand StartCommand { get; }

        private ExcelTemplate excelTemplate;

        public PersonData()
        {
            Person = new PersonDataModel();
            excelTemplate = new ExcelTemplate();
            excelTemplate.ButtonClicked += MyButtonClickHandler;
            StartCommand = new RelayCommand(excelTemplate.OnButtonClick);
        }
        private void MyButtonClickHandler()
        {
            excelTemplate.MyButtonClickMethod(Person);
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

    }
}
