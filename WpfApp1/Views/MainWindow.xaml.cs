using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using WpfApp1.ViewModels;

namespace WpfApp1
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void ButtonExl(object sender, EventArgs e)
        {
            MainViewModels mainViewModels = new MainViewModels();
            mainViewModels.CreateExcelTemplate(this.namePerson.Text, this.addressPerson.Text, this.startPerson.Text, this.finishPerson.Text, this.creditPerson.Text, this.datePerson);
            MessageBox.Show("Шаблон Excel успешно создан!");
        }
    }

}

