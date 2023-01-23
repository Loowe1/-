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
using System.Windows.Shapes;

namespace парик
{
    /// <summary>
    /// Логика взаимодействия для admin.xaml
    /// </summary>
    public partial class admin : Window
    {
        public admin()
        {
            InitializeComponent();
            DB db = new DB();
            db.openConnetion();
            admin1.Content = ID.FIO;
            nameAdmin1.Content = "Вы вошли как: " + ID.dolzhnost;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult messageBoxResult = MessageBox.Show("Вы уверены, что хотите выйти?", "Предупреждение", MessageBoxButton.YesNo, MessageBoxImage.Warning);

            if (messageBoxResult == MessageBoxResult.Yes)
            {
                MainWindow mainWindow = new MainWindow();
                mainWindow.Show();
                Hide();
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            Сотрудники сотрудники = new Сотрудники();
            сотрудники.Show();
            Hide();
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            Uslugi uslugi = new Uslugi();
            uslugi.Show();
            Hide();
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            Клиенты клиенты = new Клиенты();
            клиенты.Show();
            Hide();
        }

        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            Чек чек = new Чек();
            чек.Show();
            Hide();
        }
    }
}
