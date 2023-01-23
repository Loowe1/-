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
    /// Логика взаимодействия для Барбер.xaml
    /// </summary>
    public partial class Барбер : Window
    {
        public Барбер()
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
            Клиенты1 клиенты1 = new Клиенты1();
            клиенты1.Show();
            Hide();
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            Усулги1 усулги1 = new Усулги1();
            усулги1.Show();
            Hide();
        }
    }
}
