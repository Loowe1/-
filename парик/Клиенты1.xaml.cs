using MySql.Data.MySqlClient;
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
    /// Логика взаимодействия для Клиенты1.xaml
    /// </summary>
    public partial class Клиенты1 : Window
    {
        string вид = "select idKlienta as 'Номер', FIO as 'ФИО', Telefon as 'Телефон', kolvo as 'Количесвто посещений' from klients where ";
        string поиск = "idKlienta like '%%' ";
        string сортировка = "order by idKlienta";
        public Клиенты1()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult messageBoxResult = MessageBox.Show("Вы уверены, что хотите выйти?", "Предупреждение", MessageBoxButton.YesNo, MessageBoxImage.Warning);

            if (messageBoxResult == MessageBoxResult.Yes)
            {
                Барбер барбер = new Барбер();
                барбер.Show();
                Hide();
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            Loadkkl();
        }
        private void Loadkkl()
        {
            MySqlDataAdapter adapter = new MySqlDataAdapter(вид + поиск + сортировка, DB.connection);
            System.Data.DataTable tables = new System.Data.DataTable();
            adapter.Fill(tables);
            kli.ItemsSource = tables.DefaultView;
        }

        private void Poisk_po_Loaded(object sender, RoutedEventArgs e)
        {
            Poisk_po.Items.Add("Номер");
            Poisk_po.Items.Add("ФИО");
            Poisk_po.Items.Add("Телефон");
            Poisk_po.Items.Add("Количесвто посещений");
        }

        private void Sortirovka_po_Loaded(object sender, RoutedEventArgs e)
        {
            Sortirovka_po.Items.Add("Номер");
            Sortirovka_po.Items.Add("ФИО");
            Sortirovka_po.Items.Add("Телефон");
            Sortirovka_po.Items.Add("Количесвто посещений");
        }

        private void ASCsort_Checked(object sender, RoutedEventArgs e)
        {
            if (Sortirovka_po.SelectedIndex == 0)
            {
                сортировка = "order by idKlienta asc";
            }
            else if (Sortirovka_po.SelectedIndex == 1)
            {
                сортировка = "order by FIO asc";
            }
            else if (Sortirovka_po.SelectedIndex == 2)
            {
                сортировка = "order by Telefon asc";
            }
            else if (Sortirovka_po.SelectedIndex == 3)
            {
                сортировка = "order by kolvo asc";
            }
            Loadkkl();
        }

        private void DESCsort_Checked(object sender, RoutedEventArgs e)
        {
            if (Sortirovka_po.SelectedIndex == 0)
            {
                сортировка = "order by idKlienta desc";
            }
            else if (Sortirovka_po.SelectedIndex == 1)
            {
                сортировка = "order by FIO desc";
            }
            else if (Sortirovka_po.SelectedIndex == 2)
            {
                сортировка = "order by Telefon desc";
            }
            else if (Sortirovka_po.SelectedIndex == 3)
            {
                сортировка = "order by kolvo desc";
            }
            Loadkkl();
        }

        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            Poisk_po.SelectedIndex = -1;
            Sortirovka_po.SelectedIndex = -1;
            Poisk.Clear();
        }

        private void Poisk_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (Poisk_po.SelectedIndex == 0)
            {
                поиск = "idKlienta like '%" + Poisk.Text + "%' ";
            }
            else if (Poisk_po.SelectedIndex == 1)
            {
                поиск = "FIO like '%" + Poisk.Text + "%' ";
            }
            else if (Poisk_po.SelectedIndex == 2)
            {
                поиск = "Telefon like '%" + Poisk.Text + "%' ";
            }
            else if (Poisk_po.SelectedIndex == 3)
            {
                поиск = "kolvo like '%" + Poisk.Text + "%' ";
            }
            Loadkkl();
        }

        private void Sortirovka_po_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            сортировка = "order by idKlienta";
            ASCsort.IsChecked = false;
            DESCsort.IsChecked = false;

            Loadkkl();
        }

        private void Poisk_po_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            поиск = "idKlienta like '%%' ";
            Poisk_po.IsEnabled = true;

            Loadkkl();
        }
    }
}
