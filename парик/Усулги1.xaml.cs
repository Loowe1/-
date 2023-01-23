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
    /// Логика взаимодействия для Усулги1.xaml
    /// </summary>
    public partial class Усулги1 : Window
    {
        string вид = "select IDusugi as 'Номер', Nazvanie as 'Название', Stoimost as 'Стоимость, руб', Opisanie as 'Описание' from usugi where ";
        string поиск = "IDusugi like '%%' ";
        string сортировка = "order by IDusugi";
        public Усулги1()
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
            Loadus();
        }
        private void Loadus()
        {
            MySqlDataAdapter adapter = new MySqlDataAdapter(вид + поиск + сортировка, DB.connection);
            System.Data.DataTable tables = new System.Data.DataTable();
            adapter.Fill(tables);
            usl.ItemsSource = tables.DefaultView;
        }

        private void Poisk_po_Loaded(object sender, RoutedEventArgs e)
        {
            Poisk_po.Items.Add("Номер");
            Poisk_po.Items.Add("Название");
            Poisk_po.Items.Add("Стоимность, руб");
            Poisk_po.Items.Add("Описание");
        }

        private void Sortirovka_po_Loaded(object sender, RoutedEventArgs e)
        {
            Sortirovka_po.Items.Add("Номер");
            Sortirovka_po.Items.Add("Название");
            Sortirovka_po.Items.Add("Стоимость, руб");
            Sortirovka_po.Items.Add("Описание");
        }

        private void ASCsort_Checked(object sender, RoutedEventArgs e)
        {
            if (Sortirovka_po.SelectedIndex == 0)
            {
                сортировка = "order by IDusugi asc";
            }
            else if (Sortirovka_po.SelectedIndex == 1)
            {
                сортировка = "order by Nazvanie asc";
            }
            else if (Sortirovka_po.SelectedIndex == 2)
            {
                сортировка = "order by Stoimost asc";
            }
            else if (Sortirovka_po.SelectedIndex == 3)
            {
                сортировка = "order by Opisanie asc";
            }
            Loadus();
        }

        private void DESCsort_Checked(object sender, RoutedEventArgs e)
        {
            if (Sortirovka_po.SelectedIndex == 0)
            {
                сортировка = "order by IDusugi desc";
            }
            else if (Sortirovka_po.SelectedIndex == 1)
            {
                сортировка = "order by Nazvanie desc";
            }
            else if (Sortirovka_po.SelectedIndex == 2)
            {
                сортировка = "order by Stoimost desc";
            }
            else if (Sortirovka_po.SelectedIndex == 3)
            {
                сортировка = "order by Opisanie desc";
            }
            Loadus();
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
                поиск = "IDusugi like '%" + Poisk.Text + "%' ";
            }
            else if (Poisk_po.SelectedIndex == 1)
            {
                поиск = "Nazvanie like '%" + Poisk.Text + "%' ";
            }
            else if (Poisk_po.SelectedIndex == 2)
            {
                поиск = "Stoimost like '%" + Poisk.Text + "%' ";
            }
            else if (Poisk_po.SelectedIndex == 3)
            {
                поиск = "Opisanie like '%" + Poisk.Text + "%' ";
            }
            Loadus();
        }

        private void Sortirovka_po_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            сортировка = "order by IDusugi";
            ASCsort.IsChecked = false;
            DESCsort.IsChecked = false;

            Loadus();
        }

        private void Poisk_po_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            поиск = "IDusugi like '%%' ";
            Poisk_po.IsEnabled = true;

            Loadus();
        }
    }
}
