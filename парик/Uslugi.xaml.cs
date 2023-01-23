using Microsoft.Office.Interop.Excel;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
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
using Window = System.Windows.Window;

namespace парик
{
    /// <summary>
    /// Логика взаимодействия для Uslugi.xaml
    /// </summary>
    public partial class Uslugi : Window
    {
        string вид = "select IDusugi as 'Номер', Nazvanie as 'Название', Stoimost as 'Стоимость, руб', Opisanie as 'Описание' from usugi where ";
        string поиск = "IDusugi like '%%' ";
        string сортировка = "order by IDusugi";
        public Uslugi()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult messageBoxResult = MessageBox.Show("Вы уверены, что хотите выйти?", "Предупреждение", MessageBoxButton.YesNo, MessageBoxImage.Warning);

            if (messageBoxResult == MessageBoxResult.Yes)
            {
                admin admin = new admin();
                admin.Show();
                Hide();
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            Loadusl();
        }
        private void Loadusl()
        {
            MySqlDataAdapter adapter = new MySqlDataAdapter(вид + поиск + сортировка, DB.connection);
            System.Data.DataTable tables = new System.Data.DataTable();
            adapter.Fill(tables);
            usl.ItemsSource = tables.DefaultView;
        }
        //Добавление
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            if(nazv.Text == "" || sto.Text == "" || op.Text == "")
            {
                MessageBox.Show("Все поля должны быть заполнены", "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            else
            {
                
                int nazzvan = -1;
                bool nazvanie = true;
                DB.connection.Open();
                MySqlCommand data_login = new MySqlCommand("select IDusugi from usugi where Nazvanie = @Nazvanie", DB.connection);
                data_login.Parameters.AddWithValue("Nazvanie", nazv.Text);
                MySqlDataReader read_login = data_login.ExecuteReader();

                while (read_login.Read())
                {
                    nazzvan = Convert.ToInt32(read_login["IDusugi"]);
                }

                read_login.Close();

                if (nazzvan != -1)
                {
                    nazvanie = false;

                    MessageBox.Show("Услуга с таким названием уже есть в базе данных!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                DB.connection.Close();
                if(nazvanie == true)
                {
                    DB.connection.Open();

                    MySqlCommand insert = new MySqlCommand("insert into usugi (Nazvanie, Stoimost, Opisanie) values (@Nazvanie, @Stoimost, @Opisanie)", DB.connection);

                    insert.Parameters.AddWithValue("Nazvanie", nazv.Text);
                    insert.Parameters.AddWithValue("Stoimost", sto.Text);
                    insert.Parameters.AddWithValue("Opisanie", op.Text);

                    insert.ExecuteNonQuery();
                    DB.connection.Close();

                    Loadusl();
                }
            }
        }
        
        private void sto_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            if (!Char.IsDigit(e.Text, 0))
            {
                e.Handled = true;
                MessageBox.Show("В этой строке нельзя псиать буквы!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        //Редактирование
        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            if (nazv.Text == "" || sto.Text == "" || op.Text == "")
            {
                MessageBox.Show("Все поля должны быть заполнены", "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            else
            {
                DataRowView row = (DataRowView)usl.SelectedItems[0];
                int nazzvan = -1;
                bool nazvanie = true;
                DB.connection.Open();
                MySqlCommand data_login = new MySqlCommand("select IDusugi from usugi where Nazvanie = @Nazvanie", DB.connection);
                data_login.Parameters.AddWithValue("Nazvanie", nazv.Text);
                MySqlDataReader read_login = data_login.ExecuteReader();

                while (read_login.Read())
                {
                    nazzvan = Convert.ToInt32(read_login["IDusugi"]);
                }

                read_login.Close();

                if (nazzvan != -1 && nazzvan != Convert.ToInt32(row["Номер"]))
                {
                    nazvanie = false;

                    MessageBox.Show("Услуга с таким названием уже есть в базе данных!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                DB.connection.Close();
                if (nazvanie == true)
                {
                    DB.connection.Open();

                    MySqlCommand update = new MySqlCommand("update usugi set Nazvanie = @Nazvanie, Stoimost = @Stoimost, Opisanie = @Opisanie where IDusugi = @IDusugi", DB.connection);
                    update.Parameters.AddWithValue("IDusugi", Convert.ToInt32(row["Номер"]));
                    update.Parameters.AddWithValue("Nazvanie", nazv.Text);
                    update.Parameters.AddWithValue("Stoimost", sto.Text);
                    update.Parameters.AddWithValue("Opisanie", op.Text);

                    update.ExecuteNonQuery();
                    DB.connection.Close();

                    Loadusl();
                }
            }
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            try
            {
                DataRowView row = (DataRowView)usl.SelectedItems[0];


                MessageBoxResult messageBoxResult = MessageBox.Show("Вы уверены, что хотите удалить данные выбранного исполнителя?", "Предупреждение", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                if (messageBoxResult == MessageBoxResult.Yes)
                {
                    try
                    {
                        DB.connection.Open();

                        MySqlCommand delete = new MySqlCommand("delete from usugi where IDusugi = " + Convert.ToInt32(row["Номер"]), DB.connection);
                        delete.ExecuteNonQuery();

                        DB.connection.Close();

                        Loadusl();
                    }
                    catch
                    {
                        DB.connection.Close();

                        MessageBox.Show("Нельзя удалить исполнителя, пока он записан", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
            catch
            {
                MessageBox.Show("Выберите элемент из таблицы и нажмите на кнопку повторно", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }
        }

        private void Poisk_po_Loaded(object sender, RoutedEventArgs e)
        {
            Poisk_po.Items.Add("Номер");
            Poisk_po.Items.Add("Название");
            Poisk_po.Items.Add("Стоимость, руб");
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
            Loadusl();
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
            Loadusl();
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
            Loadusl();
        }
        //сорт вместе
        private void Sortirovka_po_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            сортировка = "order by IDusugi";
            ASCsort.IsChecked = false;
            DESCsort.IsChecked = false;

            Loadusl();
        }
        //поиск вместе
        private void Poisk_po_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            поиск = "IDusugi like '%%' ";
            Poisk_po.IsEnabled = true;

            Loadusl();
        }

        private void Button_Click_5(object sender, RoutedEventArgs e)
        {
            nazv.Clear();
            sto.Clear();
            op.Clear();
            MessageBox.Show("Поля очищены");
        }

        private void usl_PreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                DataRowView row = (DataRowView)usl.SelectedItems[0];

                DB.connection.Open();
                System.Data.DataTable table = new System.Data.DataTable();
                MySqlDataAdapter ad1 = new MySqlDataAdapter($"select Nazvanie, Stoimost, Opisanie from usugi where IDusugi = \"{Convert.ToInt32(row["Номер"])}\"", DB.connection);
                ad1.Fill(table);

                nazv.Text = table.Rows[0][0].ToString();
                sto.Text = table.Rows[0][1].ToString();
                op.Text = table.Rows[0][2].ToString();

                DB.connection.Close();
            }
            catch
            {
                MessageBox.Show("Выберите элемент из таблицы и нажмите на кнопку повторно", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }
        }

        private void Button_Click_6(object sender, RoutedEventArgs e)
        {
            DataRowView row = (DataRowView)usl.SelectedItems[0];

            Microsoft.Office.Interop.Word.Document doc = null;
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            string source = @"C:\Users\shaki\Desktop\My Fucking Works\парик\вывод с парика\УслугиАдмин\Барбершоп.docx";
            doc = app.Documents.Open(source); //открываем документ

            Microsoft.Office.Interop.Word.Bookmarks wBoookmarks = doc.Bookmarks; //Закладки
            Microsoft.Office.Interop.Word.Range wRange;

            int i = 0;

            string id = row["Номер"].ToString();
            string nazv = row["Название"].ToString();
            string stoa = row["Стоимость, руб"].ToString();
            string opa = row["Описание"].ToString();

            string[] data = new string[4] { id, nazv, stoa, opa };

            foreach (Microsoft.Office.Interop.Word.Bookmark mark in wBoookmarks)
            {
                wRange = mark.Range;
                wRange.Text = data[i];
                i++;
            }

            doc = null;
            Process.Start(@"C:\Users\shaki\Desktop\My Fucking Works\парик\вывод с парика\УслугиАдмин\Барбершоп.docx");
        }

        private void Button_Click_7(object sender, RoutedEventArgs e)
        {
            DataRowView row = (DataRowView)usl.SelectedItems[0];

            var Excel = new Microsoft.Office.Interop.Excel.Application();
            var xlWB = Excel.Workbooks.Open(@"C:\Users\shaki\Desktop\My Fucking Works\парик\вывод с парика\УслугиАдмин\Услуги.xlsx");
            var xlSht = xlWB.Worksheets[1];

            xlSht.Cells[18, 1] = row["Номер"];
            xlSht.Cells[18, 2] = row["Название"];
            xlSht.Cells[18, 4] = row["Стоимость, руб"];
            xlSht.Cells[18, 5] = row["Описание"];
            

            Excel.Visible = true;
        }
    }
}
