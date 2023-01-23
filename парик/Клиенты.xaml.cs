using Microsoft.Office.Interop.Excel;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Security.Policy;
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
    /// Логика взаимодействия для Клиенты.xaml
    /// </summary>
    public partial class Клиенты : Window
    {
        string вид = "select idKlienta as 'Номер', FIO as 'ФИО', Telefon as 'Телефон', kolvo as 'Количество посещений' from klients where ";
        string поиск = "idKlienta like '%%' ";
        string сортировка = "order by idKlienta";
        public Клиенты()
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
            Loadkl();
        }
        private void Loadkl()
        {
            MySqlDataAdapter adapter = new MySqlDataAdapter(вид + поиск + сортировка, DB.connection);
            System.Data.DataTable tables = new System.Data.DataTable();
            adapter.Fill(tables);
            kli.ItemsSource = tables.DefaultView;
        }

        private void sto_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            if (!Char.IsDigit(e.Text, 0))
            {
                e.Handled = true;
                MessageBox.Show("В этой строке нельзя псиать буквы!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            if(fio1.Text == "" || tel1.Text == "")
            {
                MessageBox.Show("Все поля должны быть заполнены", "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            else
            {
                if (tel1.Text.Length < 11)
                {
                    MessageBox.Show("Номер телефона введён неверно", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                else
                {
                    
                    int id_tel = -1;
                   
                    bool tel = true;

                    DB.connection.Open();
                    MySqlCommand data_login1 = new MySqlCommand("select idKlienta from klients where Telefon = @Telefon", DB.connection);
                    data_login1.Parameters.AddWithValue("Telefon", tel1.Text);
                    MySqlDataReader read_login1 = data_login1.ExecuteReader();

                    while (read_login1.Read())
                    {
                        id_tel = Convert.ToInt32(read_login1["idKlienta"]);
                    }

                    read_login1.Close();

                    if (id_tel != -1)
                    {
                        tel = false;

                        MessageBox.Show("Такой номер телефона уже есть в базе данных", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    DB.connection.Close();
                    
                    if(tel == true)
                    {
                        DB.connection.Open();

                        MySqlCommand insert = new MySqlCommand("insert into klients (FIO, Telefon, kolvo) values (@FIO, @Telefon, @kolvo)", DB.connection);

                        insert.Parameters.AddWithValue("FIO", fio1.Text);
                        insert.Parameters.AddWithValue("Telefon", tel1.Text);
                        insert.Parameters.AddWithValue("kolvo", kool.Text);

                        insert.ExecuteNonQuery();
                        DB.connection.Close();

                        Loadkl();
                    }
                }
            }
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            if (fio1.Text == "" || tel1.Text == "")
            {
                MessageBox.Show("Все поля должны быть заполнены", "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            else
            {
                if (tel1.Text.Length < 11)
                {
                    MessageBox.Show("Номер телефона введён неверно", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                else
                {
                    DataRowView row = (DataRowView)kli.SelectedItems[0];
                    
                    int id_tel = -1;
                    bool tel = true;
                    DB.connection.Open();
                    MySqlCommand data_login1 = new MySqlCommand("select idKlienta from klients where Telefon = @Telefon", DB.connection);
                    data_login1.Parameters.AddWithValue("Telefon", tel1.Text);
                    MySqlDataReader read_login1 = data_login1.ExecuteReader();

                    while (read_login1.Read())
                    {
                        id_tel = Convert.ToInt32(read_login1["idKlienta"]);
                    }

                    read_login1.Close();

                    if (id_tel != -1 && id_tel != Convert.ToInt32(row["Номер"]))
                    {
                        tel = false;

                        MessageBox.Show("Такой номер телефона уже есть в базе данных", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    DB.connection.Close();

                    if (tel == true)
                    {
                        DB.connection.Open();

                        MySqlCommand update = new MySqlCommand("update klients set FIO = @FIO, Telefon = @Telefon, kolvo = @kolvo where idKlienta = @idKlienta", DB.connection);

                        update.Parameters.AddWithValue("idKlienta", Convert.ToInt32(row["Номер"]));
                        update.Parameters.AddWithValue("FIO", fio1.Text);
                        update.Parameters.AddWithValue("Telefon", tel1.Text);
                        update.Parameters.AddWithValue("kolvo", kool.Text);

                        update.ExecuteNonQuery();
                        DB.connection.Close();

                        Loadkl();
                    }
                }
            }
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            try
            {
                DataRowView row = (DataRowView)kli.SelectedItems[0];
                MessageBoxResult messageBoxResult = MessageBox.Show("Вы уверены, что хотите удалить данные выбранного исполнителя?", "Предупреждение", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                if (messageBoxResult == MessageBoxResult.Yes)
                {
                    try
                    {
                        DB.connection.Open();

                        MySqlCommand delete = new MySqlCommand("delete from klients where idKlienta = " + Convert.ToInt32(row["Номер"]), DB.connection);
                        delete.ExecuteNonQuery();

                        DB.connection.Close();

                        Loadkl();
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
            Poisk_po.Items.Add("ФИО");
            Poisk_po.Items.Add("Телефон");
            Poisk_po.Items.Add("Количество посещений");
        }

        private void Sortirovka_po_Loaded(object sender, RoutedEventArgs e)
        {
            Sortirovka_po.Items.Add("Номер");
            Sortirovka_po.Items.Add("ФИО");
            Sortirovka_po.Items.Add("Телефон");
            Sortirovka_po.Items.Add("Количество посещений");
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

            Loadkl();
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
            Loadkl();
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
            Loadkl();
        }

        private void Sortirovka_po_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            сортировка = "order by idKlienta";
            ASCsort.IsChecked = false;
            DESCsort.IsChecked = false;

            Loadkl();
        }

        private void Poisk_po_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            поиск = "idKlienta like '%%' ";
            Poisk_po.IsEnabled = true;

            Loadkl();
        }

        private void Button_Click_5(object sender, RoutedEventArgs e)
        {
            fio1.Clear();
            tel1.Clear();
            kool.Clear();
            MessageBox.Show("Поля очищены");
        }

        private void usl_PreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                DataRowView row = (DataRowView)kli.SelectedItems[0];

                DB.connection.Open();
                System.Data.DataTable table = new System.Data.DataTable();
                MySqlDataAdapter ad1 = new MySqlDataAdapter($"select FIO, Telefon from klients where idKlienta = \"{Convert.ToInt32(row["Номер"])}\"", DB.connection);
                ad1.Fill(table);

                fio1.Text = table.Rows[0][0].ToString();
                tel1.Text = table.Rows[0][1].ToString();
                kool.Text = table.Rows[0][2].ToString();
                DB.connection.Close();
            }
            catch
            {
                MessageBox.Show("Выберите элемент из таблицы и нажмите на кнопку повторно", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }
        }

        private void Button_Click_6(object sender, RoutedEventArgs e)
        {
            DataRowView row = (DataRowView)kli.SelectedItems[0];

            Microsoft.Office.Interop.Word.Document doc = null;
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            string source = @"C:\Users\shaki\Desktop\My Fucking Works\парик\вывод с парика\клиентыАдмин\Клиенты вывод.docx";
            doc = app.Documents.Open(source); //открываем документ

            Microsoft.Office.Interop.Word.Bookmarks wBoookmarks = doc.Bookmarks; //Закладки
            Microsoft.Office.Interop.Word.Range wRange;

            int i = 0;

            string id = row["Номер"].ToString();
            string FIO = row["ФИО"].ToString();
            string tel = row["Телефон"].ToString();
            
            string[] data = new string[3] { id, FIO, tel};

            foreach (Microsoft.Office.Interop.Word.Bookmark mark in wBoookmarks)
            {
                wRange = mark.Range;
                wRange.Text = data[i];
                i++;
            }

            doc = null;
            Process.Start(@"C:\Users\shaki\Desktop\My Fucking Works\парик\вывод с парика\клиентыАдмин\Клиенты вывод.docx");
        }

        private void Button_Click_7(object sender, RoutedEventArgs e)
        {
            DataRowView row = (DataRowView)kli.SelectedItems[0];

            var Excel = new Microsoft.Office.Interop.Excel.Application();
            var xlWB = Excel.Workbooks.Open(@"C:\Users\shaki\Desktop\My Fucking Works\парик\вывод с парика\клиентыАдмин\Клиенты.xlsx");
            var xlSht = xlWB.Worksheets[1];

            xlSht.Cells[17, 2] = row["Номер"];
            xlSht.Cells[19, 2] = row["ФИО"];
            xlSht.Cells[21, 2] = row["Телефон"];

            Excel.Visible = true;
        }

        private void fio1_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            char inp = e.Text[0];
            if ((inp < 'А' || inp > 'Я') & (inp < 'а' || inp > 'я'))
            {
                e.Handled = true;
                MessageBox.Show("В этом поле нельзя писать цифры!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}