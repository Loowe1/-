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
using MySql.Data.MySqlClient;
using System.Data;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using System.Diagnostics;
using Window = System.Windows.Window;

namespace парик
{
    /// <summary>
    /// Логика взаимодействия для Сотрудники.xaml
    /// </summary>
    public partial class Сотрудники : Window
    {
        string вид = "select idSotrudnika as 'Номер', FIO as 'ФИО сотрудника', Pasport as 'Паспорт', Dataro as 'Дата рождения', Adres as 'Адрес', dolzhnost as 'Должность', Telefon as 'Телефон', login as 'Логин', parol as 'Пароль' from sotrudniki where ";
        string поиск = "idSotrudnika like '%%' ";
        string сортировка = "order by idSotrudnika";
        public Сотрудники()
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
            Loadsot();
        }
        private void Loadsot()
        {
            MySqlDataAdapter adapter = new MySqlDataAdapter(вид + поиск + сортировка, DB.connection);
            System.Data.DataTable tables = new System.Data.DataTable();
            adapter.Fill(tables);
            sotrr.ItemsSource = tables.DefaultView;
        }

        private void dol_Loaded(object sender, RoutedEventArgs e)
        {
            dol.Items.Add("Администратор");
            dol.Items.Add("Барбер");
        }

        private void tel_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            if (!Char.IsDigit(e.Text, 0))
            {
                e.Handled = true;
                MessageBox.Show("В этой строке нельзя псиать буквы!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void pasport_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            if (!Char.IsDigit(e.Text, 0))
            {
                e.Handled = true;
                MessageBox.Show("В этой строке нельзя псиать буквы!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void fio_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            char inp = e.Text[0];
            if ((inp < 'А' || inp > 'Я') & (inp < 'а' || inp > 'я'))
            {
                e.Handled = true;
                MessageBox.Show("В этом поле нельзя писать цифры!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        //Добавление
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            if (fio.Text == "" || pasport.Text == "" || dataro.Text == "" || adr.Text == "" || dol.Text == "" || tel.Text == "" || login1.Text == "" || parol.Text == "")
            {
                MessageBox.Show("Все поля должны быть заполнены", "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            else
            {
                if (pasport.Text.Length < 10)
                {
                    MessageBox.Show("Серия и номер паспорта введены неверно", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                else
                {
                    if (tel.Text.Length < 11)
                    {
                        MessageBox.Show("Номер телефона введён неверно", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    else
                    {
                        if (dataro.SelectedDate > DateTime.Today)
                        {
                            MessageBox.Show("Дата рождения введена неверно", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                        else
                        {
                            TimeSpan date = DateTime.Today.Subtract(Convert.ToDateTime(dataro.SelectedDate));
                            double age = date.Days;
                            age = Math.Floor(age / 365);

                            if (age < 18)
                            {
                                MessageBox.Show("Сотруднику должно быть не менее 18-ти лет", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                            }
                            else if (age > 65)
                            {
                                MessageBox.Show("Сотруднику должно быть не более 65-ти лет", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                            }
                            else
                            {
                                int id_passport_emp = -1;
                                int id_passport_own = -1;
                                int id_phone = -1;
                                int id_login = -1;
                                bool passport_emp = true;
                                bool passport_own = true;
                                bool phone = true;
                                bool login = true;

                                DB.connection.Open();

                                MySqlCommand data_passport_emp = new MySqlCommand("select idSotrudnika from sotrudniki where Pasport = @Pasport", DB.connection);
                                data_passport_emp.Parameters.AddWithValue("Pasport", pasport.Text);
                                MySqlDataReader read_passport_emp = data_passport_emp.ExecuteReader();

                                while (read_passport_emp.Read())
                                {
                                    id_passport_emp = Convert.ToInt32(read_passport_emp["idSotrudnika"]);
                                }

                                read_passport_emp.Close();

                                if (id_passport_emp != -1)
                                {
                                    passport_emp = false;

                                    MessageBox.Show("Такие серия и номер паспорта уже есть в базе данных", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                                }

                                MySqlCommand data_phone = new MySqlCommand("select idSotrudnika from sotrudniki where Telefon = @Telefon", DB.connection);
                                data_phone.Parameters.AddWithValue("Telefon", tel.Text);
                                MySqlDataReader read_phone = data_phone.ExecuteReader();

                                while (read_phone.Read())
                                {
                                    id_phone = Convert.ToInt32(read_phone["idSotrudnika"]);
                                }

                                read_phone.Close();

                                if (id_phone != -1)
                                {
                                    phone = false;

                                    MessageBox.Show("Такой номер телефона уже есть в базе данных", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                                }

                                MySqlCommand data_login = new MySqlCommand("select idSotrudnika from sotrudniki where login = @login", DB.connection);
                                data_login.Parameters.AddWithValue("login", login1.Text);
                                MySqlDataReader read_login = data_login.ExecuteReader();

                                while (read_login.Read())
                                {
                                    id_login = Convert.ToInt32(read_login["idSotrudnika"]);
                                }

                                read_login.Close();

                                if (id_login != -1)
                                {
                                    login = false;

                                    MessageBox.Show("Такой логин уже есть в базе данных", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                                }

                                DB.connection.Close();

                                if (passport_emp == true && passport_own == true && phone == true && login == true)
                                {

                                    DB.connection.Open();

                                    MySqlCommand insert = new MySqlCommand("insert into sotrudniki (FIO, Pasport, Dataro, Adres, dolzhnost, Telefon, login, parol) values (@FIO, @Pasport, @Dataro, @Adres, @dolzhnost, @Telefon, @login, @parol)", DB.connection);
                                    insert.Parameters.AddWithValue("FIO", fio.Text);
                                    insert.Parameters.AddWithValue("Pasport", pasport.Text);
                                    insert.Parameters.AddWithValue("Dataro", Convert.ToDateTime(dataro.SelectedDate).ToString("yyyy-MM-dd"));
                                    insert.Parameters.AddWithValue("Adres", adr.Text);
                                    insert.Parameters.AddWithValue("dolzhnost", dol.Text);
                                    insert.Parameters.AddWithValue("Telefon", tel.Text);
                                    insert.Parameters.AddWithValue("login", login1.Text);
                                    insert.Parameters.AddWithValue("parol", parol.Text);

                                    insert.ExecuteNonQuery();
                                    DB.connection.Close();

                                    Loadsot();
                                }
                            }
                        }
                    }
                }
            }
        }
        //Редактирование
        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            if (fio.Text == "" || pasport.Text == "" || dataro.Text == "" || adr.Text == "" || dol.Text == "" || tel.Text == "" || login1.Text == "" || parol.Text == "")
            {
                MessageBox.Show("Все поля должны быть заполнены", "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            else
            {
                if (pasport.Text.Length < 10)
                {
                    MessageBox.Show("Серия и номер паспорта введены неверно", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                else
                {
                    if (tel.Text.Length < 11)
                    {
                        MessageBox.Show("Номер телефона введён неверно", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    else
                    {
                        if (dataro.SelectedDate > DateTime.Today)
                        {
                            MessageBox.Show("Дата рождения введена неверно", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                        else
                        {
                            TimeSpan date = DateTime.Today.Subtract(Convert.ToDateTime(dataro.SelectedDate));
                            double age = date.Days;
                            age = Math.Floor(age / 365);
                            if (age < 18)
                            {
                                MessageBox.Show("Сотруднику должно быть не менее 18-ти лет", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                            }
                            else if (age > 65)
                            {
                                MessageBox.Show("Сотруднику должно быть не более 65-ти лет", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                            }
                            else
                            {
                                DataRowView row = (DataRowView)sotrr.SelectedItems[0];
                                int id_passport_emp = -1;
                                int id_passport_own = -1;
                                int id_phone = -1;
                                int id_login = -1;
                                bool passport_emp = true;
                                bool passport_own = true;
                                bool phone = true;
                                bool login = true;

                                DB.connection.Open();

                                MySqlCommand data_passport_emp = new MySqlCommand("select idSotrudnika from sotrudniki where Pasport = @Pasport", DB.connection);
                                data_passport_emp.Parameters.AddWithValue("Pasport", pasport.Text);
                                MySqlDataReader read_passport_emp = data_passport_emp.ExecuteReader();

                                while (read_passport_emp.Read())
                                {
                                    id_passport_emp = Convert.ToInt32(read_passport_emp["idSotrudnika"]);
                                }

                                read_passport_emp.Close();

                                if (id_passport_emp != -1 && id_passport_emp != Convert.ToInt32(row["Номер"]))
                                {
                                    passport_emp = false;

                                    MessageBox.Show("Такие серия и номер паспорта уже есть в базе данных", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                                }

                                MySqlCommand data_phone = new MySqlCommand("select idSotrudnika from sotrudniki where Telefon = @Telefon", DB.connection);
                                data_phone.Parameters.AddWithValue("Telefon", tel.Text);
                                MySqlDataReader read_phone = data_phone.ExecuteReader();

                                while (read_phone.Read())
                                {
                                    id_phone = Convert.ToInt32(read_phone["idSotrudnika"]);
                                }

                                read_phone.Close();

                                if (id_phone != -1 && id_phone != Convert.ToInt32(row["Номер"]))
                                {
                                    phone = false;

                                    MessageBox.Show("Такой номер телефона уже есть в базе данных", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                                }

                                MySqlCommand data_login = new MySqlCommand("select idSotrudnika from sotrudniki where login = @login", DB.connection);
                                data_login.Parameters.AddWithValue("login", login1.Text);
                                MySqlDataReader read_login = data_login.ExecuteReader();

                                while (read_login.Read())
                                {
                                    id_login = Convert.ToInt32(read_login["idSotrudnika"]);
                                }

                                read_login.Close();

                                if (id_login != -1 && id_login != Convert.ToInt32(row["Номер"]))
                                {
                                    login = false;

                                    MessageBox.Show("Такой логин уже есть в базе данных", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                                }

                                DB.connection.Close();

                                if (passport_emp == true && passport_own == true && phone == true && login == true)
                                {

                                    DB.connection.Open();

                                    MySqlCommand update = new MySqlCommand("update sotrudniki set FIO = @FIO, Pasport = @Pasport, Dataro = @Dataro, Adres = @Adres, dolzhnost = @dolzhnost, Telefon = @Telefon, login = @login, parol = @parol where idSotrudnika = @idSotrudnika", DB.connection);
                                    update.Parameters.AddWithValue("idSotrudnika", Convert.ToInt32(row["Номер"]));
                                    update.Parameters.AddWithValue("FIO", fio.Text);
                                    update.Parameters.AddWithValue("Pasport", pasport.Text);
                                    update.Parameters.AddWithValue("Dataro", Convert.ToDateTime(dataro.SelectedDate).ToString("yyyy-MM-dd"));
                                    update.Parameters.AddWithValue("Adres", adr.Text);
                                    update.Parameters.AddWithValue("dolzhnost", dol.Text);
                                    update.Parameters.AddWithValue("Telefon", tel.Text);
                                    update.Parameters.AddWithValue("login", login1.Text);
                                    update.Parameters.AddWithValue("parol", parol.Text);

                                    update.ExecuteNonQuery();
                                    DB.connection.Close();

                                    Loadsot();
                                }
                            }
                        }
                    }
                }
            }
        }

        //Удаление
        private void Button_Click_3(object sender, RoutedEventArgs e)
        {

            try
            {
                DB.connection.Open();
                DataRowView row = (DataRowView)sotrr.SelectedItems[0];

                int idst = -1;

                int dof = Convert.ToInt32(row["Номер"].ToString());
                MySqlCommand email_emp = new MySqlCommand("select idsotrudnika from sotrudniki where idsotrudnika = " + ID.idSotrudnika, DB.connection);

                MySqlDataReader read_email = email_emp.ExecuteReader();

                while (read_email.Read())
                {
                    idst = Convert.ToInt32(read_email["idsotrudnika"]);
                }
                read_email.Close();
                DB.connection.Close();
                if (ID.idSotrudnika == Convert.ToInt32(row["Номер"]))
                {

                    MessageBox.Show("Вы не можете удалить себя!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                else
                {

                    MessageBoxResult messageBoxResult = MessageBox.Show("Вы уверены, что хотите удалить данные выбранного исполнителя?", "Предупреждение", MessageBoxButton.YesNo, MessageBoxImage.Warning);

                    if (messageBoxResult == MessageBoxResult.Yes)
                    {
                        try
                        {
                            DB.connection.Open();

                            MySqlCommand delete = new MySqlCommand("delete from sotrudniki where idsotrudnika = " + Convert.ToInt32(row["Номер"]), DB.connection);
                            delete.ExecuteNonQuery();

                            DB.connection.Close();

                            Loadsot();
                        }
                        catch
                        {
                            DB.connection.Close();

                            MessageBox.Show("Нельзя удалить исполнителя, пока он записан", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                    }
                }
            }
            catch
            {
                MessageBox.Show("Выберите элемент из таблицы и нажмите на кнопку повторно", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }
        }

        private void sotrr_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                DataRowView row = (DataRowView)sotrr.SelectedItems[0];

                DB.connection.Open();
                System.Data.DataTable table = new System.Data.DataTable();
                MySqlDataAdapter ad1 = new MySqlDataAdapter($"select FIO, Pasport, Dataro, Adres, dolzhnost, Telefon, login, parol from sotrudniki where idSotrudnika = \"{Convert.ToInt32(row["Номер"])}\"", DB.connection);
                ad1.Fill(table);

                fio.Text = table.Rows[0][0].ToString();
                pasport.Text = table.Rows[0][1].ToString();
                dataro.Text = table.Rows[0][2].ToString();
                adr.Text = table.Rows[0][3].ToString();
                dol.Text = table.Rows[0][4].ToString();
                tel.Text = table.Rows[0][5].ToString();
                login1.Text = table.Rows[0][6].ToString();
                parol.Text = table.Rows[0][7].ToString();

                DB.connection.Close();
            }
            catch
            {
                MessageBox.Show("Выберите элемент из таблицы и нажмите на кнопку повторно", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }
        }

        private void ComboBox_Loaded(object sender, RoutedEventArgs e)
        {
            Poisk_po.Items.Add("Номер");
            Poisk_po.Items.Add("ФИО");
            Poisk_po.Items.Add("Паспортные данные");
            Poisk_po.Items.Add("Дата рождения");
            Poisk_po.Items.Add("Адрес");
            Poisk_po.Items.Add("Должность");
            Poisk_po.Items.Add("Телефон");
            Poisk_po.Items.Add("Логин");
            Poisk_po.Items.Add("Пароль");
        }

        private void sort_Loaded(object sender, RoutedEventArgs e)
        {
            Sortirovka_po.Items.Add("Номер");
            Sortirovka_po.Items.Add("ФИО");
            Sortirovka_po.Items.Add("Паспортные данные");
            Sortirovka_po.Items.Add("Дата рождения");
            Sortirovka_po.Items.Add("Адрес");
            Sortirovka_po.Items.Add("Должность");
            Sortirovka_po.Items.Add("Телефон");
            Sortirovka_po.Items.Add("Логин");
            Sortirovka_po.Items.Add("Пароль");
        }

        private void sort_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            сортировка = "order by idSotrudnika";
            ASCsort.IsChecked = false;
            DESCsort.IsChecked = false;

            Loadsot();
        }

        private void poi_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (Poisk_po.SelectedIndex == 0)
            {
                поиск = "idSotrudnika like '%" + Poisk.Text + "%' ";
            }
            else if (Poisk_po.SelectedIndex == 1)
            {
                поиск = "FIO like '%" + Poisk.Text + "%' ";
            }
            else if (Poisk_po.SelectedIndex == 2)
            {
                поиск = "Pasport like '%" + Poisk.Text + "%' ";
            }
            else if (Poisk_po.SelectedIndex == 3)
            {
                поиск = "Dataro like '%" + Poisk.Text + "%' ";
            }
            else if (Poisk_po.SelectedIndex == 4)
            {
                поиск = "Adres like '%" + Poisk.Text + "%' ";
            }
            else if (Poisk_po.SelectedIndex == 5)
            {
                поиск = "dolzhnost like '%" + Poisk.Text + "%' ";
            }
            else if (Poisk_po.SelectedIndex == 6)
            {
                поиск = "Telefon like '%" + Poisk.Text + "%' ";
            }
            else if (Poisk_po.SelectedIndex == 7)
            {
                поиск = "login like '%" + Poisk.Text + "%' ";
            }
            else if (Poisk_po.SelectedIndex == 8)
            {
                поиск = "parol like '%" + Poisk.Text + "%' ";
            }

            Loadsot();
        }

        private void ASCsort_Checked(object sender, RoutedEventArgs e)
        {
            if (Sortirovka_po.SelectedIndex == 0)
            {
                сортировка = "order by idSotrudnika asc";
            }
            else if (Sortirovka_po.SelectedIndex == 1)
            {
                сортировка = "order by FIO asc";
            }
            else if (Sortirovka_po.SelectedIndex == 2)
            {
                сортировка = "order by Pasport asc";
            }
            else if (Sortirovka_po.SelectedIndex == 3)
            {
                сортировка = "order by Dataro asc";
            }
            else if (Sortirovka_po.SelectedIndex == 4)
            {
                сортировка = "order by Adres asc";
            }
            else if (Sortirovka_po.SelectedIndex == 5)
            {
                сортировка = "order by dolzhnost asc";
            }
            else if (Sortirovka_po.SelectedIndex == 6)
            {
                сортировка = "order by Telefon asc";
            }
            else if (Sortirovka_po.SelectedIndex == 7)
            {
                сортировка = "order by login asc";
            }
            else if (Sortirovka_po.SelectedIndex == 8)
            {
                сортировка = "order by parol asc";
            }
            Loadsot();
        }

        private void DESCsort_Checked(object sender, RoutedEventArgs e)
        {
            if (Sortirovka_po.SelectedIndex == 0)
            {
                сортировка = "order by idSotrudnika desc";
            }
            else if (Sortirovka_po.SelectedIndex == 1)
            {
                сортировка = "order by FIO desc";
            }
            else if (Sortirovka_po.SelectedIndex == 2)
            {
                сортировка = "order by Pasport desc";
            }
            else if (Sortirovka_po.SelectedIndex == 3)
            {
                сортировка = "order by Dataro desc";
            }
            else if (Sortirovka_po.SelectedIndex == 4)
            {
                сортировка = "order by Adres desc";
            }
            else if (Sortirovka_po.SelectedIndex == 5)
            {
                сортировка = "order by dolzhnost desc";
            }
            else if (Sortirovka_po.SelectedIndex == 6)
            {
                сортировка = "order by Telefon desc";
            }
            else if (Sortirovka_po.SelectedIndex == 7)
            {
                сортировка = "order by login desc";
            }
            else if (Sortirovka_po.SelectedIndex == 8)
            {
                сортировка = "order by parol desc";
            }
            Loadsot();
        }

        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            Poisk_po.SelectedIndex = -1;
            Sortirovka_po.SelectedIndex = -1;
            Poisk.Clear();
        }

        private void Button_Click_5(object sender, RoutedEventArgs e)
        {
            fio.Clear();
            pasport.Clear();
            adr.Clear();
            tel.Clear();
            login1.Clear();
            parol.Clear();
            dol.SelectedIndex = -1;
            MessageBox.Show("Поля очищены");
        }

        private void Button_Click_6(object sender, RoutedEventArgs e)
        {
            DataRowView row = (DataRowView)sotrr.SelectedItems[0];

            Microsoft.Office.Interop.Word.Document doc = null;
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            string source = @"C:\Users\shaki\Desktop\My Fucking Works\парик\вывод с парика\сотрудники\Сотрудники вывод.docx";
            doc = app.Documents.Open(source); //открываем документ

            Microsoft.Office.Interop.Word.Bookmarks wBoookmarks = doc.Bookmarks; //Закладки
            Microsoft.Office.Interop.Word.Range wRange;

            int i = 0;

            string FIO = row["ФИО Сотрудника"].ToString();
            string dol = row["Должность"].ToString();
            string id = row["Номер"].ToString();
            
            string Pasport = row["Паспорт"].ToString();
            string dadaro = row["Дата рождения"].ToString();
            string adr = row["Адрес"].ToString();
            
            string tel = row["Телефон"].ToString();
            string Login = row["Логин"].ToString();
            string Parol = row["Пароль"].ToString();
            string[] data = new string[9] { FIO, dol, tel, id, Pasport, dadaro, adr, Login, Parol };

            foreach (Microsoft.Office.Interop.Word.Bookmark mark in wBoookmarks)
            {
                wRange = mark.Range;
                wRange.Text = data[i];
                i++;
            }

            doc = null;
            Process.Start(@"C:\Users\shaki\Desktop\My Fucking Works\парик\вывод с парика\сотрудники\Сотрудники вывод.docx");
        }

        private void Button_Click_7(object sender, RoutedEventArgs e)
        {
            DataRowView row = (DataRowView)sotrr.SelectedItems[0];

            var Excel = new Microsoft.Office.Interop.Excel.Application();
            var xlWB = Excel.Workbooks.Open(@"C:\Users\shaki\Desktop\My Fucking Works\парик\вывод с парика\сотрудники\Сотрудники вывод.xlsx");
            var xlSht = xlWB.Worksheets[1];

            xlSht.Cells[25, 1] = row["Номер"];
            xlSht.Cells[17, 2] = row["ФИО Сотрудника"];
            xlSht.Cells[25, 2] = row["Паспорт"];
            xlSht.Cells[25, 3] = row["Дата рождения"];
            xlSht.Cells[25, 4] = row["Адрес"];
            xlSht.Cells[19, 2] = row["Должность"];
            xlSht.Cells[21, 2] = row["Телефон"];
            xlSht.Cells[25, 6] = row["Логин"];
            xlSht.Cells[25, 7] = row["Пароль"];

            Excel.Visible = true;
        }

        private void Poisk_po_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            поиск = "idSotrudnika like '%%' ";
            Poisk_po.IsEnabled = true;

            Loadsot();
        }
    }
}

