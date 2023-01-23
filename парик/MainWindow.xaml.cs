using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
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

namespace парик
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

        private void Button_Click(object sender, RoutedEventArgs e)
        {

            if (login.Text.Length == 0 && parol.Text.Length == 0)
            {
                MessageBox.Show("Введите логин и пароль", "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            else
            {
                if (login.Text.Length > 0)
                {
                    if (parol.Text.Length > 0)
                    {
                        MySqlDataAdapter adapter = new MySqlDataAdapter("select idSotrudnika, FIO, dolzhnost from sotrudniki where login = '" + login.Text + "' and parol = '" + parol.Text + "'", DB.connection);
                        DataTable user = new DataTable();
                        adapter.Fill(user);

                        if (user.Rows.Count > 0)
                        {
                            ID.idSotrudnika = Convert.ToInt32(user.Rows[0][0]);
                            ID.FIO = user.Rows[0][1].ToString();
                            ID.dolzhnost = user.Rows[0][2].ToString();

                            if (ID.dolzhnost == "Администратор")
                            {
                                MessageBox.Show("Добро пожаловать, " + ID.FIO + "!", "Вход", MessageBoxButton.OK, MessageBoxImage.Information);

                                admin админ = new admin();
                                админ.Show();
                                this.Hide();
                            }
                            else if (ID.dolzhnost == "Барбер")
                            {
                                MessageBox.Show("Добро пожаловать, " + ID.FIO + "!", "Вход", MessageBoxButton.OK, MessageBoxImage.Information);

                                Барбер men = new Барбер();
                                men.Show();
                                this.Hide();
                            }
                        }
                        else
                        {
                            MessageBox.Show("Неверный логин или пароль", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Введите пароль", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
                else
                {
                    MessageBox.Show("Введите логин", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            Environment.Exit(0);
        }
    }
}
