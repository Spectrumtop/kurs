
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HOZTOVARI
{
    public partial class Form1 : Form
    {
        private bool isProcessing = false;
        private bool isClosing = false;

        public Form1()
        {
            InitializeComponent();
            InitializeControls();
        }

        private void InitializeControls()
        {
            // Настройка видимости пароля
            txtPassword.PasswordChar = '*';
            btnTogglePassword.Click += BtnTogglePassword_Click;

            // Привязка обработчиков событий
            btnLogin.Click += BtnLogin_Click;
            btnExit.Click += BtnExit_Click;

            // Обработчик закрытия формы
            this.FormClosing += Form1_FormClosing;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            isClosing = true;
        }

        private void BtnTogglePassword_Click(object sender, EventArgs e)
        {
            txtPassword.PasswordChar = txtPassword.PasswordChar == '*' ? '\0' : '*';
            btnTogglePassword.Text = txtPassword.PasswordChar == '*' ? "👁" : "👁";
        }

        private async void BtnLogin_Click(object sender, EventArgs e)
        {
            if (isProcessing || isClosing) return;
            isProcessing = true;
            btnLogin.Enabled = false;

            try
            {
                string login = txtLogin.Text.Trim();
                string password = txtPassword.Text.Trim();

                if (string.IsNullOrEmpty(login) || string.IsNullOrEmpty(password))
                {
                    MessageBox.Show("Заполните все поля!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (!IsValidInput(login) || !IsValidInput(password))
                {
                    MessageBox.Show("Логин и пароль должны содержать только латинские буквы, цифры и специальные символы!",
                        "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                using (var conn = new MySqlConnection("server=localhost;user=root;database=hoztovari;port=3306;password=root;"))
                {
                    await conn.OpenAsync();
                    string query = "SELECT role FROM users WHERE username=@username AND password=@password";

                    using (var cmd = new MySqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@username", login);
                        cmd.Parameters.AddWithValue("@password", password);

                        var result = await cmd.ExecuteScalarAsync();
                        if (result != null)
                        {
                            string role = result.ToString();
                            OpenRoleForm(role);
                        }
                        else
                        {
                            MessageBox.Show("Неверный логин или пароль!", "Ошибка",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка подключения: {ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (!isClosing)
                {
                    isProcessing = false;
                    btnLogin.Enabled = true;
                }
            }
        }

        private void OpenRoleForm(string role)
        {
            if (this.IsDisposed || isClosing) return;

            Form roleForm = null;
            if (role == "Администратор")
            {
                roleForm = new AdminForm(this);
            }
            else if (role == "Товаровед")
            {
                roleForm = new ManagerForm(this);
            }
            else if (role == "Продавец-консультант")
            {
                roleForm = new SalesForm(this);
            }

            if (roleForm != null && !roleForm.Visible)
            {
                this.Hide();
                roleForm.FormClosed += (s, args) =>
                {
                    if (!isClosing && !this.IsDisposed)
                    {
                        this.Show();
                    }
                };
                roleForm.Show();
            }
        }

        private bool IsValidInput(string input)
        {
            return Regex.IsMatch(input, @"^[a-zA-Z0-9!@#$%^&*()_+=\[{\]};:<>|./?,-]+$");
        }

        private void BtnExit_Click(object sender, EventArgs e)
        {
            isClosing = true;
            Application.Exit();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            isClosing = true;
        }
    }
}