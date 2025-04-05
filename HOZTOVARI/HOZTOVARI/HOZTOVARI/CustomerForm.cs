
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace HOZTOVARI
{
    public partial class CustomerForm : Form
    {
        private string connectionString = "server=localhost;user=root;database=hoztovari;port=3306;password=root;";

        public CustomerForm()
        {
            InitializeComponent();
            LoadCustomers();

            // Настройка маски для телефона
            txtPhoneNumber.Mask = "+70000000000";

            // Блокировка ввода английских букв, цифр и знаков
            txtName.KeyPress += new KeyPressEventHandler(BlockInvalidCharacters);
            txtPhoneNumber.KeyPress += new KeyPressEventHandler(BlockEnglishLetters);

            // Автоматическое начало ФИО с заглавной буквы
            txtName.TextChanged += new EventHandler(CapitalizeFirstLetter);
        }

        // Блокировка ввода английских букв, цифр и знаков
        private void BlockInvalidCharacters(object sender, KeyPressEventArgs e)
        {
            if (!Regex.IsMatch(e.KeyChar.ToString(), @"[А-Яа-я\s\-]") && !char.IsControl(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        // Блокировка ввода английских букв
        private void BlockEnglishLetters(object sender, KeyPressEventArgs e)
        {
            if (Regex.IsMatch(e.KeyChar.ToString(), @"[A-Za-z]"))
            {
                e.Handled = true;
            }
        }

        // Автоматическое начало ФИО с заглавной буквы
        private void CapitalizeFirstLetter(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(txtName.Text))
            {
                txtName.Text = char.ToUpper(txtName.Text[0]) + txtName.Text.Substring(1);
                txtName.SelectionStart = txtName.Text.Length;
            }
        }

        private void LoadCustomers()
        {
            using (MySqlConnection conn = new MySqlConnection(connectionString))
            {
                conn.Open();
                string query = "SELECT CustomerID AS 'ID клиента', Name AS 'ФИО', PhoneNumber AS 'Номер телефона' FROM customers";
                MySqlDataAdapter adapter = new MySqlDataAdapter(query, conn);
                DataTable dt = new DataTable();
                adapter.Fill(dt);
                dataGridViewCustomers.DataSource = dt;
            }
        }

        private bool ValidateFields()
        {
            if (string.IsNullOrEmpty(txtName.Text))
            {
                MessageBox.Show("Поле 'ФИО' обязательно для заполнения.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            if (string.IsNullOrEmpty(txtPhoneNumber.Text))
            {
                MessageBox.Show("Поле 'Телефон' обязательно для заполнения.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            return true;
        }

        private bool IsNameDuplicate(string name)
        {
            using (MySqlConnection conn = new MySqlConnection(connectionString))
            {
                conn.Open();
                string query = "SELECT COUNT(*) FROM customers WHERE Name = @Name";
                MySqlCommand cmd = new MySqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@Name", name);
                int count = Convert.ToInt32(cmd.ExecuteScalar());
                return count > 0;
            }
        }

        private bool IsPhoneNumberDuplicate(string phoneNumber)
        {
            using (MySqlConnection conn = new MySqlConnection(connectionString))
            {
                conn.Open();
                string query = "SELECT COUNT(*) FROM customers WHERE PhoneNumber = @PhoneNumber";
                MySqlCommand cmd = new MySqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@PhoneNumber", phoneNumber);
                int count = Convert.ToInt32(cmd.ExecuteScalar());
                return count > 0;
            }
        }

        private void btnAddCustomer_Click(object sender, EventArgs e)
        {
            string name = txtName.Text;
            string phoneNumber = txtPhoneNumber.Text;

            if (!ValidateFields())
                return;

            if (phoneNumber.Length > 20)
            {
                MessageBox.Show("Номер телефона слишком длинный.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (IsNameDuplicate(name))
            {
                MessageBox.Show("Клиент с таким именем уже существует.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (IsPhoneNumberDuplicate(phoneNumber))
            {
                MessageBox.Show("Клиент с таким номером телефона уже существует.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (MySqlConnection conn = new MySqlConnection(connectionString))
            {
                conn.Open();
                string query = "INSERT INTO customers (Name, PhoneNumber) VALUES (@name, @phoneNumber)";
                MySqlCommand cmd = new MySqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@name", name);
                cmd.Parameters.AddWithValue("@phoneNumber", phoneNumber);
                cmd.ExecuteNonQuery();
            }

            LoadCustomers();
            MessageBox.Show("Клиент успешно добавлен.", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnEditCustomer_Click(object sender, EventArgs e)
        {
            if (!ValidateFields())
                return;

            if (dataGridViewCustomers.CurrentRow == null)
            {
                MessageBox.Show("Выберите клиента для редактирования.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int customerId = (int)dataGridViewCustomers.CurrentRow.Cells["ID клиента"].Value;
            string name = txtName.Text;
            string phoneNumber = txtPhoneNumber.Text;

            if (IsNameDuplicate(name) && name != dataGridViewCustomers.CurrentRow.Cells["ФИО"].Value.ToString())
            {
                MessageBox.Show("Клиент с таким именем уже существует.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (IsPhoneNumberDuplicate(phoneNumber) && phoneNumber != dataGridViewCustomers.CurrentRow.Cells["Номер телефона"].Value.ToString())
            {
                MessageBox.Show("Клиент с таким номером телефона уже существует.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (MySqlConnection conn = new MySqlConnection(connectionString))
            {
                conn.Open();
                string query = "UPDATE customers SET Name = @name, PhoneNumber = @phoneNumber WHERE CustomerID = @customerId";
                MySqlCommand cmd = new MySqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@name", name);
                cmd.Parameters.AddWithValue("@phoneNumber", phoneNumber);
                cmd.Parameters.AddWithValue("@customerId", customerId);
                cmd.ExecuteNonQuery();
            }

            LoadCustomers();
            MessageBox.Show("Клиент успешно обновлен.", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnDeleteCustomer_Click(object sender, EventArgs e)
        {
            if (dataGridViewCustomers.CurrentRow == null)
            {
                MessageBox.Show("Выберите клиента для удаления.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int customerId = (int)dataGridViewCustomers.CurrentRow.Cells["ID клиента"].Value;

            using (MySqlConnection conn = new MySqlConnection(connectionString))
            {
                conn.Open();
                string query = "DELETE FROM customers WHERE CustomerID = @customerId";
                MySqlCommand cmd = new MySqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@customerId", customerId);
                cmd.ExecuteNonQuery();
            }

            LoadCustomers();
            MessageBox.Show("Клиент успешно удален.", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dataGridViewCustomers_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridViewCustomers.CurrentRow != null)
            {
                txtName.Text = dataGridViewCustomers.CurrentRow.Cells["ФИО"].Value.ToString();
                txtPhoneNumber.Text = dataGridViewCustomers.CurrentRow.Cells["Номер телефона"].Value.ToString();
            }
        }
    }
}