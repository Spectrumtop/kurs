using Microsoft.Office.Interop.Word;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HOZTOVARI
{
    public partial class AdminForm : Form
    {
        private Form1 authForm;
        private string connectionString = "server=localhost;user=root;database=hoztovari;port=3306;password=root;";

        public AdminForm(Form1 form1)
        {
            this.authForm = form1;
            InitializeComponent();
            LoadData();
            InitializeSearchAndFilter();
            LoadRoles();
        }

        private void LoadData()
        {
            try
            {
                LoadProducts();
                LoadUsers();
                LoadSuppliers();
                LoadOrders();
                LoadCategories();
                LoadSuppliersForFilter();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке данных: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void InitializeSearchAndFilter()
        {
            txtSearch.TextChanged += new EventHandler(SearchProducts);
            cmbCategoryFilter.SelectedIndexChanged += new EventHandler(FilterProducts);
            txtSearchSuppliers.TextChanged += new EventHandler(SearchSuppliers);
            cmbSupplierFilter.SelectedIndexChanged += new EventHandler(FilterSuppliers);
            btnResetSupplierFilter.Click += new EventHandler(btnResetSupplierFilter_Click);
            btnResetProductFilter.Click += new EventHandler(btnResetProductFilter_Click);
        }

        private void ExecuteNonQuery(string query, params MySqlParameter[] parameters)
        {
            using (MySqlConnection conn = new MySqlConnection(connectionString))
            {
                conn.Open();
                MySqlCommand cmd = new MySqlCommand(query, conn);
                cmd.Parameters.AddRange(parameters);
                cmd.ExecuteNonQuery();
            }
        }

        private System.Data.DataTable ExecuteQuery(string query, params MySqlParameter[] parameters)
        {
            using (MySqlConnection conn = new MySqlConnection(connectionString))
            {
                conn.Open();
                MySqlCommand cmd = new MySqlCommand(query, conn);
                cmd.Parameters.AddRange(parameters);
                MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
                System.Data.DataTable dt = new System.Data.DataTable();
                adapter.Fill(dt);

                // Безопасное переименование столбцов
                SafeRenameColumn(dt, "ProductID", "ID товара");
                SafeRenameColumn(dt, "Name", "Название");
                SafeRenameColumn(dt, "Price", "Цена");
                SafeRenameColumn(dt, "StockQuantity", "Количество на складе");
                SafeRenameColumn(dt, "CategoryName", "Категория");
                SafeRenameColumn(dt, "SupplierName", "Поставщик");
                SafeRenameColumn(dt, "UserID", "ID пользователя");
                SafeRenameColumn(dt, "Username", "Логин");
                SafeRenameColumn(dt, "Password", "Пароль");
                SafeRenameColumn(dt, "Role", "Роль");
                SafeRenameColumn(dt, "SupplierID", "ID поставщика");
                SafeRenameColumn(dt, "ContactInfo", "Контактная информация");
                SafeRenameColumn(dt, "OrderID", "ID заказа");
                SafeRenameColumn(dt, "OrderDate", "Дата заказа");
                SafeRenameColumn(dt, "TotalAmount", "Сумма заказа");
                SafeRenameColumn(dt, "UserName", "Пользователь");
                SafeRenameColumn(dt, "CategoryID", "ID категории");

                return dt;
            }
        }

        private void SafeRenameColumn(System.Data.DataTable dt, string originalName, string newName)
        {
            if (dt.Columns.Contains(originalName) && !dt.Columns.Contains(newName))
            {
                dt.Columns[originalName].ColumnName = newName;
            }
        }

        private void LoadProducts()
        {
            string query = @"
            SELECT p.ProductID, p.Name, p.Price, p.StockQuantity, c.Name AS CategoryName, s.Name AS SupplierName 
            FROM products p
            INNER JOIN categories c ON p.CategoryID = c.CategoryID
            INNER JOIN suppliers s ON p.SupplierID = s.SupplierID";
            dataGridViewProducts.DataSource = ExecuteQuery(query);
        }

        private void LoadUsers()
        {
            string query = "SELECT UserID, Username, Password, Role FROM users";
            dataGridViewCustomers.DataSource = ExecuteQuery(query);
        }

        private void LoadSuppliers()
        {
            string query = "SELECT SupplierID, Name, ContactInfo FROM suppliers";
            dataGridViewSuppliers.DataSource = ExecuteQuery(query);
        }

        private void LoadOrders()
        {
            string query = @"
            SELECT o.OrderID, o.OrderDate, o.TotalAmount, u.Username AS UserName 
            FROM orders o
            INNER JOIN users u ON o.UserID = u.UserID";
            dataGridViewOrders.DataSource = ExecuteQuery(query);
        }

        private void LoadCategories()
        {
            string query = "SELECT CategoryID, Name FROM categories";
            System.Data.DataTable dt = ExecuteQuery(query);
            cmbCategoryFilter.Items.Clear();
            cmbCategory.Items.Clear();
            foreach (System.Data.DataRow row in dt.Rows)
            {
                string name = row["Название"].ToString();
                cmbCategoryFilter.Items.Add(name);
                cmbCategory.Items.Add(name);
            }
        }

        private void LoadSuppliersForFilter()
        {
            string query = "SELECT SupplierID, Name FROM suppliers";
            System.Data.DataTable dt = ExecuteQuery(query);
            cmbSupplierFilter.Items.Clear();
            cmbSupplier.Items.Clear();
            foreach (System.Data.DataRow row in dt.Rows)
            {
                string name = row["Название"].ToString();
                cmbSupplierFilter.Items.Add(name);
                cmbSupplier.Items.Add(name);
            }
        }

        private void LoadRoles()
        {
            cmbEmployeeRole.Items.Clear();
            cmbEmployeeRole.Items.Add("Администратор");
            cmbEmployeeRole.Items.Add("Товаровед");
            cmbEmployeeRole.Items.Add("Продавец-консультант");
        }

        private void SearchProducts(object sender, EventArgs e)
        {
            try
            {
                string searchText = txtSearch.Text;
                string filterCategory = cmbCategoryFilter.SelectedItem?.ToString();

                string query = @"
                SELECT p.ProductID, p.Name, p.Price, p.StockQuantity, c.Name AS CategoryName, s.Name AS SupplierName 
                FROM products p
                INNER JOIN categories c ON p.CategoryID = c.CategoryID
                INNER JOIN suppliers s ON p.SupplierID = s.SupplierID
                WHERE p.Name LIKE @searchText";

                if (!string.IsNullOrEmpty(filterCategory))
                    query += " AND c.Name = @CategoryName";

                MySqlParameter[] parameters = {
                new MySqlParameter("@searchText", "%" + searchText + "%")
            };

                if (!string.IsNullOrEmpty(filterCategory))
                {
                    Array.Resize(ref parameters, parameters.Length + 1);
                    parameters[parameters.Length - 1] = new MySqlParameter("@CategoryName", filterCategory);
                }

                dataGridViewProducts.DataSource = ExecuteQuery(query, parameters);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при поиске товаров: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void FilterProducts(object sender, EventArgs e)
        {
            SearchProducts(sender, e);
        }

        private void SearchSuppliers(object sender, EventArgs e)
        {
            try
            {
                string searchText = txtSearchSuppliers.Text.Trim();

                if (string.IsNullOrEmpty(searchText))
                {
                    LoadSuppliers();
                    return;
                }

                string query = "SELECT SupplierID, Name, ContactInfo FROM suppliers WHERE Name LIKE @searchText OR ContactInfo LIKE @searchText";
                dataGridViewSuppliers.DataSource = ExecuteQuery(query, new MySqlParameter("@searchText", "%" + searchText + "%"));
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при поиске поставщиков: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void FilterSuppliers(object sender, EventArgs e)
        {
            try
            {
                string filterSupplier = cmbSupplierFilter.SelectedItem?.ToString();

                if (string.IsNullOrEmpty(filterSupplier))
                {
                    LoadSuppliers();
                    return;
                }

                string query = "SELECT SupplierID, Name, ContactInfo FROM suppliers WHERE Name = @SupplierName";
                dataGridViewSuppliers.DataSource = ExecuteQuery(query, new MySqlParameter("@SupplierName", filterSupplier));
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при фильтрации поставщиков: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnResetSupplierFilter_Click(object sender, EventArgs e)
        {
            cmbSupplierFilter.SelectedIndex = -1;
            txtSearchSuppliers.Clear();
            LoadSuppliers();
        }

        private void btnResetProductFilter_Click(object sender, EventArgs e)
        {
            txtSearch.Clear();
            cmbCategoryFilter.SelectedIndex = -1;
            LoadProducts();
        }

        private void btnAddProduct_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(txtProductName.Text) ||
                    string.IsNullOrEmpty(txtProductQuantity.Text) ||
                    string.IsNullOrEmpty(txtProductPrice.Text) ||
                    cmbCategory.SelectedItem == null ||
                    cmbSupplier.SelectedItem == null)
                {
                    MessageBox.Show("Заполните все поля.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                DialogResult result = MessageBox.Show(
                    "Вы уверены, что хотите добавить этот товар?",
                    "Подтверждение добавления",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    string name = txtProductName.Text;

                    if (!int.TryParse(txtProductQuantity.Text, out int quantity))
                    {
                        MessageBox.Show("Некорректное значение количества. Введите целое число.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    if (!decimal.TryParse(txtProductPrice.Text, out decimal price))
                    {
                        MessageBox.Show("Некорректное значение цены. Введите число.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    string category = cmbCategory.SelectedItem?.ToString();
                    string supplier = cmbSupplier.SelectedItem?.ToString();

                    string query = @"
                    INSERT INTO products (Name, StockQuantity, CategoryID, SupplierID, Price) 
                    VALUES (@Name, @StockQuantity, 
                    (SELECT CategoryID FROM categories WHERE Name = @CategoryName), 
                    (SELECT SupplierID FROM suppliers WHERE Name = @SupplierName), @Price)";

                    ExecuteNonQuery(query,
                        new MySqlParameter("@Name", name),
                        new MySqlParameter("@StockQuantity", quantity),
                        new MySqlParameter("@CategoryName", category),
                        new MySqlParameter("@SupplierName", supplier),
                        new MySqlParameter("@Price", price));

                    LoadProducts();
                    MessageBox.Show("Товар успешно добавлен.", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при добавлении товара: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnEditProduct_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridViewProducts.CurrentRow == null)
                {
                    MessageBox.Show("Выберите товар для редактирования.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (cmbCategory.SelectedItem == null || cmbSupplier.SelectedItem == null)
                {
                    MessageBox.Show("Выберите категорию и поставщика.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (string.IsNullOrEmpty(txtProductName.Text) ||
                    string.IsNullOrEmpty(txtProductQuantity.Text) ||
                    string.IsNullOrEmpty(txtProductPrice.Text))
                {
                    MessageBox.Show("Заполните все поля.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (!int.TryParse(txtProductQuantity.Text, out int quantity))
                {
                    MessageBox.Show("Некорректное значение количества. Введите целое число.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (!decimal.TryParse(txtProductPrice.Text, out decimal price))
                {
                    MessageBox.Show("Некорректное значение цены. Введите число.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                DialogResult result = MessageBox.Show(
                    "Вы уверены, что хотите сохранить изменения?",
                    "Подтверждение редактирования",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    int id = GetIntFromDataGridView(dataGridViewProducts, "ID товара", "ProductID");

                    string name = txtProductName.Text;
                    string category = cmbCategory.SelectedItem?.ToString();
                    string supplier = cmbSupplier.SelectedItem?.ToString();

                    string query = @"
                    UPDATE products 
                    SET Name = @Name, StockQuantity = @StockQuantity, 
                    CategoryID = (SELECT CategoryID FROM categories WHERE Name = @CategoryName), 
                    SupplierID = (SELECT SupplierID FROM suppliers WHERE Name = @SupplierName), 
                    Price = @Price 
                    WHERE ProductID = @ProductID";

                    ExecuteNonQuery(query,
                        new MySqlParameter("@Name", name),
                        new MySqlParameter("@StockQuantity", quantity),
                        new MySqlParameter("@CategoryName", category),
                        new MySqlParameter("@SupplierName", supplier),
                        new MySqlParameter("@Price", price),
                        new MySqlParameter("@ProductID", id));

                    LoadProducts();
                    MessageBox.Show("Товар успешно обновлен.", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при редактировании товара: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnDeleteProduct_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridViewProducts.CurrentRow == null)
                {
                    MessageBox.Show("Выберите товар для удаления.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                int id = GetIntFromDataGridView(dataGridViewProducts, "ID товара", "ProductID");

                DialogResult result = MessageBox.Show(
                    "Вы уверены, что хотите удалить этот товар?",
                    "Подтверждение удаления",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    string query = "DELETE FROM products WHERE ProductID = @ProductID";
                    ExecuteNonQuery(query, new MySqlParameter("@ProductID", id));
                    LoadProducts();
                    MessageBox.Show("Товар успешно удален.", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при удалении товара: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private int GetIntFromDataGridView(DataGridView dgv, params string[] possibleColumnNames)
        {
            foreach (var columnName in possibleColumnNames)
            {
                if (dgv.CurrentRow.Cells[columnName]?.Value != null)
                {
                    return Convert.ToInt32(dgv.CurrentRow.Cells[columnName].Value);
                }
            }
            throw new Exception("Не удалось найти значение в указанных столбцах");
        }

        private void btnGenerateReport_Click(object sender, EventArgs e)
        {
            try
            {
                string query = @"
                SELECT 
                    p.Name AS ProductName, 
                    SUM(oi.Quantity) AS TotalQuantitySold 
                FROM 
                    orderitems oi
                INNER JOIN 
                    products p ON oi.ProductID = p.ProductID
                GROUP BY 
                    p.Name
                ORDER BY 
                    TotalQuantitySold DESC";

                System.Data.DataTable dt = ExecuteQuery(query);

                if (dt.Rows.Count > 0)
                {
                    Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
                    Document wordDoc = wordApp.Documents.Add();
                    wordApp.Visible = true;

                    Paragraph title = wordDoc.Content.Paragraphs.Add();
                    title.Range.Text = "Отчёт о продажах";
                    title.Range.Font.Bold = 1;
                    title.Range.Font.Size = 16;
                    title.Range.InsertParagraphAfter();

                    Paragraph mostSold = wordDoc.Content.Paragraphs.Add();
                    mostSold.Range.Text = $"Товар, проданный в наибольшем количестве: {dt.Rows[0]["ProductName"]}\n";
                    mostSold.Range.Text += $"Количество проданных единиц: {dt.Rows[0]["TotalQuantitySold"]}\n\n";
                    mostSold.Range.InsertParagraphAfter();

                    Paragraph tableParagraph = wordDoc.Content.Paragraphs.Add();
                    Range tableRange = tableParagraph.Range;
                    Table salesTable = wordDoc.Tables.Add(tableRange, dt.Rows.Count + 1, 2);
                    salesTable.Borders.Enable = 1;

                    salesTable.Cell(1, 1).Range.Text = "Товар";
                    salesTable.Cell(1, 2).Range.Text = "Количество продаж";

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        salesTable.Cell(i + 2, 1).Range.Text = dt.Rows[i]["ProductName"].ToString();
                        salesTable.Cell(i + 2, 2).Range.Text = dt.Rows[i]["TotalQuantitySold"].ToString();
                    }

                    SaveFileDialog saveFileDialog = new SaveFileDialog();
                    saveFileDialog.Filter = "Word Documents|*.doc";
                    saveFileDialog.Title = "Сохранить отчёт";
                    saveFileDialog.FileName = "Отчёт_о_продажах.doc";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        wordDoc.SaveAs2(saveFileDialog.FileName);
                        MessageBox.Show("Отчёт успешно сохранён!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }

                    wordDoc.Close();
                    wordApp.Quit();
                }
                else
                {
                    MessageBox.Show("Нет данных о продажах.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при генерации отчёта: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show(
                "Вы уверены, что хотите выйти?",
                "Подтверждение выхода",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                this.Close();
                authForm.Show();
            }
        }

        private void btnAddEmployee_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(txtEmployeeName.Text) ||
                    string.IsNullOrEmpty(txtEmployeePassword.Text) ||
                    cmbEmployeeRole.SelectedItem == null)
                {
                    MessageBox.Show("Заполните все поля.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                DialogResult result = MessageBox.Show(
                    "Вы уверены, что хотите добавить этого сотрудника?",
                    "Подтверждение добавления",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    string employeeName = txtEmployeeName.Text;
                    string employeePassword = txtEmployeePassword.Text;
                    string employeeRole = cmbEmployeeRole.SelectedItem.ToString();

                    if (IsUserDuplicate(employeeName))
                    {
                        MessageBox.Show("Пользователь с таким именем уже существует.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    string query = "INSERT INTO users (Username, Password, Role) VALUES (@Username, @Password, @Role)";
                    ExecuteNonQuery(query,
                        new MySqlParameter("@Username", employeeName),
                        new MySqlParameter("@Password", employeePassword),
                        new MySqlParameter("@Role", employeeRole));

                    LoadUsers();

                    txtEmployeeName.Clear();
                    txtEmployeePassword.Clear();
                    cmbEmployeeRole.SelectedIndex = -1;

                    MessageBox.Show("Сотрудник успешно добавлен.", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при добавлении сотрудника: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnEditEmployee_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridViewCustomers.CurrentRow == null)
                {
                    MessageBox.Show("Выберите сотрудника для редактирования.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (string.IsNullOrEmpty(txtEmployeeName.Text) ||
                    string.IsNullOrEmpty(txtEmployeePassword.Text) ||
                    cmbEmployeeRole.SelectedItem == null)
                {
                    MessageBox.Show("Заполните все поля.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                DialogResult result = MessageBox.Show(
                    "Вы уверены, что хотите сохранить изменения?",
                    "Подтверждение редактирования",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    int userID = GetIntFromDataGridView(dataGridViewCustomers, "ID пользователя", "UserID");
                    string employeeName = txtEmployeeName.Text;
                    string employeePassword = txtEmployeePassword.Text;
                    string employeeRole = cmbEmployeeRole.SelectedItem.ToString();

                    string query = "UPDATE users SET Username = @Username, Password = @Password, Role = @Role WHERE UserID = @UserID";
                    ExecuteNonQuery(query,
                        new MySqlParameter("@Username", employeeName),
                        new MySqlParameter("@Password", employeePassword),
                        new MySqlParameter("@Role", employeeRole),
                        new MySqlParameter("@UserID", userID));

                    LoadUsers();

                    txtEmployeeName.Clear();
                    txtEmployeePassword.Clear();
                    cmbEmployeeRole.SelectedIndex = -1;

                    MessageBox.Show("Данные сотрудника успешно обновлены.", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при редактировании сотрудника: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnDeleteEmployee_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridViewCustomers.CurrentRow == null)
                {
                    MessageBox.Show("Выберите сотрудника для удаления.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                int userID = GetIntFromDataGridView(dataGridViewCustomers, "ID пользователя", "UserID");

                DialogResult result = MessageBox.Show(
                    "Вы уверены, что хотите удалить этого сотрудника?",
                    "Подтверждение удаления",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    string query = "DELETE FROM users WHERE UserID = @UserID";
                    ExecuteNonQuery(query, new MySqlParameter("@UserID", userID));
                    LoadUsers();

                    MessageBox.Show("Сотрудник успешно удален.", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при удалении сотрудника: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool IsUserDuplicate(string username)
        {
            string query = "SELECT COUNT(*) FROM users WHERE Username = @Username";
            System.Data.DataTable dt = ExecuteQuery(query, new MySqlParameter("@Username", username));
            return Convert.ToInt32(dt.Rows[0][0]) > 0;
        }
    }
}