using Microsoft.Office.Interop.Word;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HOZTOVARI
{
    public partial class ManagerForm : Form
    {
        private Form1 authForm;
        private string connectionString = "server=localhost;user=root;database=hoztovari;port=3306;password=root;";

        public ManagerForm(Form1 authForm)
        {
            this.authForm = authForm;
            InitializeComponent();
            LoadData();
            InitializeSearchAndFilter();
        }

        private void LoadData()
        {
            try
            {
                LoadProducts();
                LoadSuppliers();
                LoadCategories();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке данных: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void InitializeSearchAndFilter()
        {
            cmbCategoryFilter.SelectedIndexChanged += new EventHandler(cmbCategoryFilter_SelectedIndexChanged);
            txtSearchSuppliers.TextChanged += new EventHandler(txtSearchSuppliers_TextChanged);
            txtSearchProducts.TextChanged += new EventHandler(txtSearchProducts_TextChanged);
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

                // Переименование столбцов на русский
                SafeRenameColumn(dt, "ProductID", "ID товара");
                SafeRenameColumn(dt, "Name", "Название");
                SafeRenameColumn(dt, "Price", "Цена");
                SafeRenameColumn(dt, "StockQuantity", "Количество");
                SafeRenameColumn(dt, "CategoryName", "Категория");
                SafeRenameColumn(dt, "SupplierName", "Поставщик");
                SafeRenameColumn(dt, "SupplierID", "ID поставщика");
                SafeRenameColumn(dt, "ContactInfo", "Контакты");
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

        private void LoadSuppliers()
        {
            string query = "SELECT SupplierID, Name, ContactInfo FROM suppliers";
            dataGridViewSuppliers.DataSource = ExecuteQuery(query);
        }

        private void LoadCategories()
        {
            string query = "SELECT CategoryID, Name FROM categories";
            System.Data.DataTable dt = ExecuteQuery(query);
            cmbCategoryFilter.Items.Clear();
            foreach (System.Data.DataRow row in dt.Rows)
            {
                cmbCategoryFilter.Items.Add(row["Название"].ToString());
            }
        }

        private void cmbCategoryFilter_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                string category = cmbCategoryFilter.SelectedItem?.ToString();

                if (string.IsNullOrEmpty(category))
                {
                    LoadSuppliers();
                    return;
                }

                string query = @"
                    SELECT DISTINCT s.SupplierID, s.Name, s.ContactInfo 
                    FROM suppliers s
                    INNER JOIN products p ON s.SupplierID = p.SupplierID
                    INNER JOIN categories c ON p.CategoryID = c.CategoryID
                    WHERE c.Name = @CategoryName";

                dataGridViewSuppliers.DataSource = ExecuteQuery(query,
                    new MySqlParameter("@CategoryName", category));
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при фильтрации: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnResetFilter_Click(object sender, EventArgs e)
        {
            cmbCategoryFilter.SelectedIndex = -1;
            LoadSuppliers();
        }

        private void btnAddSupplier_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(txtSupplierName.Text) || string.IsNullOrEmpty(txtSupplierContactInfo.Text))
                {
                    MessageBox.Show("Заполните все поля.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string name = txtSupplierName.Text;
                string contactInfo = txtSupplierContactInfo.Text;

                if (IsSupplierDuplicate(name, contactInfo))
                {
                    MessageBox.Show("Такой поставщик уже существует.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                DialogResult result = MessageBox.Show(
                    "Добавить нового поставщика?",
                    "Подтверждение",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    string query = "INSERT INTO suppliers (Name, ContactInfo) VALUES (@Name, @ContactInfo)";
                    ExecuteNonQuery(query,
                        new MySqlParameter("@Name", name),
                        new MySqlParameter("@ContactInfo", contactInfo));

                    LoadSuppliers();
                    MessageBox.Show("Поставщик добавлен.", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при добавлении: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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

        private bool IsSupplierDuplicate(string name, string contactInfo)
        {
            string query = "SELECT COUNT(*) FROM suppliers WHERE Name = @Name OR ContactInfo = @ContactInfo";
            System.Data.DataTable dt = ExecuteQuery(query,
                new MySqlParameter("@Name", name),
                new MySqlParameter("@ContactInfo", contactInfo));
            return Convert.ToInt32(dt.Rows[0][0]) > 0;
        }

        private void btnEditSupplier_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridViewSuppliers.CurrentRow == null)
                {
                    MessageBox.Show("Выберите поставщика.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                int supplierID = GetIntFromDataGridView(dataGridViewSuppliers, "ID поставщика", "SupplierID");
                string name = txtSupplierName.Text;
                string contactInfo = txtSupplierContactInfo.Text;

                DialogResult result = MessageBox.Show(
                    "Сохранить изменения?",
                    "Подтверждение",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    string query = "UPDATE suppliers SET Name = @Name, ContactInfo = @ContactInfo WHERE SupplierID = @SupplierID";
                    ExecuteNonQuery(query,
                        new MySqlParameter("@Name", name),
                        new MySqlParameter("@ContactInfo", contactInfo),
                        new MySqlParameter("@SupplierID", supplierID));

                    LoadSuppliers();
                    MessageBox.Show("Изменения сохранены.", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при редактировании: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            throw new Exception("Не удалось найти значение");
        }

        private void btnDeleteSupplier_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridViewSuppliers.CurrentRow == null)
                {
                    MessageBox.Show("Выберите поставщика.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                int supplierID = GetIntFromDataGridView(dataGridViewSuppliers, "ID поставщика", "SupplierID");

                DialogResult result = MessageBox.Show(
                    "Удалить поставщика и связанные товары?",
                    "Подтверждение",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);

                if (result == DialogResult.Yes)
                {
                    // Удаление связанных товаров
                    string deleteProductsQuery = "DELETE FROM products WHERE SupplierID = @SupplierID";
                    ExecuteNonQuery(deleteProductsQuery, new MySqlParameter("@SupplierID", supplierID));

                    // Удаление поставщика
                    string deleteSupplierQuery = "DELETE FROM suppliers WHERE SupplierID = @SupplierID";
                    ExecuteNonQuery(deleteSupplierQuery, new MySqlParameter("@SupplierID", supplierID));

                    LoadSuppliers();
                    MessageBox.Show("Поставщик удален.", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при удалении: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnGenerateReport_Click(object sender, EventArgs e)
        {
            try
            {
                string query = @"
                    SELECT 
                        p.Name AS ProductName, 
                        p.StockQuantity, 
                        c.Name AS CategoryName, 
                        s.Name AS SupplierName
                    FROM 
                        products p
                    INNER JOIN 
                        categories c ON p.CategoryID = c.CategoryID
                    INNER JOIN 
                        suppliers s ON p.SupplierID = s.SupplierID";

                System.Data.DataTable dt = ExecuteQuery(query);
                SaveReportToDoc(dt);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при создании отчета: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SaveReportToDoc(System.Data.DataTable dt)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Word Documents|*.doc";
            saveFileDialog.Title = "Сохранить отчет";
            saveFileDialog.FileName = "Отчет_по_товарам.doc";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    var wordApp = new Microsoft.Office.Interop.Word.Application();
                    var wordDoc = wordApp.Documents.Add();
                    wordApp.Visible = false;

                    // Заголовок
                    var title = wordDoc.Content.Paragraphs.Add();
                    title.Range.Text = "Отчёт по товарам";
                    title.Range.Font.Bold = 1;
                    title.Range.Font.Size = 16;
                    title.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    title.Range.InsertParagraphAfter();

                    // Таблица
                    var table = wordDoc.Tables.Add(wordDoc.Range(wordDoc.Content.End - 1), dt.Rows.Count + 1, dt.Columns.Count);
                    table.Borders.Enable = 1;

                    // Заголовки
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        table.Cell(1, i + 1).Range.Text = dt.Columns[i].ColumnName;
                        table.Cell(1, i + 1).Range.Font.Bold = 1;
                    }

                    // Данные
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            table.Cell(i + 2, j + 1).Range.Text = dt.Rows[i][j].ToString();
                        }
                    }

                    wordDoc.SaveAs2(saveFileDialog.FileName);
                    wordDoc.Close();
                    wordApp.Quit();

                    MessageBox.Show("Отчёт сохранён.", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при сохранении: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show(
                "Выйти из системы?",
                "Подтверждение",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                this.Close();
                authForm.Show();
            }
        }

        private void txtSearchSuppliers_TextChanged(object sender, EventArgs e)
        {
            string searchText = txtSearchSuppliers.Text.Trim();

            if (string.IsNullOrEmpty(searchText))
            {
                LoadSuppliers();
                return;
            }

            System.Data.DataTable dt = (System.Data.DataTable)dataGridViewSuppliers.DataSource;
            if (dt != null)
            {
                DataView dv = new DataView(dt);
                dv.RowFilter = $"[Название] LIKE '%{searchText}%' OR [Контакты] LIKE '%{searchText}%'";
                dataGridViewSuppliers.DataSource = dv.ToTable();
            }
        }

        private void txtSearchProducts_TextChanged(object sender, EventArgs e)
        {
            string searchText = txtSearchProducts.Text.Trim();

            if (string.IsNullOrEmpty(searchText))
            {
                LoadProducts();
                return;
            }

            System.Data.DataTable dt = (System.Data.DataTable)dataGridViewProducts.DataSource;
            if (dt != null)
            {
                DataView dv = new DataView(dt);
                dv.RowFilter = $"[Название] LIKE '%{searchText}%' OR [Категория] LIKE '%{searchText}%' OR [Поставщик] LIKE '%{searchText}%'";
                dataGridViewProducts.DataSource = dv.ToTable();
            }
        }
    }
}
