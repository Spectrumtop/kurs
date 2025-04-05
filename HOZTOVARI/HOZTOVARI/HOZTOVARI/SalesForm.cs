

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
using Microsoft.Office.Interop.Word;
using DataTable = System.Data.DataTable;
using Application = Microsoft.Office.Interop.Word.Application;
using MySql.Data.MySqlClient;
namespace HOZTOVARI
{
    public partial class SalesForm : Form
    {
        private Form1 authForm;
        private string connectionString = "server=localhost;user=root;database=hoztovari;port=3306;password=root;";

        private bool isProcessing = false;
        private Application wordApp = null;

        public SalesForm(Form1 authForm)
        {
            this.authForm = authForm;
            InitializeComponent();
            InitializeComponents();
        }

        private void InitializeComponents()
        {
            dataGridViewProducts.DataError += dataGridViewProducts_DataError;
            dataGridViewProducts.RowTemplate.Height = 100;

            LoadProducts();
            LoadCategories();
            LoadSuppliers();
            LoadCustomers();
            InitializeSearchAndFilter();
            InitializeSaleForm();
        }

        private void InitializeSaleForm()
        {
            using (MySqlConnection conn = new MySqlConnection(connectionString))
            {
                conn.Open();
                string query = "SELECT ProductID, Name FROM products";
                MySqlDataAdapter adapter = new MySqlDataAdapter(query, conn);
                DataTable dt = new DataTable();
                adapter.Fill(dt);
                comboBoxProducts.DataSource = dt;
                comboBoxProducts.DisplayMember = "Name";
                comboBoxProducts.ValueMember = "ProductID";
            }

            txtQuantity.TextChanged += UpdateTotalPrice;
            btnClearFilters.Click += btnClearFilters_Click;
            btnUploadImage.Click += btnUploadImage_Click;
            btnGoToCustomerForm.Click += btnGoToCustomerForm_Click;
            btnDeleteImage.Click += btnDeleteImage_Click;
            btnExit.Click += btnExit_Click;
            btnSell.Click += btnSell_Click;
        }

        private void LoadProducts()
        {
            try
            {
                using (MySqlConnection conn = new MySqlConnection(connectionString))
                {
                    conn.Open();
                    string query = "SELECT ProductID AS 'ID товара', Name AS 'Название', Price AS 'Цена', StockQuantity AS 'Количество на складе', Image FROM products";
                    MySqlDataAdapter adapter = new MySqlDataAdapter(query, conn);
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);

                    if (!dt.Columns.Contains("ImageColumn"))
                    {
                        dt.Columns.Add("ImageColumn", typeof(Image));
                    }

                    foreach (DataRow row in dt.Rows)
                    {
                        try
                        {
                            if (row["Image"] != DBNull.Value && row["Image"] != null)
                            {
                                byte[] imageBytes = (byte[])row["Image"];
                                using (MemoryStream ms = new MemoryStream(imageBytes))
                                {
                                    row["ImageColumn"] = Image.FromStream(ms);
                                }
                            }
                            else
                            {
                                row["ImageColumn"] = null;
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Ошибка обработки изображения: " + ex.Message);
                            row["ImageColumn"] = null;
                        }
                    }

                    dataGridViewProducts.DataSource = dt;

                    if (dataGridViewProducts.Columns["ImageColumn"] == null)
                    {
                        DataGridViewImageColumn imageColumn = new DataGridViewImageColumn();
                        imageColumn.Name = "ImageColumn";
                        imageColumn.HeaderText = "Изображение";
                        imageColumn.ImageLayout = DataGridViewImageCellLayout.Zoom;
                        dataGridViewProducts.Columns.Add(imageColumn);
                    }

                    if (dataGridViewProducts.Columns["Image"] != null)
                    {
                        dataGridViewProducts.Columns["Image"].Visible = false;
                    }

                    dataGridViewProducts.RowTemplate.Height = 100;
                    dataGridViewProducts.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка загрузки товаров: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadCategories()
        {
            using (MySqlConnection conn = new MySqlConnection(connectionString))
            {
                conn.Open();
                string query = "SELECT CategoryID, Name FROM categories";
                MySqlDataAdapter adapter = new MySqlDataAdapter(query, conn);
                DataTable dt = new DataTable();
                adapter.Fill(dt);
                comboBoxCategory.DataSource = dt;
                comboBoxCategory.DisplayMember = "Name";
                comboBoxCategory.ValueMember = "CategoryID";
                comboBoxCategory.SelectedIndex = -1;
            }
        }

        private void LoadSuppliers()
        {
            using (MySqlConnection conn = new MySqlConnection(connectionString))
            {
                conn.Open();
                string query = "SELECT SupplierID, Name FROM suppliers";
                MySqlDataAdapter adapter = new MySqlDataAdapter(query, conn);
                DataTable dt = new DataTable();
                adapter.Fill(dt);
                comboBoxSupplier.DataSource = dt;
                comboBoxSupplier.DisplayMember = "Name";
                comboBoxSupplier.ValueMember = "SupplierID";
                comboBoxSupplier.SelectedIndex = -1;
            }
        }

        private void LoadCustomers()
        {
            using (MySqlConnection conn = new MySqlConnection(connectionString))
            {
                conn.Open();
                string query = "SELECT CustomerID, Name, PhoneNumber FROM customers";
                MySqlDataAdapter adapter = new MySqlDataAdapter(query, conn);
                DataTable dt = new DataTable();
                adapter.Fill(dt);
                comboBoxCustomer.DataSource = dt;
                comboBoxCustomer.DisplayMember = "Name";
                comboBoxCustomer.ValueMember = "CustomerID";
            }
        }

        private void InitializeSearchAndFilter()
        {
            txtSearch.TextChanged += SearchProducts;
            comboBoxCategory.SelectedIndexChanged += FilterProducts;
            comboBoxSupplier.SelectedIndexChanged += FilterProducts;
        }

        private void SearchProducts(object sender, EventArgs e)
        {
            string searchText = txtSearch.Text;
            string filterCategory = comboBoxCategory.SelectedValue?.ToString();
            string filterSupplier = comboBoxSupplier.SelectedValue?.ToString();

            using (MySqlConnection conn = new MySqlConnection(connectionString))
            {
                conn.Open();
                string query = @"
SELECT ProductID AS 'ID товара', Name AS 'Название', Price AS 'Цена', StockQuantity AS 'Количество на складе', Image 
FROM products 
WHERE (Name LIKE @search OR Price LIKE @search OR StockQuantity LIKE @search)";

                if (!string.IsNullOrEmpty(filterCategory))
                    query += " AND CategoryID = @categoryId";
                if (!string.IsNullOrEmpty(filterSupplier))
                    query += " AND SupplierID = @supplierId";

                MySqlCommand cmd = new MySqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@search", $"%{searchText}%");
                if (!string.IsNullOrEmpty(filterCategory))
                    cmd.Parameters.AddWithValue("@categoryId", filterCategory);
                if (!string.IsNullOrEmpty(filterSupplier))
                    cmd.Parameters.AddWithValue("@supplierId", filterSupplier);

                MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                adapter.Fill(dt);

                if (!dt.Columns.Contains("ImageColumn"))
                {
                    dt.Columns.Add("ImageColumn", typeof(Image));
                }

                foreach (DataRow row in dt.Rows)
                {
                    try
                    {
                        if (row["Image"] != DBNull.Value && row["Image"] != null)
                        {
                            byte[] imageBytes = (byte[])row["Image"];
                            using (MemoryStream ms = new MemoryStream(imageBytes))
                            {
                                row["ImageColumn"] = Image.FromStream(ms);
                            }
                        }
                        else
                        {
                            row["ImageColumn"] = null;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Ошибка обработки изображения: " + ex.Message);
                        row["ImageColumn"] = null;
                    }
                }

                dataGridViewProducts.DataSource = dt;
            }
        }

        private void FilterProducts(object sender, EventArgs e)
        {
            SearchProducts(sender, e);
        }

        private void btnSell_Click(object sender, EventArgs e)
        {
            if (isProcessing) return;

            isProcessing = true;
            btnSell.Enabled = false;
            Cursor = Cursors.WaitCursor;

            try
            {
                if (dataGridViewProducts.CurrentRow == null)
                {
                    MessageBox.Show("Выберите товар для продажи.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (comboBoxCustomer.SelectedValue == null)
                {
                    MessageBox.Show("Выберите клиента.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (!int.TryParse(txtQuantity.Text, out int quantity) || quantity <= 0)
                {
                    MessageBox.Show("Введите корректное количество товара.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                int productId = (int)dataGridViewProducts.CurrentRow.Cells["ID товара"].Value;
                int customerId = (int)comboBoxCustomer.SelectedValue;

                int stockQuantity = GetStockQuantity(productId);
                if (quantity > stockQuantity)
                {
                    MessageBox.Show($"Недостаточно товара на складе. Доступно: {stockQuantity}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                decimal totalPrice = CalculateTotalPrice(productId, quantity);

                DialogResult result = MessageBox.Show(
                    $"Вы уверены, что хотите оформить продажу?\nТовар: {dataGridViewProducts.CurrentRow.Cells["Название"].Value}\nКоличество: {quantity}\nОбщая сумма: {totalPrice:C}",
                    "Подтверждение продажи",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    int orderId = ProcessSale(customerId, productId, quantity, totalPrice);
                    if (orderId > 0)
                    {
                        MessageBox.Show("Продажа оформлена!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        LoadProducts();
                        GenerateReceipt(orderId, customerId, productId, quantity, totalPrice);
                    }
                }
            }
            finally
            {
                isProcessing = false;
                btnSell.Enabled = true;
                Cursor = Cursors.Default;
            }
        }

        private int ProcessSale(int customerId, int productId, int quantity, decimal totalPrice)
        {
            using (MySqlConnection conn = new MySqlConnection(connectionString))
            {
                conn.Open();
                using (MySqlTransaction transaction = conn.BeginTransaction())
                {
                    try
                    {
                        string orderQuery = @"
INSERT INTO orders (CustomerID, OrderDate, TotalAmount) 
VALUES (@customerId, @orderDate, @totalAmount);
SELECT LAST_INSERT_ID();";
                        MySqlCommand orderCmd = new MySqlCommand(orderQuery, conn, transaction);
                        orderCmd.Parameters.AddWithValue("@customerId", customerId);
                        orderCmd.Parameters.AddWithValue("@orderDate", DateTime.Now);
                        orderCmd.Parameters.AddWithValue("@totalAmount", totalPrice);
                        int orderId = Convert.ToInt32(orderCmd.ExecuteScalar());

                        string orderItemQuery = @"
INSERT INTO orderitems (OrderID, ProductID, Quantity, Price) 
VALUES (@orderId, @productId, @quantity, @price)";
                        MySqlCommand orderItemCmd = new MySqlCommand(orderItemQuery, conn, transaction);
                        orderItemCmd.Parameters.AddWithValue("@orderId", orderId);
                        orderItemCmd.Parameters.AddWithValue("@productId", productId);
                        orderItemCmd.Parameters.AddWithValue("@quantity", quantity);
                        orderItemCmd.Parameters.AddWithValue("@price", totalPrice / quantity);
                        orderItemCmd.ExecuteNonQuery();

                        string updateStockQuery = "UPDATE products SET StockQuantity = StockQuantity - @quantity WHERE ProductID = @productId";
                        MySqlCommand updateStockCmd = new MySqlCommand(updateStockQuery, conn, transaction);
                        updateStockCmd.Parameters.AddWithValue("@quantity", quantity);
                        updateStockCmd.Parameters.AddWithValue("@productId", productId);
                        updateStockCmd.ExecuteNonQuery();

                        transaction.Commit();
                        return orderId;
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        MessageBox.Show("Ошибка при оформлении продажи: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return -1;
                    }
                }
            }
        }

        private void GenerateReceipt(int orderId, int customerId, int productId, int quantity, decimal totalPrice)
        {
            try
            {
                if (wordApp == null)
                {
                    wordApp = new Application();
                    wordApp.Visible = false;
                }

                Document doc = wordApp.Documents.Add();

                AddReceiptParagraph(doc, "Чек о продаже", true, 16, WdParagraphAlignment.wdAlignParagraphCenter);
                AddReceiptParagraph(doc, $"Номер заказа: {orderId}\nДата: {DateTime.Now}");
                AddReceiptParagraph(doc, GetCustomerInfo(customerId));
                AddReceiptParagraph(doc, GetProductInfo(productId));
                AddReceiptParagraph(doc, $"Количество: {quantity}\nОбщая сумма: {totalPrice:C}");

                string receiptPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), $"Receipt_Order_{orderId}.docx");
                doc.SaveAs2(receiptPath);
                doc.Close(false);

                MessageBox.Show($"Чек сохранен в файл: {receiptPath}", "Чек о продаже", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при создании чека: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AddReceiptParagraph(Document doc, string text, bool bold = false, int fontSize = 12, WdParagraphAlignment alignment = WdParagraphAlignment.wdAlignParagraphLeft)
        {
            Paragraph paragraph = doc.Content.Paragraphs.Add();
            paragraph.Range.Text = text;
            paragraph.Range.Font.Bold = bold ? 1 : 0;
            paragraph.Range.Font.Size = fontSize;
            paragraph.Range.ParagraphFormat.Alignment = alignment;
            paragraph.Range.InsertParagraphAfter();
        }

        private string GetCustomerInfo(int customerId)
        {
            using (MySqlConnection conn = new MySqlConnection(connectionString))
            {
                conn.Open();
                string query = "SELECT Name, PhoneNumber FROM customers WHERE CustomerID = @customerId";
                MySqlCommand cmd = new MySqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@customerId", customerId);
                using (MySqlDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        return $"Клиент: {reader["Name"]}\nТелефон: {reader["PhoneNumber"]}";
                    }
                }
            }
            return "Клиент не найден";
        }

        private string GetProductInfo(int productId)
        {
            using (MySqlConnection conn = new MySqlConnection(connectionString))
            {
                conn.Open();
                string query = "SELECT Name, Price FROM products WHERE ProductID = @productId";
                MySqlCommand cmd = new MySqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@productId", productId);
                using (MySqlDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        return $"Товар: {reader["Name"]}\nЦена за единицу: {Convert.ToDecimal(reader["Price"]):C}";
                    }
                }
            }
            return "Товар не найден";
        }

        private int GetStockQuantity(int productId)
        {
            using (MySqlConnection conn = new MySqlConnection(connectionString))
            {
                conn.Open();
                string query = "SELECT StockQuantity FROM products WHERE ProductID = @productId";
                MySqlCommand cmd = new MySqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@productId", productId);
                return Convert.ToInt32(cmd.ExecuteScalar());
            }
        }

        private decimal CalculateTotalPrice(int productId, int quantity)
        {
            decimal price = 0;

            using (MySqlConnection conn = new MySqlConnection(connectionString))
            {
                conn.Open();
                string query = "SELECT Price FROM products WHERE ProductID = @productId";
                MySqlCommand cmd = new MySqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@productId", productId);
                price = Convert.ToDecimal(cmd.ExecuteScalar());
            }

            decimal totalPrice = price * quantity;

            if (quantity >= 3)
            {
                int discountQuantity = quantity / 3;
                decimal discountAmount = price * 0.4m * discountQuantity;
                totalPrice -= discountAmount;
            }

            return totalPrice;
        }

        private void btnUploadImage_Click(object sender, EventArgs e)
        {
            if (dataGridViewProducts.CurrentRow == null)
            {
                MessageBox.Show("Выберите товар для загрузки изображения.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    Image originalImage = Image.FromFile(openFileDialog.FileName);
                    Bitmap resizedImage = new Bitmap(100, 100);
                    using (Graphics g = Graphics.FromImage(resizedImage))
                    {
                        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                        g.DrawImage(originalImage, 0, 0, 100, 100);
                    }

                    byte[] imageBytes;
                    using (MemoryStream ms = new MemoryStream())
                    {
                        resizedImage.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                        imageBytes = ms.ToArray();
                    }

                    int productId = (int)dataGridViewProducts.CurrentRow.Cells["ID товара"].Value;

                    using (MySqlConnection conn = new MySqlConnection(connectionString))
                    {
                        conn.Open();
                        string query = "UPDATE products SET Image = @Image WHERE ProductID = @ProductID";
                        MySqlCommand cmd = new MySqlCommand(query, conn);
                        cmd.Parameters.AddWithValue("@Image", imageBytes);
                        cmd.Parameters.AddWithValue("@ProductID", productId);
                        cmd.ExecuteNonQuery();
                    }

                    LoadProducts();
                    MessageBox.Show("Изображение успешно загружено!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка при загрузке изображения: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btnDeleteImage_Click(object sender, EventArgs e)
        {
            if (dataGridViewProducts.CurrentRow == null)
            {
                MessageBox.Show("Выберите товар для удаления изображения.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int productId = (int)dataGridViewProducts.CurrentRow.Cells["ID товара"].Value;

            DialogResult result = MessageBox.Show(
                "Вы уверены, что хотите удалить изображение?",
                "Подтверждение удаления",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                using (MySqlConnection conn = new MySqlConnection(connectionString))
                {
                    conn.Open();
                    string query = "UPDATE products SET Image = NULL WHERE ProductID = @ProductID";
                    MySqlCommand cmd = new MySqlCommand(query, conn);
                    cmd.Parameters.AddWithValue("@ProductID", productId);
                    cmd.ExecuteNonQuery();
                }

                LoadProducts();
                MessageBox.Show("Изображение успешно удалено!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnGoToCustomerForm_Click(object sender, EventArgs e)
        {
            CustomerForm customerForm = new CustomerForm();
            customerForm.ShowDialog();
            LoadCustomers();
        }

        private void btnClearFilters_Click(object sender, EventArgs e)
        {
            txtSearch.Text = string.Empty;
            comboBoxCategory.SelectedIndex = -1;
            comboBoxSupplier.SelectedIndex = -1;
            LoadProducts();
        }

        private void UpdateTotalPrice(object sender, EventArgs e)
        {
            if (comboBoxProducts.SelectedValue == null || string.IsNullOrEmpty(txtQuantity.Text) || !int.TryParse(txtQuantity.Text, out int quantity))
            {
                lblTotalPrice.Text = "0 ₽";
                return;
            }

            int productId = (int)comboBoxProducts.SelectedValue;
            decimal totalPrice = CalculateTotalPrice(productId, quantity);
            lblTotalPrice.Text = totalPrice.ToString("C");
        }

        private void dataGridViewProducts_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            Console.WriteLine("Ошибка в DataGridView: " + e.Exception.Message);
            e.ThrowException = false;
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
                if (wordApp != null)
                {
                    wordApp.Quit(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
                    wordApp = null;
                }
                this.Close();
                authForm.Show();
            }
        }
    }
}