using System;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.Windows.Forms;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;

namespace InventoryApp
{
    public partial class MainForm : Form
    {
        private string connectionString = "Data Source=inventory.db;Version=3;";
        private DataTable turnoverDataTable;
        private DataTable writeOffDataTable;

        public MainForm()
        {
            InitializeComponent();
            this.KeyPreview = true;
            this.KeyDown += new KeyEventHandler(MainForm_KeyDown);
            LoadComboBoxes();
            LoadTurnoverData();
            LoadWriteOffData();
            SetupGrids();
            CalculateTurnoverTotals();
            CalculateWriteOffTotals();
            ApplyVisualStyles();
            this.Text = "Учет расходных материалов";
        }

        private void MainForm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F1)
            {
                string helpFilePath = Path.Combine(Application.StartupPath, "Руководство пользователя.chm");

                if (File.Exists(helpFilePath))
                {
                    System.Diagnostics.Process.Start(helpFilePath);
                }
                else
                {
                    MessageBox.Show("Help file not found!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void ApplyVisualStyles()
        {
            this.Font = new System.Drawing.Font("Arial", 10F);
            tabControl.Font = new System.Drawing.Font("Arial", 10F, System.Drawing.FontStyle.Bold);
            groupWriteOffInput.Font = new System.Drawing.Font("Arial", 10F, System.Drawing.FontStyle.Bold);
            groupTurnoverInput.Font = new System.Drawing.Font("Arial", 10F, System.Drawing.FontStyle.Bold);
            groupTurnoverData.Font = new System.Drawing.Font("Arial", 10F, System.Drawing.FontStyle.Bold);
            groupWriteOffHistory.Font = new System.Drawing.Font("Arial", 10F, System.Drawing.FontStyle.Bold);

            lblTurnoverTitle.ForeColor = Color.DarkBlue;
            lblWriteOffHistoryTitle.ForeColor = Color.DarkBlue;
            lblTurnoverTotalQuantity.ForeColor = Color.DarkRed;
            lblTurnoverTotalSum.ForeColor = Color.DarkRed;
            lblWriteOffTotalQuantity.ForeColor = Color.DarkRed;

            btnSaveWriteOff.BackColor = Color.LightGreen;
            btnAddTurnover.BackColor = Color.LightGreen;
            btnExportTurnover.BackColor = Color.LightBlue;
            btnExportTurnoverByMonth.BackColor = Color.LightBlue;
            btnExportWriteOffHistory.BackColor = Color.LightBlue;
        }

        private void LoadComboBoxes()
        {
            using (SQLiteConnection conn = new SQLiteConnection(connectionString))
            {
                conn.Open();

                SQLiteDataAdapter materialAdapter = new SQLiteDataAdapter("SELECT DISTINCT MaterialName FROM WriteOffs UNION SELECT DISTINCT MaterialName FROM Turnover", conn);
                DataTable materialTable = new DataTable();
                materialAdapter.Fill(materialTable);
                comboMaterialName.Items.Clear();
                foreach (DataRow row in materialTable.Rows)
                {
                    comboMaterialName.Items.Add(row["MaterialName"].ToString());
                }

                SQLiteDataAdapter unitAdapter = new SQLiteDataAdapter("SELECT DISTINCT Unit FROM WriteOffs UNION SELECT DISTINCT Unit FROM Turnover", conn);
                DataTable unitTable = new DataTable();
                unitAdapter.Fill(unitTable);
                comboUnit.Items.Clear();
                foreach (DataRow row in unitTable.Rows)
                {
                    comboUnit.Items.Add(row["Unit"].ToString());
                }

                SQLiteDataAdapter deptAdapter = new SQLiteDataAdapter("SELECT DISTINCT Department FROM WriteOffs", conn);
                DataTable deptTable = new DataTable();
                deptAdapter.Fill(deptTable);
                comboDepartment.Items.Clear();
                foreach (DataRow row in deptTable.Rows)
                {
                    comboDepartment.Items.Add(row["Department"].ToString());
                }

                conn.Close();
            }
        }

        private void LoadTurnoverData(string filter = "")
        {
            using (SQLiteConnection conn = new SQLiteConnection(connectionString))
            {
                conn.Open();
                string query = "SELECT * FROM Turnover";
                if (!string.IsNullOrEmpty(filter))
                {
                    query += " WHERE MaterialName LIKE @filter OR Unit LIKE @filter";
                }
                using (SQLiteDataAdapter adapter = new SQLiteDataAdapter(query, conn))
                {
                    if (!string.IsNullOrEmpty(filter))
                    {
                        adapter.SelectCommand.Parameters.AddWithValue("@filter", $"%{filter}%");
                    }
                    turnoverDataTable = new DataTable();
                    adapter.Fill(turnoverDataTable);
                    dgvTurnover.DataSource = turnoverDataTable;
                }
                conn.Close();
            }
            CalculateTurnoverTotals();
        }

        private void LoadWriteOffData(string filter = "", string dateFilter = "")
        {
            using (SQLiteConnection conn = new SQLiteConnection(connectionString))
            {
                conn.Open();
                string query = "SELECT * FROM WriteOffs";
                if (!string.IsNullOrEmpty(filter) || !string.IsNullOrEmpty(dateFilter))
                {
                    query += " WHERE 1=1";
                    if (!string.IsNullOrEmpty(filter))
                    {
                        query += " AND MaterialName LIKE @filter";
                    }
                    if (!string.IsNullOrEmpty(dateFilter))
                    {
                        query += " AND Date = @dateFilter";
                    }
                }
                using (SQLiteDataAdapter adapter = new SQLiteDataAdapter(query, conn))
                {
                    if (!string.IsNullOrEmpty(filter))
                    {
                        adapter.SelectCommand.Parameters.AddWithValue("@filter", $"%{filter}%");
                    }
                    if (!string.IsNullOrEmpty(dateFilter))
                    {
                        adapter.SelectCommand.Parameters.AddWithValue("@dateFilter", dateFilter);
                    }
                    writeOffDataTable = new DataTable();
                    adapter.Fill(writeOffDataTable);
                    dgvWriteOffHistory.DataSource = writeOffDataTable;
                }
                conn.Close();
            }
            CalculateWriteOffTotals();
        }

        private void SetupGrids()
        {
            dgvTurnover.Columns["Id"].Visible = false;
            dgvTurnover.Columns["MaterialName"].HeaderText = "Наименование";
            dgvTurnover.Columns["Unit"].HeaderText = "Ед.изм";
            dgvTurnover.Columns["Price"].HeaderText = "Цена";
            dgvTurnover.Columns["QuantityEnd"].HeaderText = "Кол-во";
            dgvTurnover.Columns["TotalEnd"].HeaderText = "Сумма";
            dgvTurnover.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            dgvWriteOffHistory.Columns["Id"].Visible = false;
            dgvWriteOffHistory.Columns["MaterialName"].HeaderText = "Наименование";
            dgvWriteOffHistory.Columns["Unit"].HeaderText = "Ед.изм";
            dgvWriteOffHistory.Columns["Quantity"].HeaderText = "Кол-во";
            dgvWriteOffHistory.Columns["Department"].HeaderText = "Отдел";
            dgvWriteOffHistory.Columns["DeviceName"].HeaderText = "Устройство";
            dgvWriteOffHistory.Columns["InventoryNumber"].HeaderText = "Инв.номер";
            dgvWriteOffHistory.Columns["Reason"].HeaderText = "Причина";
            dgvWriteOffHistory.Columns["Date"].HeaderText = "Дата";
            dgvWriteOffHistory.Columns["Note"].HeaderText = "Примечание";
            dgvWriteOffHistory.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }

        private void CalculateTurnoverTotals()
        {
            double totalQuantity = 0;
            double totalSum = 0;

            foreach (DataRow row in turnoverDataTable.Rows)
            {
                totalQuantity += Convert.ToDouble(row["QuantityEnd"]);
                totalSum += Convert.ToDouble(row["TotalEnd"]);
            }

            lblTurnoverTotalQuantity.Text = $"Итого Кол-во: {totalQuantity:F2}";
            lblTurnoverTotalSum.Text = $"Итого Сумма: {totalSum:F2}";
        }

        private void CalculateWriteOffTotals()
        {
            double totalQuantity = 0;

            foreach (DataRow row in writeOffDataTable.Rows)
            {
                if (row["Quantity"] != DBNull.Value && double.TryParse(row["Quantity"].ToString(), out double quantity))
                {
                    totalQuantity += quantity;
                }
            }

            lblWriteOffTotalQuantity.Text = $"Итого списано: {totalQuantity:F2}";
        }

        private void btnSaveWriteOff_Click(object sender, EventArgs e)
        {
            try
            {
                using (SQLiteConnection conn = new SQLiteConnection(connectionString))
                {
                    conn.Open();
                    string query = "INSERT INTO WriteOffs (MaterialName, Unit, Quantity, Department, DeviceName, InventoryNumber, Reason, Date, Note) " +
                                   "VALUES (@MaterialName, @Unit, @Quantity, @Department, @DeviceName, @InventoryNumber, @Reason, @Date, @Note)";
                    using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@MaterialName", comboMaterialName.Text);
                        cmd.Parameters.AddWithValue("@Unit", comboUnit.Text);
                        cmd.Parameters.AddWithValue("@Quantity", Convert.ToDouble(txtQuantity.Text));
                        cmd.Parameters.AddWithValue("@Department", comboDepartment.Text);
                        cmd.Parameters.AddWithValue("@DeviceName", txtDeviceName.Text);
                        cmd.Parameters.AddWithValue("@InventoryNumber", txtInventoryNumber.Text);
                        cmd.Parameters.AddWithValue("@Reason", txtReason.Text);
                        cmd.Parameters.AddWithValue("@Date", dtpWriteOffDate.Value.ToString("dd.MM.yyyy"));
                        cmd.Parameters.AddWithValue("@Note", txtNote.Text);

                        cmd.ExecuteNonQuery();
                    }
                    conn.Close();
                }
                MessageBox.Show("Запись сохранена!");
                LoadComboBoxes();
                LoadWriteOffData(txtWriteOffSearch.Text, dtpWriteOffFilterDate.Value.ToString("dd.MM.yyyy"));
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при сохранении: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnAddTurnover_Click(object sender, EventArgs e)
        {
            try
            {
                using (SQLiteConnection conn = new SQLiteConnection(connectionString))
                {
                    conn.Open();
                    string query = "INSERT INTO Turnover (MaterialName, Unit, Price, QuantityEnd, TotalEnd, Date) " +
                                   "VALUES (@MaterialName, @Unit, @Price, @QuantityEnd, @TotalEnd, @Date)";
                    using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
                    {
                        double price = Convert.ToDouble(txtTurnoverPrice.Text);
                        double quantity = Convert.ToDouble(txtTurnoverQuantity.Text);
                        double total = price * quantity;

                        cmd.Parameters.AddWithValue("@MaterialName", txtTurnoverMaterial.Text);
                        cmd.Parameters.AddWithValue("@Unit", txtTurnoverUnit.Text);
                        cmd.Parameters.AddWithValue("@Price", price);
                        cmd.Parameters.AddWithValue("@QuantityEnd", quantity);
                        cmd.Parameters.AddWithValue("@TotalEnd", total);
                        cmd.Parameters.AddWithValue("@Date", DateTime.Now.ToString("dd.MM.yyyy"));

                        cmd.ExecuteNonQuery();
                    }
                    conn.Close();
                }
                LoadTurnoverData(txtTurnoverSearch.Text);
                LoadComboBoxes();
                MessageBox.Show("Запись добавлена!");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при добавлении: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void txtTurnoverSearch_TextChanged(object sender, EventArgs e)
        {
            LoadTurnoverData(txtTurnoverSearch.Text);
        }

        private void txtWriteOffSearch_TextChanged(object sender, EventArgs e)
        {
            LoadWriteOffData(txtWriteOffSearch.Text, dtpWriteOffFilterDate.Value.ToString("dd.MM.yyyy"));
        }

        private void dtpWriteOffFilterDate_ValueChanged(object sender, EventArgs e)
        {
            LoadWriteOffData(txtWriteOffSearch.Text, dtpWriteOffFilterDate.Value.ToString("dd.MM.yyyy"));
        }

        private void btnExportTurnover_Click(object sender, EventArgs e)
        {
            try
            {
                using (SaveFileDialog sfd = new SaveFileDialog())
                {
                    sfd.Filter = "PDF files (*.pdf)|*.pdf";
                    sfd.FileName = "TurnoverReport.pdf";
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        Document doc = new Document(PageSize.A4, 30, 30, 30, 30);
                        PdfWriter writer = PdfWriter.GetInstance(doc, new FileStream(sfd.FileName, FileMode.Create));
                        doc.Open();

                        BaseFont baseFont = BaseFont.CreateFont("c:\\windows\\fonts\\arial.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                        iTextSharp.text.Font font = new iTextSharp.text.Font(baseFont, 12);
                        iTextSharp.text.Font headerFont = new iTextSharp.text.Font(baseFont, 14, iTextSharp.text.Font.BOLD);

                        doc.Add(new Paragraph("Оборотная ведомость по НФА", headerFont) { Alignment = Element.ALIGN_CENTER });
                        doc.Add(new Paragraph($"Период: с {dtpTurnoverStart.Value.ToString("dd.MM.yyyy")} по {dtpTurnoverEnd.Value.ToString("dd.MM.yyyy")}", font) { Alignment = Element.ALIGN_CENTER });
                        doc.Add(new Paragraph($"Дата формирования: {DateTime.Now.ToString("dd.MM.yyyy")}", font) { Alignment = Element.ALIGN_CENTER });
                        doc.Add(new Paragraph(" ", font));

                        PdfPTable table = new PdfPTable(5);
                        table.WidthPercentage = 100;
                        float[] widths = new float[] { 3f, 1f, 1f, 2f, 2f };
                        table.SetWidths(widths);

                        table.AddCell(new PdfPCell(new Phrase("Наименование", font)) { BackgroundColor = BaseColor.LIGHT_GRAY, HorizontalAlignment = Element.ALIGN_CENTER });
                        table.AddCell(new PdfPCell(new Phrase("Ед.изм", font)) { BackgroundColor = BaseColor.LIGHT_GRAY, HorizontalAlignment = Element.ALIGN_CENTER });
                        table.AddCell(new PdfPCell(new Phrase("Цена", font)) { BackgroundColor = BaseColor.LIGHT_GRAY, HorizontalAlignment = Element.ALIGN_CENTER });
                        table.AddCell(new PdfPCell(new Phrase("Остаток (Кол-во)", font)) { BackgroundColor = BaseColor.LIGHT_GRAY, HorizontalAlignment = Element.ALIGN_CENTER });
                        table.AddCell(new PdfPCell(new Phrase("Остаток (Сумма)", font)) { BackgroundColor = BaseColor.LIGHT_GRAY, HorizontalAlignment = Element.ALIGN_CENTER });

                        double totalQuantity = 0;
                        double totalSum = 0;

                        foreach (DataRow row in turnoverDataTable.Rows)
                        {
                            table.AddCell(new Phrase(row["MaterialName"].ToString(), font));
                            table.AddCell(new Phrase(row["Unit"].ToString(), font));
                            table.AddCell(new Phrase(Convert.ToDouble(row["Price"]).ToString("F2"), font));
                            table.AddCell(new Phrase(Convert.ToDouble(row["QuantityEnd"]).ToString("F2"), font));
                            table.AddCell(new Phrase(Convert.ToDouble(row["TotalEnd"]).ToString("F2"), font));
                            totalQuantity += Convert.ToDouble(row["QuantityEnd"]);
                            totalSum += Convert.ToDouble(row["TotalEnd"]);
                        }

                        doc.Add(table);
                        doc.Add(new Paragraph($"Итого Кол-во: {totalQuantity:F2}, Итого Сумма: {totalSum:F2}", font) { Alignment = Element.ALIGN_RIGHT });
                        doc.Close();
                        MessageBox.Show("PDF-файл создан! Откройте файл для предпросмотра.");
                        System.Diagnostics.Process.Start(sfd.FileName);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте: {ex.Message}\n\nStack Trace: {ex.StackTrace}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void btnExportWriteOffHistory_Click(object sender, EventArgs e)
        {
            try
            {
                using (SaveFileDialog sfd = new SaveFileDialog())
                {
                    sfd.Filter = "PDF files (*.pdf)|*.pdf";
                    sfd.FileName = "WriteOffHistory.pdf";
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        Document doc = new Document(PageSize.A4.Rotate(), 30, 30, 30, 30);
                        PdfWriter writer = PdfWriter.GetInstance(doc, new FileStream(sfd.FileName, FileMode.Create));
                        doc.Open();

                        BaseFont baseFont = BaseFont.CreateFont("c:\\windows\\fonts\\arial.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                        iTextSharp.text.Font font = new iTextSharp.text.Font(baseFont, 10);
                        iTextSharp.text.Font headerFont = new iTextSharp.text.Font(baseFont, 14, iTextSharp.text.Font.BOLD);

                        doc.Add(new Paragraph("История списаний", headerFont) { Alignment = Element.ALIGN_CENTER });
                        doc.Add(new Paragraph($"Дата формирования: {DateTime.Now.ToString("dd.MM.yyyy")}", font) { Alignment = Element.ALIGN_CENTER });
                        doc.Add(new Paragraph(" ", font));

                        PdfPTable table = new PdfPTable(9);
                        float[] widths = new float[] { 3f, 1f, 1f, 2f, 3f, 2f, 2f, 2f, 3f };
                        table.SetWidths(widths);
                        table.WidthPercentage = 100;

                        table.AddCell(new PdfPCell(new Phrase("Наименование", font)) { BackgroundColor = BaseColor.LIGHT_GRAY, HorizontalAlignment = Element.ALIGN_CENTER });
                        table.AddCell(new PdfPCell(new Phrase("Ед.изм", font)) { BackgroundColor = BaseColor.LIGHT_GRAY, HorizontalAlignment = Element.ALIGN_CENTER });
                        table.AddCell(new PdfPCell(new Phrase("Кол-во", font)) { BackgroundColor = BaseColor.LIGHT_GRAY, HorizontalAlignment = Element.ALIGN_CENTER });
                        table.AddCell(new PdfPCell(new Phrase("Отдел", font)) { BackgroundColor = BaseColor.LIGHT_GRAY, HorizontalAlignment = Element.ALIGN_CENTER });
                        table.AddCell(new PdfPCell(new Phrase("Устройство", font)) { BackgroundColor = BaseColor.LIGHT_GRAY, HorizontalAlignment = Element.ALIGN_CENTER });
                        table.AddCell(new PdfPCell(new Phrase("Инв.номер", font)) { BackgroundColor = BaseColor.LIGHT_GRAY, HorizontalAlignment = Element.ALIGN_CENTER });
                        table.AddCell(new PdfPCell(new Phrase("Причина", font)) { BackgroundColor = BaseColor.LIGHT_GRAY, HorizontalAlignment = Element.ALIGN_CENTER });
                        table.AddCell(new PdfPCell(new Phrase("Дата", font)) { BackgroundColor = BaseColor.LIGHT_GRAY, HorizontalAlignment = Element.ALIGN_CENTER });
                        table.AddCell(new PdfPCell(new Phrase("Примечание", font)) { BackgroundColor = BaseColor.LIGHT_GRAY, HorizontalAlignment = Element.ALIGN_CENTER });

                        double totalQuantity = 0;

                        foreach (DataRow row in writeOffDataTable.Rows)
                        {
                            table.AddCell(new Phrase(row["MaterialName"].ToString(), font));
                            table.AddCell(new Phrase(row["Unit"].ToString(), font));
                            table.AddCell(new Phrase(Convert.ToDouble(row["Quantity"]).ToString("F2"), font));
                            table.AddCell(new Phrase(row["Department"].ToString(), font));
                            table.AddCell(new Phrase(row["DeviceName"].ToString(), font));
                            table.AddCell(new Phrase(row["InventoryNumber"].ToString(), font));
                            table.AddCell(new Phrase(row["Reason"].ToString(), font));
                            table.AddCell(new Phrase(row["Date"].ToString(), font));
                            table.AddCell(new Phrase(row["Note"].ToString(), font));
                            totalQuantity += Convert.ToDouble(row["Quantity"]);
                        }

                        doc.Add(table);
                        doc.Add(new Paragraph($"Итого списано: {totalQuantity:F2}", font) { Alignment = Element.ALIGN_RIGHT });

                        doc.Close();
                        MessageBox.Show("PDF-файл создан! Откройте файл для предпросмотра.");
                        System.Diagnostics.Process.Start(sfd.FileName);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void btnResetWriteOffSearch_Click(object sender, EventArgs e)
        {
            txtWriteOffSearch.Clear();
            LoadWriteOffData();
        }

        private void btnExportTurnoverByMonth_Click(object sender, EventArgs e)
        {
            try
            {
                if (cmbMonth.SelectedIndex == -1)
                {
                    MessageBox.Show("Выберите месяц!");
                    return;
                }

                using (SaveFileDialog sfd = new SaveFileDialog())
                {
                    sfd.Filter = "PDF files (*.pdf)|*.pdf";
                    sfd.FileName = $"TurnoverByMonthReport_{cmbMonth.SelectedItem}.pdf";
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        Document doc = new Document(PageSize.A4, 30, 30, 30, 30);
                        PdfWriter writer = PdfWriter.GetInstance(doc, new FileStream(sfd.FileName, FileMode.Create));
                        doc.Open();

                        BaseFont baseFont = BaseFont.CreateFont("c:\\windows\\fonts\\arial.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                        iTextSharp.text.Font font = new iTextSharp.text.Font(baseFont, 12);
                        iTextSharp.text.Font headerFont = new iTextSharp.text.Font(baseFont, 14, iTextSharp.text.Font.BOLD);

                        int selectedMonthIndex = cmbMonth.SelectedIndex + 1;
                        string monthName = cmbMonth.SelectedItem.ToString();

                        doc.Add(new Paragraph($"Оборотная ведомость по НФА за {monthName}", headerFont) { Alignment = Element.ALIGN_CENTER });
                        doc.Add(new Paragraph($"Дата формирования: {DateTime.Now.ToString("dd.MM.yyyy")}", font) { Alignment = Element.ALIGN_CENTER });
                        doc.Add(new Paragraph(" ", font));

                        using (SQLiteConnection conn = new SQLiteConnection(connectionString))
                        {
                            conn.Open();
                            string query = @"SELECT MaterialName, Unit, Price, QuantityEnd, TotalEnd
                    FROM Turnover
                    WHERE substr(Date, 4, 2) = @Month";

                            using (SQLiteDataAdapter adapter = new SQLiteDataAdapter(query, conn))
                            {
                                adapter.SelectCommand.Parameters.AddWithValue("@Month", selectedMonthIndex.ToString("00"));
                                DataTable monthlyDataTable = new DataTable();
                                adapter.Fill(monthlyDataTable);

                                if (monthlyDataTable.Rows.Count == 0)
                                {
                                    doc.Add(new Paragraph($"Нет данных за {monthName}.", font));
                                }
                                else
                                {
                                    PdfPTable table = new PdfPTable(5);
                                    table.WidthPercentage = 100;
                                    float[] widths = new float[] { 3f, 1f, 1f, 2f, 2f };
                                    table.SetWidths(widths);

                                    table.AddCell(new PdfPCell(new Phrase("Наименование", font)) { BackgroundColor = BaseColor.LIGHT_GRAY, HorizontalAlignment = Element.ALIGN_CENTER });
                                    table.AddCell(new PdfPCell(new Phrase("Ед.изм", font)) { BackgroundColor = BaseColor.LIGHT_GRAY, HorizontalAlignment = Element.ALIGN_CENTER });
                                    table.AddCell(new PdfPCell(new Phrase("Цена", font)) { BackgroundColor = BaseColor.LIGHT_GRAY, HorizontalAlignment = Element.ALIGN_CENTER });
                                    table.AddCell(new PdfPCell(new Phrase("Остаток (Кол-во)", font)) { BackgroundColor = BaseColor.LIGHT_GRAY, HorizontalAlignment = Element.ALIGN_CENTER });
                                    table.AddCell(new PdfPCell(new Phrase("Остаток (Сумма)", font)) { BackgroundColor = BaseColor.LIGHT_GRAY, HorizontalAlignment = Element.ALIGN_CENTER });

                                    double totalQuantity = 0;
                                    double totalSum = 0;

                                    foreach (DataRow row in monthlyDataTable.Rows)
                                    {
                                        table.AddCell(new Phrase(row["MaterialName"].ToString(), font));
                                        table.AddCell(new Phrase(row["Unit"].ToString(), font));
                                        table.AddCell(new Phrase(Convert.ToDouble(row["Price"]).ToString("F2"), font));
                                        table.AddCell(new Phrase(Convert.ToDouble(row["QuantityEnd"]).ToString("F2"), font));
                                        table.AddCell(new Phrase(Convert.ToDouble(row["TotalEnd"]).ToString("F2"), font));
                                        totalQuantity += Convert.ToDouble(row["QuantityEnd"]);
                                        totalSum += Convert.ToDouble(row["TotalEnd"]);
                                    }

                                    doc.Add(table);
                                    doc.Add(new Paragraph($"Итого Кол-во: {totalQuantity:F2}, Итого Сумма: {totalSum:F2}", font) { Alignment = Element.ALIGN_RIGHT });
                                }
                            }
                            conn.Close();
                        }

                        doc.Close();
                        MessageBox.Show("PDF-файл создан! Откройте файл для предпросмотра.");
                        System.Diagnostics.Process.Start(sfd.FileName);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}