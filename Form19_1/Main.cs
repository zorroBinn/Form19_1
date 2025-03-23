using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Globalization;

namespace Form19_1
{
    public partial class Main : Form
    {
        private int currentSet = 1;
        private Button[] buttons;
        private System.Windows.Forms.Label[] sumLabels;
        private System.Windows.Forms.Label[] nomberLabels;
        private string organizationBaseFilePath = "organizationBase.csv";
        private Dictionary<string, List<string>> organizationData = new Dictionary<string, List<string>>();
        private Dictionary<int, string> workersBySet = new Dictionary<int, string>();

        public Main()
        {
            InitializeComponent();
            this.dateTimePicker_PO.MinDate = this.dateTimePicker_S.Value;
            this.dateTimePicker.MinDate = this.dateTimePicker_S.Value;
            buttons = new Button[] { button1, button2, button3, button4 };
            sumLabels = new System.Windows.Forms.Label[] { label_sum6, label_sum7, label_sum8, label_sum9 };
            nomberLabels = new System.Windows.Forms.Label[] { label_1, label_2, label_3, label_4, label_5, label_6, label_7, label_8, label_9 };
            LoadOrganizations();
        }

        ///<summary>
        ///Автоинкрементирование номера строк
        ///</summary>
        private void UpdateRowNumbers()
        {
            for (int i = 0; i < dataGridView.Rows.Count; i++)
            {
                dataGridView.Rows[i].Cells["Column1"].Value = i + 1;
            }
        }

        ///<summary>
        ///Загрузка организаций и подразделений
        ///</summary>
        private void LoadOrganizations()
        {
            if (!File.Exists(organizationBaseFilePath)) return;

            string[] lines = File.ReadAllLines(organizationBaseFilePath);
            foreach (string line in lines)
            {
                string[] parts = line.Split(';');
                if (parts.Length > 1)
                {
                    string organization = parts[0];
                    List<string> departments = parts.Skip(1).ToList();
                    organizationData[organization] = departments;
                }
                else if (parts.Length == 1)
                {
                    //Если нет подразделений, создаём пустой список
                    organizationData[parts[0]] = new List<string>();
                }
            }

            comboBox_organiz.Items.AddRange(organizationData.Keys.ToArray());
        }

        private void Main_Load(object sender, EventArgs e)
        {
            UpdateRowNumbers();
        }

        private void dataGridView_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            UpdateRowNumbers();
            UpdateSums();
        }

        private void dataGridView_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            UpdateRowNumbers();
            UpdateSums();
        }

        ///<summary>
        ///Удаление выбранной строки
        ///</summary>
        private void DeleteSelectedRow()
        {
            if (dataGridView.CurrentRow != null && !dataGridView.CurrentRow.IsNewRow)
            {
                dataGridView.Rows.Remove(dataGridView.CurrentRow);
            }
        }

        private void dataGridView_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete && dataGridView.CurrentCell != null)
            {
                if (dataGridView.CurrentCell.ColumnIndex == 0)
                {
                    DeleteSelectedRow();
                }
            }

        }

        private void linkLabel_Utverd_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            using (Podpis child = new Podpis(this))
            {
                this.Enabled = false;
                child.ShowDialog();
                this.Enabled = true;
            }
        }

        public void UpdateButtonColor()
        {
            linkLabel_Utverd.LinkColor = Color.Green;
        }

        ///<summary>
        ///При изменении размеров окна размеры столбцов делать равным размеру лейблов-заголовков
        ///</summary>
        private void ResizeColumnsToLabels()
        {
            for (int i = 0; i < 9; i++)
            {
                string baseColName = $"Column{i + 1}";

                //Для первых 5 столбцов
                if (i < 5)
                {
                    if (dataGridView.Columns.Contains(baseColName))
                    {
                        int width = nomberLabels[i].Width;
                        if (i == 0) width -= 2; //Делаем первый столбец уже на 2 пикселя
                        dataGridView.Columns[baseColName].Width = width;
                    }
                }
                else //Для столбцов 6-9
                {
                    for (int j = 1; j <= 4; j++) //4 скрытых копии у каждого
                    {
                        string colName = $"{baseColName}_{j}";
                        if (dataGridView.Columns.Contains(colName))
                            dataGridView.Columns[colName].Width = nomberLabels[i].Width;
                    }
                }
            }
        }

        private void Main_Resize(object sender, EventArgs e)
        {
            ResizeColumnsToLabels();
        }

        private void dateTimePicker_S_ValueChanged(object sender, EventArgs e)
        {
            this.dateTimePicker_PO.MinDate = this.dateTimePicker_S.Value;
            this.dateTimePicker_PO.MaxDate = new DateTime(this.dateTimePicker_S.Value.Year, 12, 31);
            this.dateTimePicker.MinDate = this.dateTimePicker_S.Value;
        }

        ///<summary>
        ///Функция проверки, заполнена ли первая строка таблицы
        ///</summary>
        ///<param name="dgv"> Таблица для проверки </param>
        ///<returns></returns>
        private bool IsFirstRowEmpty(DataGridView dgv)
        {
            if (dgv.Rows.Count == 0) return true;
            DataGridViewRow firstRow = dgv.Rows[0];

            //Проверяем, что столбцы 1-5 заполнены
            for (int col = 0; col < 5; col++)
            {
                DataGridViewCell cell = firstRow.Cells[col];

                if (cell.Value == null || string.IsNullOrWhiteSpace(cell.Value.ToString()))
                    return true;
            }

            for (int set = 1; set <= 4; set++)
            {
                bool isSetFilled = true;

                for (int col = 6; col <= 9; col++)
                {
                    string columnName = $"column{col}_{set}";

                    if (dgv.Columns.Contains(columnName))
                    {
                        DataGridViewCell cell = firstRow.Cells[columnName];

                        if (cell.Value == null || string.IsNullOrWhiteSpace(cell.Value.ToString()))
                        {
                            isSetFilled = false;
                            break;
                        }
                    }
                }

                if (isSetFilled) return false; //Если хотя бы один сет заполнен, строка НЕ пустая
            }

            return true;
        }

        ///<summary>
        ///Выделение незаполненного элемента
        ///</summary>
        ///<param name="control"> Элемент </param>
        ///<returns></returns>
        private bool HighlightIfEmpty(Control control)
        {
            Dictionary<Control, Color> originalColors = new Dictionary<Control, Color>();

            if ((control is ComboBox comboBox && string.IsNullOrWhiteSpace(comboBox.Text)) ||
                (control is TextBox textBox && string.IsNullOrWhiteSpace(textBox.Text)) ||
                (control is DataGridView dgv && IsFirstRowEmpty(dgv)) ||
                (control is LinkLabel linkLabel && linkLabel.LinkColor != Color.Green))
            {
                //Запоминаем оригинальный цвет
                if (!originalColors.ContainsKey(control))
                    originalColors[control] = control.BackColor;

                control.BackColor = Color.LightCoral;

                if (control is DataGridView dgv0)
                {
                    if (!originalColors.ContainsKey(dgv0))
                        originalColors[dgv0] = dgv0.BackgroundColor;
                    dgv0.BackgroundColor = Color.LightCoral;
                    MessageBox.Show("Таблица заполнена некорректно!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                Timer timer = new Timer { Interval = 500 };
                timer.Tick += (s, e) =>
                {
                    if (originalColors.ContainsKey(control))
                        control.BackColor = originalColors[control];

                    if (control is DataGridView dgv1 && originalColors.ContainsKey(dgv1))
                        dgv1.BackgroundColor = originalColors[dgv1];

                    timer.Stop();
                    timer.Dispose();
                };
                timer.Start();
                return false;
            }
            return true;
        }

        ///<summary>
        ///Запоминание вводимых организаций/подразделений
        ///</summary>
        private void SaveOrganizations()
        {
            List<string> lines = new List<string>();

            foreach (var entry in organizationData)
            {
                string line = entry.Key + ";" + string.Join(";", entry.Value);
                lines.Add(line);
            }

            File.WriteAllLines(organizationBaseFilePath, lines);
        }

        private void button_form_Click(object sender, EventArgs e)
        {
            string org = comboBox_organiz.Text.Trim();
            string dep = comboBox_podrazdel.Text.Trim();
            bool allFilled = true;
            allFilled &= HighlightIfEmpty(comboBox_organiz);
            allFilled &= HighlightIfEmpty(comboBox_podrazdel);
            allFilled &= HighlightIfEmpty(textBox_OKPO);
            allFilled &= HighlightIfEmpty(textBox_OKDP);
            allFilled &= HighlightIfEmpty(linkLabel_Utverd);
            allFilled &= HighlightIfEmpty(dataGridView);

            if (allFilled)
            {
                //Экспорт
                {
                    //Добавляем организацию, если её нет
                    if (!organizationData.ContainsKey(org))
                    {
                        organizationData[org] = new List<string>();
                        comboBox_organiz.Items.Add(org);
                    }
                    //Добавляем подразделение, если его ещё нет
                    if (!string.IsNullOrWhiteSpace(dep) && !organizationData[org].Contains(dep))
                    {
                        organizationData[org].Add(dep);
                        comboBox_podrazdel.Items.Add(dep);
                    }
                    SaveOrganizations();
                }
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Excel Files|*.xls";
                    saveFileDialog.Title = "Сохранить файл";
                    saveFileDialog.FileName = $"Form_ОП-19_{dateTimePicker.Value:yyyy-MM-dd}.xls";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        string targetPath = saveFileDialog.FileName;
                        string templatePath = "obrazec_FormN_OP-19.XLS";

                        try
                        {
                            File.Copy(templatePath, targetPath, true); //Копируем шаблон
                            FillExcelFile(targetPath); //Заполняем файл
                            MessageBox.Show("Экспорт выполнен успешно!", "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            this.linkLabel_NewForm.Visible = true;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
        }

        ///<summary>
        ///Экспорт в Excel файл (заполнение по шаблону)
        ///</summary>
        ///<param name="filePath"></param>
        private void FillExcelFile(string filePath)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);
            Excel.Worksheet worksheet = workbook.Sheets[1]; //Первый лист

            try
            {
                //Функция для получения названия месяца в родительном падеже
                string GetMonthName(DateTime date)
                {
                    return date.ToString("MMMM", new CultureInfo("ru-RU"));
                }

                //Заполняем ячейки по заданным координатам
                worksheet.Cells[6, 1] = comboBox_organiz.Text.Trim(); //A6
                worksheet.Cells[8, 1] = comboBox_podrazdel.Text.Trim(); //A8
                worksheet.Cells[6, 69] = textBox_OKPO.Text.Trim(); //BQ6
                worksheet.Cells[9, 69] = textBox_OKDP.Text.Trim(); //BQ9

                worksheet.Cells[18, 25] = dateTimePicker_S.Value.Day + " " + GetMonthName(dateTimePicker_S.Value); //Y18 (день месяц)
                worksheet.Cells[18, 39] = dateTimePicker_PO.Value.Day + " " + GetMonthName(dateTimePicker_PO.Value); //AM18 (день месяц)
                worksheet.Cells[18, 52] = dateTimePicker_S.Value.Year; //AZ18 (год)

                worksheet.Cells[35, 57] = dateTimePicker.Value.Day; //BE35 (день)
                worksheet.Cells[35, 59] = GetMonthName(dateTimePicker.Value); //BG35 (месяц)
                worksheet.Cells[35, 70] = dateTimePicker.Value.Year; //BR35 (год)

                //Читаем данные из файла подписей
                string podpisFile = "00p11o22d33p44i55s66.csv";
                if (File.Exists(podpisFile))
                {
                    string[] podpisLines = File.ReadAllLines(podpisFile);
                    if (podpisLines.Length >= 4)
                    {
                        worksheet.Cells[23, 25] = podpisLines[0].Trim(); //Y23
                        worksheet.Cells[23, 37] = podpisLines[1].Trim(); //AK23
                        worksheet.Cells[67, 19] = podpisLines[2].Trim(); //S67
                        worksheet.Cells[67, 45] = podpisLines[3].Trim(); //AS67
                    }
                }

                //Заполняем таблицу
                Dictionary<int, string> workerCells = new Dictionary<int, string>
                {
                    { 1, "Y39" }, { 2, "AL39" }, { 3, "AY39" }, { 4, "BL39" }
                };

                foreach (var set in workersBySet)
                {
                    if (workerCells.ContainsKey(set.Key))
                    {
                        worksheet.Range[workerCells[set.Key]].Value = set.Value.Trim();
                    }
                }

                int maxRows = Math.Min(20, dataGridView.Rows.Count);
                int excelRowStart = 44; // Начальная строка в Excel
                List<string> orderedColumns = new List<string>();
                // Добавляем сначала первые 5 столбцов (Column1 - Column5)
                for (int i = 1; i <= 5; i++)
                {
                    orderedColumns.Add($"Column{i}");
                }
                // Затем добавляем столбцы по сетам (Column6_1 - Column9_4)
                for (int set = 1; set <= 4; set++)
                {
                    for (int col = 6; col <= 9; col++)
                    {
                        orderedColumns.Add($"Column{col}_{set}");
                    }
                }
                // Сопоставление с колонками Excel
                Dictionary<string, string> columnMap = new Dictionary<string, string>
                {
                    { "Column1", "A" }, { "Column2", "E" }, { "Column3", "O" }, { "Column4", "R" }, { "Column5", "U" },
                    { "Column6_1", "Y" }, { "Column7_1", "AB" }, { "Column8_1", "AF" }, { "Column9_1", "AI" },
                    { "Column6_2", "AL" }, { "Column7_2", "AO" }, { "Column8_2", "AS" }, { "Column9_2", "AV" },
                    { "Column6_3", "AY" }, { "Column7_3", "BB" }, { "Column8_3", "BF" }, { "Column9_3", "BI" },
                    { "Column6_4", "BL" }, { "Column7_4", "BO" }, { "Column8_4", "BS" }, { "Column9_4", "BV" }
                };
                Dictionary<string, double> columnSums = new Dictionary<string, double>(); //Словарь для хранения сумм по каждому столбцу
                for (int rowIndex = 0; rowIndex < maxRows; rowIndex++)
                {
                    int excelRow = excelRowStart + rowIndex; // Excel строки 44-63

                    foreach (string columnName in orderedColumns)
                    {
                        if (!dataGridView.Columns.Contains(columnName)) continue;
                        int colIndex = dataGridView.Columns[columnName].Index;
                        if (dataGridView.Rows[rowIndex].Cells[colIndex].Value != null)
                        {
                            string excelColumn = columnMap[columnName];
                            worksheet.Cells[excelRow, excelColumn] = dataGridView.Rows[rowIndex].Cells[colIndex].Value.ToString();
                        }
                    }
                }

                // Подсчитываем суммы по каждому столбцу сетов
                foreach (int set in new int[] { 1, 2, 3, 4 }) // Перебираем 4 сета
                {
                    bool isSetEmpty = true; // Флаг, пуст ли сет полностью
                    Dictionary<string, double> tempSums = new Dictionary<string, double>(); // Временное хранилище сумм столбцов сета

                    for (int col = 6; col <= 9; col++) // Перебираем столбцы Column6_X - Column9_X
                    {
                        string columnName = $"Column{col}_{set}";

                        if (!dataGridView.Columns.Contains(columnName)) continue;

                        int colIndex = dataGridView.Columns[columnName].Index;
                        double sum = 0;

                        for (int rowIndex = 0; rowIndex < maxRows; rowIndex++)
                        {
                            if (dataGridView.Rows[rowIndex].Cells[colIndex].Value != null &&
                                double.TryParse(dataGridView.Rows[rowIndex].Cells[colIndex].Value.ToString(), out double value))
                            {
                                sum += value;
                                isSetEmpty = false; // Нашли хотя бы одно число — сет не пустой
                            }
                        }

                        tempSums[columnName] = sum;
                    }

                    // Если сет не пустой, записываем его суммы
                    if (!isSetEmpty)
                    {
                        foreach (var kvp in tempSums)
                        {
                            if (columnMap.ContainsKey(kvp.Key))
                            {
                                string excelColumn = columnMap[kvp.Key];
                                worksheet.Cells[64, excelColumn] = kvp.Value;
                            }
                        }
                    }
                }

                workbook.Save();
                workbook.Close();
                excelApp.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при заполнении файла: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                ReleaseObject(worksheet);
                ReleaseObject(workbook);
                ReleaseObject(excelApp);
            }
        }

        ///<summary>
        ///Освобождение ресурсов Excel
        ///</summary>
        ///<param name="obj"> Освобождаемый объект </param>
        private void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch { obj = null; }
            finally { GC.Collect(); }
        }

        ///<summary>
        ///Изменение сета работников
        ///</summary>
        ///<param name="set"> Номер сета </param>
        private void SwitchColumnSet(int set)
        {
            if (set < 1 || set > 4) return;

            currentSet = set;

            for (int i = 6; i <= 9; i++)
            {
                for (int j = 1; j <= 4; j++)
                {
                    string colName = $"Column{i}_{j}";
                    if (dataGridView.Columns.Contains(colName))
                        dataGridView.Columns[colName].Visible = (j == set);
                }
            }

            //Обновляем стиль кнопок
            foreach (var btn in buttons)
            {
                btn.Font = new Font(btn.Font, FontStyle.Regular);
                btn.BackColor = SystemColors.Control;
            }

            buttons[set - 1].Font = new Font(buttons[set - 1].Font, FontStyle.Bold);
            buttons[set - 1].BackColor = Color.Silver;

            //Обновляем суммы
            UpdateSums();
            //Обновляем текст в textBox_rab
            textBox_rab.Text = workersBySet.ContainsKey(set) ? workersBySet[set] : "";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SwitchColumnSet(1);
            if(this.Size != this.MinimumSize) ResizeColumnsToLabels();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SwitchColumnSet(2);
            if (this.Size != this.MinimumSize) ResizeColumnsToLabels();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SwitchColumnSet(3);
            if (this.Size != this.MinimumSize) ResizeColumnsToLabels();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            SwitchColumnSet(4);
            if (this.Size != this.MinimumSize) ResizeColumnsToLabels();
        }

        private void button_prev_Click(object sender, EventArgs e)
        {
            SwitchColumnSet(Math.Max(1, currentSet - 1));
            if (this.Size != this.MinimumSize) ResizeColumnsToLabels();
        }

        private void button_next_Click(object sender, EventArgs e)
        {
            SwitchColumnSet(Math.Min(4, currentSet + 1));
            if (this.Size != this.MinimumSize) ResizeColumnsToLabels();
        }

        private void dataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex >= 0 && e.RowIndex >= 0)
                UpdateSums();
        }

        ///<summary>
        ///Сводка по количеству (сумма)
        ///</summary>
        private void UpdateSums()
        {
            for (int i = 6; i <= 9; i++) //Столбцы 6-9
            {
                string colName = $"Column{i}_{currentSet}";
                if (!dataGridView.Columns.Contains(colName)) continue;

                decimal sum = 0;
                bool hasValues = false;

                foreach (DataGridViewRow row in dataGridView.Rows)
                {
                    if (row.IsNewRow) continue; //Пропускаем пустую строку ввода

                    if (decimal.TryParse(row.Cells[colName].Value?.ToString(), out decimal value))
                    {
                        sum += value;
                        hasValues = true;
                    }
                }

                sumLabels[i - 6].Text = hasValues ? sum.ToString("N0") : ""; //Если нет значений - очищаем текст
            }
        }

        private void textBox_OKPO_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && e.KeyChar != (char)Keys.Back && e.KeyChar != (char)Keys.Delete)
            {
                e.Handled = true;
            }
        }

        private void textBox_OKDP_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && e.KeyChar != ' ' && e.KeyChar != '.' && e.KeyChar != (char)Keys.Back && e.KeyChar != (char)Keys.Delete)
            {
                e.Handled = true;
            }
        }

        private void textBox_rab_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsLetter(e.KeyChar) && e.KeyChar != ' ' && e.KeyChar != (char)Keys.Back && e.KeyChar != (char)Keys.Delete)
            {
                e.Handled = true;
            }
        }

        private void comboBox_organiz_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && !char.IsPunctuation(e.KeyChar) && !char.IsLetter(e.KeyChar) && e.KeyChar != ' ' && e.KeyChar != (char)Keys.Back && e.KeyChar != (char)Keys.Delete)
            {
                e.Handled = true;
            }
        }

        private void comboBox_podrazdel_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && !char.IsPunctuation(e.KeyChar) && !char.IsLetter(e.KeyChar) && e.KeyChar != ' ' && e.KeyChar != (char)Keys.Back && e.KeyChar != (char)Keys.Delete)
            {
                e.Handled = true;
            }
        }

        private void dataGridView_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (e.Control is TextBox textBox)
            {
                textBox.KeyPress -= OnlyNumbers_KeyPress;
                textBox.KeyPress -= textBox_OKDP_KeyPress;

                string colName = dataGridView.CurrentCell.OwningColumn.Name;

                if (colName.StartsWith("Column6_") || colName.StartsWith("Column7_") ||
                    colName.StartsWith("Column8_") || colName.StartsWith("Column9_"))
                {
                    textBox.KeyPress += OnlyNumbers_KeyPress;
                }
                if (colName == "Column3" || colName == "Column5")
                    textBox.KeyPress += textBox_OKDP_KeyPress;
            }
        }

        private void OnlyNumbers_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && e.KeyChar != (char)Keys.Back && e.KeyChar != (char)Keys.Delete)
            {
                e.Handled = true;
            }
        }

        private void comboBox_organiz_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox_podrazdel.Items.Clear();
            comboBox_podrazdel.Text = string.Empty;

            if (comboBox_organiz.SelectedItem is string selectedOrg && organizationData.ContainsKey(selectedOrg))
            {
                comboBox_podrazdel.Items.AddRange(organizationData[selectedOrg].ToArray());
            }
        }

        private void Main_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (File.Exists("00p11o22d33p44i55s66.csv"))
            {
                File.Delete("00p11o22d33p44i55s66.csv");
            }
        }

        private void linkLabel_NewForm_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            DialogResult result = MessageBox.Show("Вы уверены, что хотите очистить форму и перейти к новой?",
                                          "Подтверждение",
                                          MessageBoxButtons.YesNo,
                                          MessageBoxIcon.Warning);

            if (result == DialogResult.Yes)
            {
                foreach (Control ctrl in this.Controls)
                {
                    if (ctrl is TextBox textBox)
                        textBox.Clear();
                }
                this.textBox_rab.Clear();
                foreach (Control ctrl in this.Controls)
                {
                    if (ctrl is ComboBox comboBox)
                        comboBox.SelectedIndex = -1;
                }
                dataGridView.Rows.Clear();
                foreach (Control ctrl in this.Controls)
                {
                    if (ctrl is DateTimePicker dateTimePicker)
                        dateTimePicker.Value = dateTimePicker.MinDate;
                }
            }
            this.linkLabel_NewForm.Visible = false;
            if (File.Exists("00p11o22d33p44i55s66.csv"))
            {
                File.Delete("00p11o22d33p44i55s66.csv");
            }
        }

        private void textBox_rab_TextChanged(object sender, EventArgs e)
        {
            workersBySet[currentSet] = textBox_rab.Text.Trim();
        }
    }
}
