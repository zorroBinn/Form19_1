using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Form19_1
{
    public partial class Main : Form
    {
        private int currentSet = 1;
        private Button[] buttons;

        public Main()
        {
            InitializeComponent();
            this.dateTimePicker_PO.MinDate = this.dateTimePicker_S.Value;
            this.dateTimePicker.MinDate = this.dateTimePicker_S.Value;
            buttons = new Button[] { button1, button2, button3, button4 };
        }

        private void UpdateRowNumbers()
        {
            for (int i = 0; i < dataGridView.Rows.Count; i++)
            {
                dataGridView.Rows[i].Cells["Column1"].Value = i + 1;
            }
        }

        private void Main_Load(object sender, EventArgs e)
        {
            UpdateRowNumbers();
        }

        private void dataGridView_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            UpdateRowNumbers();
        }

        private void dataGridView_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            UpdateRowNumbers();
        }

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

        private void linkLabel_New_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
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
            linkLabel_New.LinkColor = Color.Green;
        }

        private void Main_Resize(object sender, EventArgs e)
        {

        }

        private void dateTimePicker_S_ValueChanged(object sender, EventArgs e)
        {
            this.dateTimePicker_PO.MinDate = this.dateTimePicker_S.Value;
            this.dateTimePicker.MinDate = this.dateTimePicker_S.Value;
        }

        //Функция проверки, заполнена ли первая строка таблицы
        private bool IsFirstRowEmpty(DataGridView dgv)
        {
            if (dgv.Rows.Count == 0) return true;
            DataGridViewRow firstRow = dgv.Rows[0];
            foreach (DataGridViewCell cell in firstRow.Cells)
            {
                if (cell.Value == null || string.IsNullOrWhiteSpace(cell.Value.ToString()))
                    return true;
            }
            return false;
        }

        private bool HighlightIfEmpty(Control control)
        {
            Dictionary<Control, Color> originalColors = new Dictionary<Control, Color>();

            if ((control is ComboBox comboBox && string.IsNullOrWhiteSpace(comboBox.Text)) ||
                (control is TextBox textBox && string.IsNullOrWhiteSpace(textBox.Text)) ||
                (control is DataGridView dgv && IsFirstRowEmpty(dgv)) ||
                (control is LinkLabel linkLabel && linkLabel.LinkColor != Color.Green))
            {
                // Запоминаем оригинальный цвет
                if (!originalColors.ContainsKey(control))
                    originalColors[control] = control.BackColor;

                control.BackColor = Color.LightCoral;

                if (control is DataGridView dgv0)
                {
                    if (!originalColors.ContainsKey(dgv0))
                        originalColors[dgv0] = dgv0.BackgroundColor;
                    dgv0.BackgroundColor = Color.LightCoral;
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

        private void button_form_Click(object sender, EventArgs e)
        {
            bool allFilled = true;
            allFilled &= HighlightIfEmpty(comboBox_organiz);
            allFilled &= HighlightIfEmpty(comboBox_podrazdel);
            allFilled &= HighlightIfEmpty(textBox_OKPO);
            allFilled &= HighlightIfEmpty(textBox_OKDP);
            allFilled &= HighlightIfEmpty(linkLabel_New);
            allFilled &= HighlightIfEmpty(dataGridView);

            if (allFilled)
            {
                //Экспорт
            }
        }

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
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SwitchColumnSet(1);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SwitchColumnSet(2);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SwitchColumnSet(3);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            SwitchColumnSet(4);
        }

        private void button_prev_Click(object sender, EventArgs e)
        {
            SwitchColumnSet(Math.Max(1, currentSet - 1));
        }

        private void button_next_Click(object sender, EventArgs e)
        {
            SwitchColumnSet(Math.Min(4, currentSet + 1));
        }
    }
}
