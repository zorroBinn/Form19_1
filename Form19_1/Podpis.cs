using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox;

namespace Form19_1
{
    public partial class Podpis : Form
    {
        private Main parentForm;
        private string dolzhnostFile = "dolzhnostBase.csv";
        private string podpisFile = "00p11o22d33p44i55s66.csv";

        public Podpis(Main parent)
        {
            InitializeComponent();
            this.parentForm = parent;
            LoadDolzhnosti();
            LoadPodpisData();
        }

        private void LoadPodpisData()
        {
            if (!File.Exists(podpisFile))
            {
                //Создаём файл с пустыми строками при первом запуске
                File.WriteAllLines(podpisFile, new string[] { "", "", "", "" });
            }

            string[] lines = File.ReadAllLines(podpisFile);
            if (lines.Length >= 4)
            {
                comboBox_dolg1.Text = lines[0].Trim();
                textBox_fio1.Text = lines[1].Trim();
                comboBox_dolg2.Text = lines[2].Trim();
                textBox_fio2.Text = lines[3].Trim();
            }
        }

        private void LoadDolzhnosti()
        {
            if (File.Exists(dolzhnostFile))
            {
                var dolzhnosti = File.ReadAllLines(dolzhnostFile).Distinct().ToList();
                comboBox_dolg1.Items.Clear();
                comboBox_dolg2.Items.Clear();
                comboBox_dolg1.Items.AddRange(dolzhnosti.ToArray());
                comboBox_dolg2.Items.AddRange(dolzhnosti.ToArray());
            }
        }

        private void SaveDolzhnost(string dolzhnost)
        {
            if (!string.IsNullOrWhiteSpace(dolzhnost))
            {
                var dolzhnosti = File.Exists(dolzhnostFile) ? File.ReadAllLines(dolzhnostFile).ToList() : new List<string>();

                if (!dolzhnosti.Contains(dolzhnost))
                {
                    dolzhnosti.Add(dolzhnost);
                    File.WriteAllLines(dolzhnostFile, dolzhnosti);
                }
            }
        }

        private bool HighlightIfEmpty(Control control)
        {
            if (control is System.Windows.Forms.ComboBox comboBox && string.IsNullOrWhiteSpace(comboBox.Text) ||
                control is System.Windows.Forms.TextBox textBox && string.IsNullOrWhiteSpace(textBox.Text))
            {
                control.BackColor = Color.LightCoral;
                Timer timer = new Timer { Interval = 500 };
                timer.Tick += (s, e) =>
                {
                    control.BackColor = Color.White;
                    timer.Stop();
                    timer.Dispose();
                };
                timer.Start();
                return false;
            }
            return true;
        }

        private void button_zapomn_Click(object sender, EventArgs e)
        {
            bool allFilled = true;
            allFilled &= HighlightIfEmpty(comboBox_dolg1);
            allFilled &= HighlightIfEmpty(comboBox_dolg2);
            allFilled &= HighlightIfEmpty(textBox_fio1);
            allFilled &= HighlightIfEmpty(textBox_fio2);

            if (allFilled)
            {
                SaveDolzhnost(comboBox_dolg1.Text);
                SaveDolzhnost(comboBox_dolg2.Text);
                parentForm.UpdateButtonColor();
                this.Close();
                {
                    string[] lines =
                    {
                        comboBox_dolg1.Text.Trim(),
                        textBox_fio1.Text.Trim(),
                        comboBox_dolg2.Text.Trim(),
                        textBox_fio2.Text.Trim()
                    };

                    File.WriteAllLines(podpisFile, lines);
                }
            }
        }

        private void textBox_fio1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsLetter(e.KeyChar) && e.KeyChar != ' ' && e.KeyChar != (char)Keys.Back && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true;
            }
        }

        private void textBox_fio2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsLetter(e.KeyChar) && e.KeyChar != ' ' && e.KeyChar != (char)Keys.Back && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true;
            }
        }
    }
}
