﻿using System;
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

        public Podpis(Main parent)
        {
            InitializeComponent();
            this.parentForm = parent;
            LoadDolzhnosti();
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
