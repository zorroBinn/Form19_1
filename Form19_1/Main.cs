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
        public Main()
        {
            InitializeComponent();
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
                // Проверяем, выбрана ли ячейка в первом столбце
                if (dataGridView.CurrentCell.ColumnIndex == 0)
                {
                    DeleteSelectedRow();
                }
            }

        }
    }
}
