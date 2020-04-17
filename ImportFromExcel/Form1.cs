using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ImportFromExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                string file = openFileDialog1.FileName;
                textBox1.Text = file;

                Excel dataFromExcel = new Excel();
                DataTable dt = dataFromExcel.LoadDataFromExcel(file);

                dataGridView1.DataSource = dt;
                Excel.dataTable = dt;
                textBox2.Enabled = true;
            }
            else
                MessageBox.Show("Not selected file");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DataTable dataTable = Excel.dataTable;

            var testData = dataTable.Select().Where(x => x.ItemArray[0].ToString() == textBox2.Text && x.ItemArray[1].ToString() == textBox3.Text).SingleOrDefault();
            if (testData != null)
            {
                textBox4.Enabled = true;
                textBox4.Text = testData.ItemArray[2].ToString();
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            textBox3.Enabled = true;
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            button2.Enabled = true;
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            button3.Enabled = true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "XLS files (*.xls)|*.xls|XLT files (*.xlt)|*.xlt|XLSX files (*.xlsx)|*.xlsx|XLSM files (*.xlsm)|*.xlsm|XLTX (*.xltx)|*.xltx|XLTM (*.xltm)|*.xltm|ODS (*.ods)|*.ods|OTS (*.ots)|*.ots|CSV (*.csv)|*.csv|TSV (*.tsv)|*.tsv|HTML (*.html)|*.html|MHTML (.mhtml)|*.mhtml|PDF (*.pdf)|*.pdf|XPS (*.xps)|*.xps|BMP (*.bmp)|*.bmp|GIF (*.gif)|*.gif|JPEG (*.jpg)|*.jpg|PNG (*.png)|*.png|TIFF (*.tif)|*.tif|WMP (*.wdp)|*.wdp";
            saveFileDialog.FilterIndex = 3;

            DialogResult result = saveFileDialog.ShowDialog(); // Show the dialog.
            DataTable dataTable = Excel.dataTable;

            dataTable.Select().Where(x => x.ItemArray[0].ToString() == textBox2.Text && x.ItemArray[1].ToString() == textBox3.Text)
                .ToList().ForEach(D => D.SetField("CATALOG_DESC", textBox4.Text));

            dataGridView1.DataSource = dataTable;

            if (result == DialogResult.OK) // Test result.
            {
                string file = saveFileDialog.FileName;

                Excel dataToExcel = new Excel();
                dataToExcel.SaveDataToExcel(dataGridView1, file);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DialogResult dialog = new DialogResult();

            dialog = MessageBox.Show("Do you want to close?", "Alert!", MessageBoxButtons.YesNo);

            if (dialog == DialogResult.Yes)
            {
                System.Environment.Exit(1);
            }
        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Control.ModifierKeys == Keys.Control && e.KeyChar == 5)
            {
                button3_Click(this, new EventArgs());
            }
        }
    }
}
