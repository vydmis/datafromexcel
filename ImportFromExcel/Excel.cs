using System.Data;
using System.Linq;
using System.Windows.Forms;
using GemBox.Spreadsheet;
using GemBox.Spreadsheet.WinFormsUtilities;

namespace ImportFromExcel
{
    class Excel
    {
        public static DataTable dataTable;

        public DataTable LoadDataFromExcel(string file)
        {
            DataTable dt = new DataTable();
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            var workbook = ExcelFile.Load(file);

            foreach (var worksheet in workbook.Worksheets)
            {
                dt.Columns.Add(worksheet.Rows[0].AllocatedCells[0].Value.ToString());
                dt.Columns.Add(worksheet.Rows[0].AllocatedCells[1].Value.ToString());
                dt.Columns.Add(worksheet.Rows[0].AllocatedCells[2].Value.ToString());

                foreach (var row in worksheet.Rows.Skip(1))
                {
                    DataRow dr = dt.NewRow();
                    int i = 0;

                    foreach (var cell in row.AllocatedCells)
                    {
                        if (cell.ValueType != CellValueType.Null)
                        {
                            dr[i] = cell.Value;
                        }
                        i++;
                    }
                    dt.Rows.Add(dr);
                }
            }

            return dt;
        }

        public void SaveDataToExcel(DataGridView dataGridView1, string file)
        {
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

            var workbook = new ExcelFile();
            var worksheet = workbook.Worksheets.Add("Sheet1");

            // From DataGridView to ExcelFile.
            DataGridViewConverter.ImportFromDataGridView(worksheet, dataGridView1, new ImportFromDataGridViewOptions() { ColumnHeaders = true });

            workbook.Save(file);
        }
    }
}
