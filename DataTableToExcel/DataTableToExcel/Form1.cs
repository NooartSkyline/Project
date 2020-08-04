using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DataTableToExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btn_Load_Click(object sender, EventArgs e)
        {
            dataGridView.DataSource = dt_Table();
        }

        private DataTable dt_Table()
        {
			try
			{
				DataTable dt = new DataTable();
				string ConString = ConfigurationManager.ConnectionStrings["ConString"].ConnectionString;
				SqlConnection con = new SqlConnection(ConString);
				SqlCommand cmd = new SqlCommand("select * from TBREQUESTAPPROVAL_DT", con);

				con.Open();

				SqlDataReader reader = cmd.ExecuteReader();

				dt.Load(reader);

				return dt;
			}
			catch (Exception)
			{

				throw;
			}
           
        }

        private void btn_Export_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel.Worksheet worksheet = worksheet = null;
            worksheet = workbook.Sheets["Sheet1"];
            worksheet = workbook.ActiveSheet;
            worksheet.Name = "Arty";

            for (int i = 1; i < dataGridView.Columns.Count+1; i++)
            {
                //worksheet.Cells[i, 1] = dataGridView.Columns[i - 1].HeaderText;
                worksheet.Cells[1, i] = dataGridView.Columns[i - 1].HeaderText;
            }
            for (int i = 0; i < dataGridView.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView.Columns.Count; j++)
                {
                    if (dataGridView.Rows[i].Cells[j].Value == null)
                    {
                        worksheet.Cells[i + 2, j + 1] = "";
                    }
                    else {
                        worksheet.Cells[i + 2, j + 1] = dataGridView.Rows[i].Cells[j].Value.ToString();
                    }
                }
            }

            var saveFileDialoge = new SaveFileDialog();
            saveFileDialoge.FileName = "OUTPUT";
            saveFileDialoge.DefaultExt = ".xlsx";
            if (saveFileDialoge.ShowDialog() == DialogResult.OK)
            {
                workbook.SaveAs(saveFileDialoge.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            }
            app.Quit();
        }
		public String SelectHD()
		{
			return "select * from TBREQUESTAPPROVAL_HD";
		}
	}
}
