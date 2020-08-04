using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Check__Dupicate_excels
{
    public partial class Comparable_Excel : Form
    {
        public Comparable_Excel()
        {
            InitializeComponent();
        }
        DataSet result1, result2;
        DataTable dt_1, dt_2;

        private DataTable compareDatatable(DataTable dt1, DataTable dt2)
        {
            var differences =
            dt1.AsEnumerable().Except(dt2.AsEnumerable(), DataRowComparer.Default);
            return differences.Any() ? differences.CopyToDataTable() : new DataTable();
        }

        private void btn_click_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog file = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx|Excel Workbook 97-2003|*.xls", ValidateNames = true })
            {
                if (file.ShowDialog() == DialogResult.OK)
                {
                    FileStream fs = File.Open(file.FileName, FileMode.Open, FileAccess.Read);
                    IExcelDataReader reader;
                    if (file.FilterIndex != 1)
                    {
                        reader = ExcelReaderFactory.CreateBinaryReader(fs);
                    }
                    else
                    {
                        reader = ExcelReaderFactory.CreateOpenXmlReader(fs);

                        result1 = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                            {
                                UseHeaderRow = true
                            }
                        });
                    }
                    com_1.Items.Clear();
                    foreach (DataTable dt in result1.Tables)
                    {
                        com_1.Items.Add(dt.TableName);

                    }
                    com_1.SelectedIndex = 0;
                    lbPath1.Text = file.FileName;
                    reader.Close();
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog file = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx|Excel Workbook 97-2003|*.xls", ValidateNames = true })
            {
                if (file.ShowDialog() == DialogResult.OK)
                {
                    FileStream fs = File.Open(file.FileName, FileMode.Open, FileAccess.Read);
                    IExcelDataReader reader;
                    if (file.FilterIndex != 1)
                    {
                        reader = ExcelReaderFactory.CreateBinaryReader(fs);
                    }
                    else
                    {
                        reader = ExcelReaderFactory.CreateOpenXmlReader(fs);

                        result2 = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                            {
                                UseHeaderRow = true
                            }
                        });
                    }
                    com_2.Items.Clear();
                    foreach (DataTable dt in result2.Tables)
                    {
                        com_2.Items.Add(dt.TableName);
                    }
                    com_2.SelectedIndex = 0;
                    lbPath2.Text = file.FileName;
                    reader.Close();
                }
            }
        }

        private void com_1_SelectedIndexChanged(object sender, EventArgs e)
        {

            dt_1 = result1.Tables[com_1.SelectedIndex];
            dataGridView_list_1.DataSource = dt_1;
        }

        private void com_2_SelectedIndexChanged(object sender, EventArgs e)
        {
            dt_2 = result2.Tables[com_2.SelectedIndex];
            dataGridView_list_2.DataSource = dt_2;
        }

        private void btn_comparable_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(lbPath1.Text)|| lbPath1.Text == "ไม่พบไฟล์")
            {
                MessageBox.Show("เลือกไฟล์ 1");
                return;
            }
            if (string.IsNullOrEmpty(lbPath2.Text)|| lbPath2.Text == "ไม่พบไฟล์")
            {
                MessageBox.Show("เลือกไฟล์ 2");
                return;
            }

            if (dt_1.Rows.Count == dt_2.Rows.Count)
            {
                this.lb_Status.Text = "จำนวนแถว " + dt_1.Rows.Count + " เท่ากัน";
                this.lb_Status.ForeColor = System.Drawing.Color.Green;

                DataTable dtDiff = new DataTable();

                List<string> listcolumn_1 = new List<string>();
                List<string> listcolumn_2 = new List<string>();
                Boolean checkColumn;

                foreach (DataColumn column in dt_1.Columns)
                {
                    listcolumn_1.Add(column.ColumnName);
                }

                foreach (DataColumn column in dt_2.Columns)
                {
                    listcolumn_2.Add(column.ColumnName);
                }

                if (listcolumn_1.Count == listcolumn_2.Count)
                {

                    for (int i = 0; i < listcolumn_1.Count; i++)
                    {
                        if (listcolumn_1[i] != listcolumn_2[i])
                        {

                            MessageBox.Show("Column ไม่ตรงกัน  " + listcolumn_1[i] + " , " + listcolumn_2[i]);
                            checkColumn = false;
                            return;
                        }
                        checkColumn = true;
                    }

                    if (checkColumn = true)
                    {

                        if (dt_1.Rows.Count > dt_2.Rows.Count)
                        {
                            dtDiff = compareDatatable(dt_1, dt_2);
                        }
                        else
                        {
                            dtDiff = compareDatatable(dt_2, dt_1);
                        }
                    }
                }
                if (dtDiff.Rows.Count == 0)
                {
                    dataGridView_Comparable.DataSource = dtDiff;
                    MessageBox.Show("ไฟล์ตรงกัน");
                }
                else
                {
                    dataGridView_Comparable.DataSource = dtDiff;
                }
            }
            else
            {
                this.lb_Status.Text = "จำนวนแถวไม่เท่ากันเท่ากัน ไฟล์ที่ 1 : " + dt_1.Rows.Count + " แถว ไฟล์ที่ 2: " + dt_2.Rows.Count + " แถว";
                this.lb_Status.ForeColor = System.Drawing.Color.Red;

                DataTable dtDiff = new DataTable();

                List<string> listcolumn_1 = new List<string>();
                List<string> listcolumn_2 = new List<string>();
                Boolean checkColumn;

                foreach (DataColumn column in dt_1.Columns)
                {
                    listcolumn_1.Add(column.ColumnName);
                }

                foreach (DataColumn column in dt_2.Columns)
                {
                    listcolumn_2.Add(column.ColumnName);
                }

                    for (int i = 0; i < listcolumn_1.Count; i++)
                    {
                        if (listcolumn_1[i] != listcolumn_2[i])
                        {

                            MessageBox.Show("Column ไม่ตรงกัน  " + listcolumn_1[i] + " , " + listcolumn_2[i]);
                            checkColumn = false;
                            return;
                        }
                        checkColumn = true;
                    }

                    if (checkColumn = true)
                    {

                        if (dt_1.Rows.Count > dt_2.Rows.Count)
                        {
                            dtDiff = compareDatatable(dt_1, dt_2);
                        }
                        else
                        {
                            dtDiff = compareDatatable(dt_2, dt_1);
                        }
                    }
                if (dtDiff.Rows.Count == 0)
                {
                    dataGridView_Comparable.DataSource = dtDiff;
                    MessageBox.Show("ไฟล์ตรงกัน");
                }
                else {
                    dataGridView_Comparable.DataSource = dtDiff;
                }
            }

        }
        
    }
}
