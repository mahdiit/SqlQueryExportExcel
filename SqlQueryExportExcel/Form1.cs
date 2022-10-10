using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SqlQueryExportExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        }

        private string GetConnection()
        {
            SqlConnectionStringBuilder conn = new SqlConnectionStringBuilder();
            conn.InitialCatalog = txtDb.Text;
            conn.DataSource = txtIP.Text;
            if (chkWinAthu.Checked)
            {
                conn.IntegratedSecurity = true;
            }
            else
            {
                conn.Password = txtPassword.Text;
                conn.UserID = txtUsername.Text;
            }

            return conn.ConnectionString;
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            btnExport.Enabled = false;
            Application.DoEvents();

            var sqlConnection = new SqlConnection(GetConnection());
            var sqlCommand = new SqlCommand(txtQuery.Text, sqlConnection)
            {
                CommandTimeout = 300
            };

            var exportFile = "";
            if (exportDialog.ShowDialog() == DialogResult.OK)
            {
                exportFile = exportDialog.FileName;
            }
            else
            {
                MessageBox.Show("Error");
                return;
            }

            sqlConnection.Open();
            SqlDataReader reader = null;
            try
            {
                var f = new FileInfo(exportFile);
                ExcelPackage xlsx = new ExcelPackage(f);

                reader = sqlCommand.ExecuteReader();
                int sheetIndex = 1;
                do
                {
                    int totalCount = 0;
                    var colPos = new Dictionary<string, int>();
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        colPos.Add(reader.GetName(i), i);
                    }

                    var ws = xlsx.Workbook.Worksheets.Add("Result-" + sheetIndex);
                    ws.Cells.Style.Font.Name = "Tahoma";
                    ws.Cells.Style.Font.Size = 9.5f;
                    ws.Cells.Style.Numberformat.Format = "@";

                    var firstHeader = ws.Cells[1, 1, 1, colPos.Count + 1];
                    firstHeader.Merge = true;
                    firstHeader.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    firstHeader.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    firstHeader.Style.Font.Bold = true;
                    firstHeader.Value = DateTime.Now.ToShortDateString();
                    ws.Row(1).Height = 25;

                    ws.Cells[2, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[2, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    ws.Cells[2, 1].Style.WrapText = false;
                    ws.Cells[2, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[2, 1].Style.Fill.BackgroundColor.SetColor(Color.Gainsboro);
                    ws.Cells[2, 1].Value = "No";

                    while (reader.Read())
                    {
                        if (totalCount == 0)
                        {
                            int i = 1;
                            foreach (var item in colPos)
                            {
                                ws.Cells[2, i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                ws.Cells[2, i + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                ws.Cells[2, i + 1].Style.WrapText = false;
                                ws.Cells[2, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                ws.Cells[2, i + 1].Style.Fill.BackgroundColor.SetColor(Color.Gainsboro);
                                ws.Cells[2, i + 1].Value = item.Key;
                                i++;
                            }
                            ws.Row(2).Height = 20;
                            ws.Cells.AutoFitColumns();
                        }

                        totalCount++;

                        ws.Row(totalCount + 1).Height = 16.5;
                        ws.Cells[totalCount + 2, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Cells[totalCount + 2, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        ws.Cells[totalCount + 2, 1].Value = totalCount;

                        var j = 1; //ستون اول ستون ردیف هست
                        foreach (var item in colPos)
                        {
                            var col = ws.Cells[totalCount + 2, j + 1];
                            col.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            var colValue = reader.GetValue(colPos[item.Key]);
                            //if (OnRenderColumn != null)
                            //    OnRenderColumn(item.Key, colValue, col);
                            //else
                            col.Value = colValue;
                            //if (dataRead != null)
                            //    dataRead(item.Key, colValue);
                            j++;
                        }

                        if (totalCount % 100 == 0 && totalCount < 3000)
                        {
                            ws.Cells.AutoFitColumns();
                            //OnProgress("تعداد رکوردهای پردازش شده:  " + totalCount);
                        }

                        //بیش اندازه بزرگ شده است
                        if (totalCount == 1000000)
                        {
                            break;
                        }
                    }

                    ws.Cells.AutoFitColumns();
                    ws.View.PageLayoutView = false;
                    //ws.View.RightToLeft = true;

                    sheetIndex++;

                } while (reader.NextResult());

                xlsx.Save();
                xlsx.Dispose();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                //clean reader
                if (reader != null)
                    reader.Close();

                //clean command
                sqlCommand.Dispose();

                //clean connection
                sqlConnection.Close();
                sqlConnection.Dispose();
            }

            btnExport.Enabled = true;
        }
    }
}
