using System.Windows;
using System;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;

namespace JournalHospOut
{
    public partial class MainWindow
    {
        Excel.Application eapp = new Excel.Application();
        private Excel.Range ecells;
        private Excel.Sheets esheets;
        private Excel.Worksheet eworksheet;
        int[,] columnConditions = new int[11, 2] { { 0, 0 }, { 15, 19 }, { 20, 24 }, { 25, 29 }, { 30, 34 }, { 35, 39 }, { 40, 44 }, { 45, 49 }, { 50, 54 }, { 55, 59 }, { 60, 999 } };
        string[,] rowConditions = new string[,] { { "A00", "B99", null, null },
                                                  { "A15", "A19", null, null },
                                                  { "C00", "D48", null, null },
                                                  { "C00", "D09", null, null },
                                                  { "D50", "D89", null, null },
                                                  { "E00", "E89", "E90", "E90" },
                                                  { "E10", "E14", null, null },
                                                  { "F00", "F99", null, null },
                                                  { "G00", "G98", "G99", "G99" },
                                                  { "H00", "H59", null, null },
                                                  { "H60", "H95", null, null },
                                                  { "I00", "I99", null, null },
                                                  { "I20", "I25", null, null },
                                                  { "I60", "I69", null, null },
                                                  { "J00", "J98", "J99", "J99" },
                                                  { "J00", "J01", "J04", "J06" },
                                                  { "J09", "J09", "J11", "J11" },
                                                  { "J12", "J12", "J18", "J18" },
                                                  { "K00", "K92", "K93", "K93" },
                                                  { "L00", "L99", null, null },
                                                  { "M00", "M99", null, null },
                                                  { "N00", "N99", null, null },
                                                  { "O00", "O99", null, null },
                                                  { "Q00", "Q99", null, null },
                                                  { "S00", "T98", null, null },
                                                  { "O03", "O08", null, null } };
        public void test()
        {
            eapp.SheetsInNewWorkbook = 1;
            eapp.Workbooks.Add(Type.Missing);
            esheets = eapp.Worksheets;
            eworksheet = (Excel.Worksheet)esheets.get_Item(1);
            eapp.Visible = true;
            eapp.Interactive = false;
            cn.Open();
            int WriteRow =1;
            for (int i = 0; i < rowConditions.GetLength(0); i++)
            {
                
                if (rowConditions[i, 0] != "O00")
                {
                    eworksheet.Range[eworksheet.Cells[WriteRow, 1], eworksheet.Cells[WriteRow + 1, 1]].Merge();
                    eworksheet.Range[eworksheet.Cells[WriteRow, 1], eworksheet.Cells[WriteRow + 1, 1]] = rowConditions[i, 0] + "-" + rowConditions[i, 1] + "\n\r" + rowConditions[i, 2] + "-" + rowConditions[i, 3];
                    eworksheet.Cells[WriteRow, 2] = "М";
                    for (int j = 0; j < columnConditions.GetLength(0); j++)
                    {
                        if (columnConditions[j, 1] == 0)
                        {
                            makeCmdText(true, i, j, "М");
                            dtToCell(dtForReport, WriteRow, 3, eworksheet);
                        }
                        else
                        {
                            makeCmdText(false, i, j, "М");
                            dtToCell(dtForReport, WriteRow, j + 4, eworksheet);
                        }
                    }
                    WriteRow++;
                }
                else
                {
                    eworksheet.Cells[WriteRow, 1] = rowConditions[i, 0] + "-" + rowConditions[i, 1];
                }
                
                eworksheet.Cells[WriteRow, 2] = "Ж";
                for (int j = 0; j < columnConditions.GetLength(0); j++)
                {
                    if (columnConditions[j, 1] == 0)
                    {
                        makeCmdText(true, i, j, "Ж");
                        dtToCell(dtForReport, WriteRow, 3, eworksheet);
                    }
                    else
                    {
                        makeCmdText(false, i, j, "Ж");
                        dtToCell(dtForReport, WriteRow, j + 4, eworksheet);
                    }
                }
                WriteRow++;
            }
            
            cn.Close();
            ecells = eworksheet.Range[eworksheet.Cells[1, 1], eworksheet.Cells[rowConditions.GetLength(0) * 2 - 1, columnConditions.GetLength(0) + 3]];
            ecells.Borders.ColorIndex = 1;
            ecells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            eapp.Interactive = true;
            MessageBox.Show("Отчет сформирован");
        }

        private void makeCmdText(bool tmp, int i, int j, string pol)
        {
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            dtForReport.Reset();
            if (tmp)
                cmd.CommandText = @"SELECT jho.pol, Sum(jho.kd) AS [Sum-kd], Count(jho.kd) AS [Count-kd] FROM jho WHERE (([mkb]>='" + rowConditions[i, 0] + "' AND [mkb]<='" + rowConditions[i, 1] + "') OR ([mkb]>='" + rowConditions[i, 2] + "' AND [mkb]<='" + rowConditions[i, 3] + "')) AND [pol]='" + pol + "' GROUP BY jho.pol ORDER BY jho.pol DESC;";
            else
                cmd.CommandText = @"SELECT jho.pol, Count(jho.kd) AS [Count-kd] FROM jho WHERE (([mkb]>='" + rowConditions[i, 0] + "' AND [mkb]<='" + rowConditions[i, 1] + "') OR ([mkb]>='" + rowConditions[i, 2] + "' AND [mkb]<='" + rowConditions[i, 3] + "')) AND [pol]='" + pol + "' AND [age]>=" + columnConditions[j, 0] + " AND [age]<=" + columnConditions[j, 1] + " GROUP BY jho.pol ORDER BY jho.pol DESC;";
            da.Fill(dtForReport);
        }

        private void dtToCell(DataTable dt, int startRow, int startColumn, dynamic worksheets)
        {
            try
            {
                for (var i = 0; i < dt.Rows.Count; i++)
                    for (var j = 1; j < dt.Columns.Count; j++)
                        worksheets.Cells[i + startRow, j + startColumn - 1] = dt.Rows[i][j];
            }
            catch (Exception ex)
            {
                MessageBox.Show("" + ex);
            }
        }
    }
}