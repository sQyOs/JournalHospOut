using System.Windows;
using System;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using System.Configuration;
using System.Text.RegularExpressions;

namespace JournalHospOut
{
    public partial class MainWindow
    {
        Excel.Application eapp = new Excel.Application();
        private Excel.Range ecells;
        private Excel.Sheets esheets;
        private Excel.Worksheet eworksheet;
        int[,] columnConditions = new int[11, 2] { { 0, 0 }, { 15, 19 }, { 20, 24 }, { 25, 29 }, { 30, 34 }, { 35, 39 }, { 40, 44 }, { 45, 49 }, { 50, 54 }, { 55, 59 }, { 60, 999 } };
        string[] head = new string[14] {"Шифр по МКБ Х пересмотра",
                                        "Пол",
                                        "Число дней временной нетрудоспособности",
                                        "Число случаев временной нетрудоспособности",
                                        "15 - 19",
                                        "20 - 24",
                                        "25 - 29",
                                        "30 - 34",
                                        "35 - 39",
                                        "40 - 44",
                                        "45 - 49",
                                        "50 - 54",
                                        "55 - 59",
                                        "60 лет и старше"};
        public void test()
        {

            //string[,] rowConditions = new string[28, 4] { { "A00", "B99", null, null },
            //                                      { "A15", "A19", null, null },
            //                                      { "C00", "D48", null, null },
            //                                      { "C00", "D97", null, null },
            //                                      { "D50", "D89", null, null },
            //                                      { "E00", "E89", "E90", "E90" },
            //                                      { "E10", "E14", null, null },
            //                                      { "F00", "F99", null, null },
            //                                      { "G00", "G98", "G99", "G99" },
            //                                      { "H00", "H59", null, null },
            //                                      { "H60", "H95", null, null },
            //                                      { "I00", "I99", null, null },
            //                                      { "I20", "I25", null, null },
            //                                      { "I60", "I69", null, null },
            //                                      { "J00", "J98", "J99", "J99" },
            //                                      { "J00", "J01", "J04", "J06" },
            //                                      { "J09", "J09", "J11", "J11" },
            //                                      { "J12", "J18", null, null },
            //                                      { "K00", "K92", "K93", "K93" },
            //                                      { "L00", "L99", null, null },
            //                                      { "M00", "M99", null, null },
            //                                      { "N00", "N99", null, null },
            //                                      { "O00", "O99", null, null },
            //                                      { "Q00", "Q99", null, null },
            //                                      { "S00", "T98", null, null },
            //                                      { "A00", "Z99", null, null },
            //                                      { "O03", "O08", null, null },
            //                                      { "Z75", "Z75", null, null } };
            string[] conditionsTmp = ConfigurationManager.AppSettings["condition"].Split(';');
            string[][] conditions = new string[conditionsTmp.Length][];
            for (int i = 0; i < conditionsTmp.Length; i++)
            {
                conditions[i] = new string[conditionsTmp[i].Split(',').Length];
            }
            Regex regex = new Regex(@"\W");
            for (int i = 0; i < conditionsTmp.Length; i++)
            {
                for (int j = 0; j < conditionsTmp[i].Split(',').Length; j++)
                {
                    conditions[i][j] = regex.Replace(conditionsTmp[i].Split(',')[j], "");
                }
            }


            eapp.SheetsInNewWorkbook = 1;
            eapp.Workbooks.Add(Type.Missing);
            esheets = eapp.Worksheets;
            eworksheet = (Excel.Worksheet)esheets.get_Item(1);
            eapp.Visible = true;
            eapp.Interactive = false;
            cn.Open();
            int WriteRow = 1;
            eworksheet.Range[eworksheet.Cells[WriteRow, 1], eworksheet.Cells[WriteRow, head.Length]].WrapText = true;
            for (int i = 0; i < head.Length; i++)
            {
                eworksheet.Cells[WriteRow, i + 1] = head[i];
            }
            WriteRow++;
            

            for (int i = 0; i < conditions.GetLength(0); i++)
            {
                string textCondition = "";
                for (int k = 0; k < conditions[i].Length / 2; k++)
                {
                    if ((k % 2) != 0)
                    {
                        //textCondition += "\n\r";
                        textCondition += ", ";
                    }
                    if (conditions[i][2 * k] != conditions[i][2 * k + 1])
                    {
                        textCondition += conditions[i][2 * k] + "-" + conditions[i][2 * k + 1];
                    }
                    else
                    {
                        textCondition += conditions[i][2 * k];
                    }
                }
                if (!conditions[i][0].Contains("O"))
                {
                    eworksheet.Range[eworksheet.Cells[WriteRow, 1], eworksheet.Cells[WriteRow + 1, 1]].Merge();
                    eworksheet.Range[eworksheet.Cells[WriteRow, 1], eworksheet.Cells[WriteRow + 1, 1]].WrapText = true;
                    eworksheet.Range[eworksheet.Cells[WriteRow, 1], eworksheet.Cells[WriteRow + 1, 1]] = textCondition;
                    eworksheet.Cells[WriteRow, 2] = "М";
                    for (int j = 0; j < columnConditions.GetLength(0); j++)
                    {
                        if (columnConditions[j, 1] == 0)
                        {
                            dtToCell(makeCmdTextMultiMas(true, i, j, "М", conditions[i]), WriteRow, 3, eworksheet);
                        }
                        else
                        {
                            dtToCell(makeCmdTextMultiMas(false, i, j, "М", conditions[i]), WriteRow, j + 4, eworksheet);
                        }
                    }
                    WriteRow++;
                }
                else
                {
                    eworksheet.Cells[WriteRow, 1] = textCondition;
                }

                eworksheet.Cells[WriteRow, 2] = "Ж";
                for (int j = 0; j < columnConditions.GetLength(0); j++)
                {
                    if (columnConditions[j, 1] == 0)
                    {
                        dtToCell(makeCmdTextMultiMas(true, i, j, "Ж", conditions[i]), WriteRow, 3, eworksheet);
                    }
                    else
                    {
                        dtToCell(makeCmdTextMultiMas(false, i, j, "Ж", conditions[i]), WriteRow, j + 4, eworksheet);
                    }
                }
                WriteRow++;
            }

            cn.Close();
            ecells = eworksheet.Range[eworksheet.Cells[1, 1], eworksheet.Cells[conditions.Length * 2 - 1, columnConditions.GetLength(0) + 3]];
            ecells.Borders.ColorIndex = 1;
            ecells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            eapp.Interactive = true;
            MessageBox.Show("Отчет сформирован");
        }


        private DataTable makeCmdTextMultiMas(bool tmp, int i, int j, string pol, string[] rowConditions)
        {
            DataTable dtForReport = new DataTable();
            string whereRowCondition = "";
            for (int k = 0; k < rowConditions.Length / 2; k++)
            {
                if ((k % 2) != 0)
                {
                    whereRowCondition += " OR ";
                }
                whereRowCondition += $"([mkb] >= '{rowConditions[2 * k]}' AND [mkb] <= '{rowConditions[2 * k + 1]}')";
            }

            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            dtForReport.Reset();
            if (tmp)
                cmd.CommandText = @"SELECT jho.pol, Sum(jho.kd) AS [Sum-kd], Count(jho.kd) AS [Count-kd] FROM jho WHERE (" + whereRowCondition + ") AND [pol]='" + pol + "' GROUP BY jho.pol ORDER BY jho.pol DESC;";
            else
                cmd.CommandText = @"SELECT jho.pol, Count(jho.kd) AS [Count-kd] FROM jho WHERE (" + whereRowCondition + ") AND [pol]='" + pol + "' AND [age]>=" + columnConditions[j, 0] + " AND [age]<=" + columnConditions[j, 1] + " GROUP BY jho.pol ORDER BY jho.pol DESC;";
            da.Fill(dtForReport);
            return dtForReport;
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