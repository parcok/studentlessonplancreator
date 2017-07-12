using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections;
using System.IO;

namespace NavExcel {
    public partial class Form1 : Form {
        public Form1() {
            InitializeComponent();
        }
        string path = "";

        private void button1_Click(object sender, EventArgs e) {
            using (var folderDialog = new FolderBrowserDialog()) {
                if (folderDialog.ShowDialog() == DialogResult.OK) {
                    path = folderDialog.SelectedPath;
                    textBox1.Text = path;
                }
            }
        }

        private void button2_Click(object sender, EventArgs e) {
            if (textBox1.Text == "") {
                MessageBox.Show("Please select the root folder with the browse button.");
            } else {
                string[] excelFiles = Directory.GetFiles(path, "*", SearchOption.AllDirectories).Where(s => s.EndsWith(".xls") || s.EndsWith(".xlsx")).ToArray();
                int fileAmount = excelFiles.Length;
                progressBar1.Maximum = fileAmount;
                foreach (string file in excelFiles) {
                    try {

                        //Console.Write("FILE: " + file + " - ");
                        Excel.Application excel = new Excel.Application();
                        Excel.Workbook workbook = excel.Workbooks.Open(file);
                        Excel.Worksheet airSheet = excel.Sheets[1];
                        Excel.Worksheet groundSheet = excel.Sheets[2];
                        excel.DisplayAlerts = false;

                        // Cells go down then across
                        if ((string)airSheet.Cells[4, 1].value2 == "TIME") {
                            // THE WORKING SOLUTION TO FIND THE AMOUNT OF USED ROWS
                            int airRowAmount = airSheet.Cells.Find("*", System.Reflection.Missing.Value, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
                            int groundRowAmount = groundSheet.Cells.Find("*", System.Reflection.Missing.Value, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
                            // END OF THE WORKING SOLUTION TO FIND THE RIGHT AMOUNT OF USED ROWS


                            // Add colour to the tables
                            Excel.Range airTable = airSheet.Range["A4:G" + airRowAmount];
                            airTable.Cells.Interior.Color = System.Drawing.ColorTranslator.FromHtml("#C6E2E9"); // MATCH WITH AIRSHEET
                            groundSheet.Cells.Interior.Color = System.Drawing.ColorTranslator.FromHtml("#F1FFC4"); // MATCH WITH TITLERANGE

                            // Copy from ground to air
                            string rangeToCopy = "A4:G" + groundRowAmount;
                            Excel.Range rangeFrom = groundSheet.Range[rangeToCopy];
                            Excel.Range rangeTo = airSheet.Range["A" + (airRowAmount + 2)];
                            rangeFrom.Copy(rangeTo);

                            // Title for ground
                            airSheet.Range["A" + (airRowAmount + 1)].Cells.Font.Size = 36;
                            airSheet.Range["A" + (airRowAmount + 1)].Value2 = "GROUND";
                            Excel.Range titleRange = airSheet.Range["A" + (airRowAmount + 1) + ":G" + (airRowAmount + 1)];
                            MergeAndCenter(titleRange);
                            titleRange.Cells.Interior.Color = System.Drawing.ColorTranslator.FromHtml("#F1FFC4"); // MATCH WITH GROUNDSHEET
                            titleRange.Cells.Font.Bold = true;

                            // Get the total amount of rows
                            int totalRowAmount = airSheet.Cells.Find("*", System.Reflection.Missing.Value, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

                            // Need to get the range of the data we copied over, isolate the vehicles, then sort again
                            Excel.Range groundTable = airSheet.Range["A" + (airRowAmount + 3) + ":G" + totalRowAmount];
                            groundTable.Sort(groundTable.Columns[3, Type.Missing], Excel.XlSortOrder.xlDescending, groundTable.Columns[3, Type.Missing], Type.Missing, Excel.XlSortOrder.xlAscending, Type.Missing, Excel.XlSortOrder.xlAscending, Excel.XlYesNoGuess.xlYes, Type.Missing, Type.Missing, Excel.XlSortOrientation.xlSortColumns, Excel.XlSortMethod.xlPinYin, Excel.XlSortDataOption.xlSortNormal, Excel.XlSortDataOption.xlSortNormal, Excel.XlSortDataOption.xlSortNormal);

                            int vehicleAmount = 0;
                            Excel.Range vehicleGrab = airSheet.Range["C" + (airRowAmount + 3) + ":C" + totalRowAmount];
                            foreach (Excel.Range item in vehicleGrab.Cells) {
                                if (item.Text == "") {
                                    vehicleAmount++;
                                }
                            }

                            // Only if there are vehicles
                            if (vehicleAmount > 0) {
                                Excel.Range sortNonVehicles = airSheet.Range["A" + (airRowAmount + 3) + ":G" + (totalRowAmount - vehicleAmount)];
                                sortNonVehicles.Sort(groundTable.Columns[1, Type.Missing], Excel.XlSortOrder.xlAscending, groundTable.Columns[1, Type.Missing], Type.Missing, Excel.XlSortOrder.xlAscending, Type.Missing, Excel.XlSortOrder.xlAscending, Excel.XlYesNoGuess.xlYes, Type.Missing, Type.Missing, Excel.XlSortOrientation.xlSortColumns, Excel.XlSortMethod.xlPinYin, Excel.XlSortDataOption.xlSortNormal, Excel.XlSortDataOption.xlSortNormal, Excel.XlSortDataOption.xlSortNormal);

                                Excel.Range sortVehicles = airSheet.Range["A" + (totalRowAmount - vehicleAmount + 1) + ":G" + totalRowAmount];
                                sortVehicles.Cells.Interior.Color = System.Drawing.ColorTranslator.FromHtml("#FFCAAF");
                                sortVehicles.Sort(groundTable.Columns[1, Type.Missing], Excel.XlSortOrder.xlAscending, groundTable.Columns[1, Type.Missing], Type.Missing, Excel.XlSortOrder.xlAscending, Type.Missing, Excel.XlSortOrder.xlAscending, Excel.XlYesNoGuess.xlYes, Type.Missing, Type.Missing, Excel.XlSortOrientation.xlSortColumns, Excel.XlSortMethod.xlPinYin, Excel.XlSortDataOption.xlSortNormal, Excel.XlSortDataOption.xlSortNormal, Excel.XlSortDataOption.xlSortNormal);
                            }

                            // Title for tower
                            Excel.Range towerTitle = (Excel.Range)airSheet.Rows[4];
                            towerTitle.Insert();
                            airSheet.Range["A4"].Cells.Font.Size = 36;
                            airSheet.Range["A4"].Value2 = "TOWER";
                            towerTitle = airSheet.Range["A4:G4"];
                            MergeAndCenter(towerTitle);
                            towerTitle.Cells.Interior.Color = System.Drawing.ColorTranslator.FromHtml("#C6E2E9");
                            towerTitle.Cells.Font.Bold = true;

                            // Fix borders
                            Excel.Range fixBorders = airSheet.Range["A4:G" + (totalRowAmount + 1)];
                            fixBorders.Cells.Borders.Weight = Excel.XlBorderWeight.xlThick;// (Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick, Excel.XlColorIndex.xlColorIndexAutomatic, System.Drawing.Color.Black, System.Drawing.Color.Black);

                            airSheet.Sort.SortFields.Clear();

                            airSheet.Columns.AutoFit();
                            airSheet.Rows.AutoFit();

                            airSheet.PageSetup.PrintArea = "A1:G" + (totalRowAmount + 1);

                            excel.DisplayAlerts = false;
                            if (file.EndsWith(".xls")) {
                                workbook.SaveAs(file + "x", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
                            } else {
                                workbook.SaveAs(file, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
                            }

                        }
                        workbook.Sheets[2].delete();
                        workbook.Sheets[2].delete();
                        workbook.Close(true, Type.Missing, Type.Missing);
                        excel.Quit();

                    } catch {
                        textBox2.Text += "Couldn't open " + file + "\r\n";
                    }
                    progressBar1.Value++;
                    progressBar1.Update();
                }
                Process[] excelClients = Process.GetProcessesByName("excel");
                foreach (Process p in excelClients) {
                    p.Kill();
                }
                if (textBox2.Text == "") {
                    MessageBox.Show("All completed successfully.");
                } else {
                    MessageBox.Show("Complete. Please verify the template of files listed above.");
                }
            }
        }

        public static void MergeAndCenter(Excel.Range MergeRange) {
            MergeRange.Select();

            MergeRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            MergeRange.VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
            MergeRange.WrapText = false;
            MergeRange.Orientation = 0;
            MergeRange.AddIndent = false;
            MergeRange.IndentLevel = 0;
            MergeRange.ShrinkToFit = false;
            MergeRange.ReadingOrder = (int)(Excel.Constants.xlContext);
            MergeRange.MergeCells = false;

            MergeRange.Merge(System.Type.Missing);
        }
    }
}
