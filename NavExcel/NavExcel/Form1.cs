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
using Word = Microsoft.Office.Interop.Word;
using System.Collections;
using System.IO;

namespace NavExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        string path = "";

        // Browse for folder
        private void button1_Click(object sender, EventArgs e)
        {
            path = "\\\\central\\ops\\OperationalTraining\\YYZ\\2 Generic Training\\Kevin Testing (Please Don't Touch)\\Temp\\";
            //path = "D:\\Temp2\\";
            using (var folderDialog = new FolderBrowserDialog()) {
                if (folderDialog.ShowDialog() == DialogResult.OK) {
                    path = folderDialog.SelectedPath;
                    textBox1.Text = path;
                }
            }
            textBox1.Text = path;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (path == "") {
                MessageBox.Show("Please select the root folder with the browse button.");
            } else {
                string[] wordFiles = Directory.GetFiles(path, "*", SearchOption.AllDirectories).Where(s => s.EndsWith("Instructor.doc") && !s.StartsWith("~") || s.EndsWith("Instructor.docx") && !s.StartsWith("~") || s.EndsWith("Instructor.docm") && !s.StartsWith("~")).ToArray();
                int fileAmount = wordFiles.Length;
                progressBar1.Maximum = fileAmount;
                Stopwatch watch = new Stopwatch();
                watch.Start();
                foreach (string file in wordFiles) {
                    Console.WriteLine("FILE: " + file);
                    try {

                        object missing = System.Reflection.Missing.Value;
                        Word.Application wordApp = new Word.ApplicationClass();
                        Word.Document aDoc = null;

                        wordApp.Visible = false;
                        wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                        wordApp.Options.WarnBeforeSavingPrintingSendingMarkup = false;

                        aDoc = wordApp.Documents.Open(file/*filename*/, false/*convert file prompt*/,
                              false/*readonly*/, false/*recent files*/, ref missing/*read pass*/,
                              ref missing/*read template pass*/, true/*reopen*/, ref missing/*write pass*/,
                              ref missing/*write template pass*/, ref missing/*format*/, ref missing/*encoding*/,
                              false/*visible client*/, ref missing/*repair*/, ref missing/*direction*/,
                              ref missing/*NoEncodingDialog*/, ref missing/*XMLTransform*/);
                        aDoc = wordApp.Documents.Add(file/*template*/, ref missing/*new template*/, ref missing/*doc type*/, false/*visible*/);
                        aDoc.Activate();
                        aDoc.AcceptAllRevisions();
                        aDoc.TrackRevisions = false;

                        if (aDoc.Comments.Count > 0) {
                            aDoc.DeleteAllComments();
                        }

                        // START OF VBA MACRO FROM NAV CANADA
                        foreach (Word.Style style in aDoc.Styles) {
                            if (style.NameLocal.Length > 4) {
                                if (style.NameLocal.Substring(0, 5) == "Instr") {
                                    Console.WriteLine("Removing all of " + style.NameLocal);
                                    wordApp.Selection.Find.set_Style(style);
                                    wordApp.Selection.Find.Execute();
                                    while (wordApp.Selection.Find.Found) {
                                        wordApp.Selection.Delete();
                                        wordApp.Selection.Find.Execute();
                                    }
                                }
                            }
                            wordApp.Selection.HomeKey(Unit: Word.WdUnits.wdStory);
                        }

                        Word.Find findObject = wordApp.Application.Selection.Find;
                        findObject.ClearFormatting();
                        findObject.Text = "Instructor Manual";
                        findObject.Replacement.ClearFormatting();
                        findObject.Replacement.Text = "Student Manual";

                        object replaceAll = Word.WdReplace.wdReplaceAll;
                        findObject.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing, ref missing, ref missing,
                            ref replaceAll, ref missing, ref missing, ref missing, ref missing);
                        // END OF VBA MACRO FROM NAV CANADA

                        
                        // Saving docx
                        string newFile = file.Replace("Instructor", "Student");
                        if (newFile.EndsWith(".docm")) {
                            newFile = newFile.Replace(".docm", ".docx");
                        } else if (newFile.EndsWith(".doc")) {
                            newFile = newFile.Replace(".doc", ".docx");
                        }
                        // Saving PDF
                        string newPDF = "";
                        if (newFile.EndsWith(".doc")) {
                            newPDF = newFile.Replace(".doc", ".pdf");
                        } else if (newFile.EndsWith(".docm")) {
                            newPDF = newFile.Replace(".docm", ".pdf");
                        } else {
                            newPDF = newFile.Replace(".docx", ".pdf");
                        }
                        Console.WriteLine("Trying to save as " + newFile);
                        aDoc.SaveAs(newFile, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocumentDefault);
                        Console.WriteLine("Trying to save as " + newPDF);
                        aDoc.SaveAs(newPDF, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF);
                    } catch (Exception error) {
                        textBox2.Text += "Couldn't open " + file + " " + e + "\r\n" + error;
                    }
                    progressBar1.Value++;
                    progressBar1.Update();
                }
                Process[] wordClients = Process.GetProcessesByName("WINWORD");
                foreach (Process p in wordClients) {
                    p.Kill();
                }
                watch.Stop();
                Console.WriteLine("Total runtime: " + watch.ElapsedMilliseconds);
                if (textBox2.Text == "") {
                    MessageBox.Show("All completed successfully.");
                    MessageBox.Show("Took a total of " + watch.ElapsedMilliseconds + "ms.");
                } else {
                    MessageBox.Show("Complete. Please verify the template of files listed above.");
                }
            }
        }
    }
}
