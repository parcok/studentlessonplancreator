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
using Dropbox.Api;

namespace NavExcel {
    public partial class Form1 : Form {
        public Form1() {
            InitializeComponent();
            if (System.IO.File.Exists(Environment.CurrentDirectory + "/token.txt")) {
                token = File.ReadAllText(Environment.CurrentDirectory + "/token.txt");
            }
        }
        string token = "";
        string path = "";
        ArrayList dropboxPDFs = new ArrayList();
        string[] edmontonSpecialties = { "Alberta High", "Arctic High", "Calgary Enroute", "Calgary Terminal", "Calgary Tower", "Edmonton Enroute", "Edmonton Terminal", "North Low" };
        string[] ganderSpecialties = { "ATOS", "FSS", "High Level Domestic", "IFSS", "Low Level Domestic", "Ocean", "Planner" };
        string[] monctonSpecialties = { "Generic", "High Level", "Maritime", "Terminal" };
        string[] montrealSpecialties = { "Capitales", "Est", "Nord", "St-Laurent", "Sud", "Montreal Terminal", "Montreal De Tour" };
        string[] torontoSpecialties = { "Airports", "East High", "East Low", "North", "North Bay", "Pearson Tower", "TMU", "Terminal", "West High", "West Low" };
        string[] vancouverSpecialties = { "ATOS", "Airports", "East", "TMU West", "Vancouver High", "Vancouver Terminal", "Vancouver Tower", "Victoria Terminal", "Victoria Tower" };
        string[] winnipegSpecialties = { "Airports", "East High", "East Low", "North", "West High", "West Low", "Winnipeg Tower" };

        // Browse for folder
        private void button1_Click(object sender, EventArgs e) {
            using (var folderDialog = new FolderBrowserDialog()) {
                if (folderDialog.ShowDialog() == DialogResult.OK) {
                    path = folderDialog.SelectedPath;
                    textBox1.Text = path;
                }
            }
        }

        private void button2_Click(object sender, EventArgs e) {
            if (path == "") {
                MessageBox.Show("Please select the root folder with the browse button.");
            } else {
                button2.Text = "Running...";
                string[] wordFiles = Directory.GetFiles(path, "*", SearchOption.AllDirectories).Where(s => s.EndsWith("Instructor.doc") && !s.StartsWith("~") || s.EndsWith("Instructor.docx") && !s.StartsWith("~") || s.EndsWith("Instructor.docm") && !s.StartsWith("~")).ToArray();
                int fileAmount = wordFiles.Length;
                progressBar1.Maximum = fileAmount;
                foreach (string file in wordFiles) {
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

                        string instructorPDF = file;
                        if (file.EndsWith(".doc")) {
                            instructorPDF = instructorPDF.Replace(".doc", ".pdf");
                        } else if (instructorPDF.EndsWith(".docm")) {
                            instructorPDF = instructorPDF.Replace(".docm", ".pdf");
                        } else {
                            instructorPDF = instructorPDF.Replace(".docx", ".pdf");
                        }
                        //Console.WriteLine("Trying to save as " + instructorPDF);
                        aDoc.SaveAs(instructorPDF, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF);

                        // START OF VBA MACRO FROM NAV CANADA
                        foreach (Word.Style style in aDoc.Styles) {
                            if (style.NameLocal.Length > 4) {
                                if (style.NameLocal.Substring(0, 5) == "Instr") {
                                    //Console.WriteLine("Removing all of " + style.NameLocal);
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

                        //Console.WriteLine("Trying to save as " + newFile);
                        //aDoc.SaveAs(newFile, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocumentDefault);
                        //Console.WriteLine("Trying to save as " + newPDF);
                        aDoc.SaveAs(newPDF, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF);
                        dropboxPDFs.Add(newPDF);
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
                if (textBox2.Text == "") {
                    if (checkBox1.Checked) {
                        MessageBox.Show("All completed successfully. Now uploading to dropbox.");
                    } else {
                        MessageBox.Show("All completed successfully.");
                    }
                } else {
                    MessageBox.Show("Complete. Please verify the template of files listed above.");
                }
                if (checkBox1.Checked && token != "") {
                    button2.Text = "Uploading...";
                    progressBar1.Value = 0;
                    progressBar1.Maximum = dropboxPDFs.Count;
                    foreach (string s in dropboxPDFs) {
                        DropboxClient dropboxClient = new DropboxClient(token);
                        string folder = "/" + comboBox2.Text;
                        string[] filenamesplit = s.Split('\\');
                        string filename = filenamesplit[filenamesplit.Length - 1];
                        var content = System.IO.File.ReadAllBytes(s);
                        try {
                            Form1.UploadAsync(dropboxClient, folder, filename, content).Wait();
                            progressBar1.Value++;
                            progressBar1.Update();
                        } catch {
                            textBox2.Text += "Could not upload. " + filename + "\r\n";
                        }
                    }
                    MessageBox.Show("Completed uploading files.");
                } else if (checkBox1.Checked && token == "") {
                    MessageBox.Show("Please put your token in a file named token.txt and restart the program.");
                }
                button2.Text = "Start";
            }
        }

        static async Task UploadAsync(DropboxClient dbx, string folder, string file, byte[] content) {
            using (var mem = new MemoryStream(content)) {
                var updated = await dbx.Files.UploadAsync(folder + "/" + file, Dropbox.Api.Files.WriteMode.Overwrite.Instance, body: mem).ConfigureAwait(false);
            }
        }

        private void Form1_Load(object sender, EventArgs e) {
            comboBox1.SelectedIndex = 0;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e) {
            switch (comboBox1.SelectedIndex) {
                case 1: // Edmonton
                    comboBox2.Items.Clear();
                    foreach (string s in edmontonSpecialties) {
                        comboBox2.Items.Add(s);
                    }
                    comboBox2.SelectedIndex = 0;
                    break;
                case 2: // Gander
                    comboBox2.Items.Clear();
                    foreach (string s in ganderSpecialties) {
                        comboBox2.Items.Add(s);
                    }
                    comboBox2.SelectedIndex = 0;
                    break;
                case 3: // Moncton
                    comboBox2.Items.Clear();
                    foreach (string s in monctonSpecialties) {
                        comboBox2.Items.Add(s);
                    }
                    comboBox2.SelectedIndex = 0;
                    break;
                case 4: // Montreal
                    comboBox2.Items.Clear();
                    foreach (string s in montrealSpecialties) {
                        comboBox2.Items.Add(s);
                    }
                    comboBox2.SelectedIndex = 0;
                    break;
                case 5: // Toronto
                    comboBox2.Items.Clear();
                    foreach (string s in torontoSpecialties) {
                        comboBox2.Items.Add(s);
                    }
                    comboBox2.SelectedIndex = 0;
                    break;
                case 6: // Vancouver
                    comboBox2.Items.Clear();
                    foreach (string s in vancouverSpecialties) {
                        comboBox2.Items.Add(s);
                    }
                    comboBox2.SelectedIndex = 0;
                    break;
                case 7: // Winnipeg
                    comboBox2.Items.Clear();
                    foreach (string s in winnipegSpecialties) {
                        comboBox2.Items.Add(s);
                    }
                    comboBox2.SelectedIndex = 0;
                    break;
                default:
                    comboBox2.Items.Clear();
                    break;
            }
        }
    }
}
