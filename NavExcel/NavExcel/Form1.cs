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

namespace NavExcel {
    public partial class Form1 : Form {
        public Form1() {
            InitializeComponent();
        }
        string path = "";
        string template = "";

        // Browse for folder
        private void button1_Click(object sender, EventArgs e) {
            path = "\\\\central\\ops\\OperationalTraining\\YYZ\\2 Generic Training\\Kevin Testing (Please Don't Touch)\\";
            //path = "D:\\Temp2\\";
            using (var folderDialog = new FolderBrowserDialog()) {
                if (folderDialog.ShowDialog() == DialogResult.OK) {
                    path = folderDialog.SelectedPath;
                    textBox1.Text = path;
                }
            }
            textBox1.Text = path;
        }

        // Browse for template
        private void button3_Click(object sender, EventArgs e) {
            using (var fileSelection = new OpenFileDialog()) {
                fileSelection.Filter = "Word Templates | *.dotm";
                if (fileSelection.ShowDialog() == DialogResult.OK) {
                    template = fileSelection.FileName;
                    textBox3.Text = template;
                }
            }
        }

        private void button2_Click(object sender, EventArgs e) {
            if (path == "") {
                MessageBox.Show("Please select the root folder with the browse button.");
            /*} else if (template == "") {
                MessageBox.Show("Please select the template you wish to apply.");*/
            } else {
                string[] wordFiles = Directory.GetFiles(path, "*", SearchOption.AllDirectories).Where(s => s.EndsWith(".doc") || s.EndsWith(".docx")).ToArray();
                int fileAmount = wordFiles.Length;
                progressBar1.Maximum = fileAmount;
                foreach (string file in wordFiles) {
                    try {

                        object missing = System.Reflection.Missing.Value;
                        Word.Application wordApp = new Word.ApplicationClass();
                        Word.Document aDoc = null;

                        wordApp.Visible = false;
                        wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                        aDoc = wordApp.Documents.Add(template/*template*/, ref missing/*new template*/, ref missing/*doc type*/, false/*visible*/);

                        aDoc = wordApp.Documents.Open(file/*filename*/, false/*convert file prompt*/,
                                                      false/*readonly*/, false/*recent files*/, ref missing/*read pass*/,
                                                      ref missing/*read template pass*/, true/*reopen*/, ref missing/*write pass*/,
                                                      ref missing/*write template pass*/, ref missing/*format*/, ref missing/*encoding*/,
                                                      false/*visible client*/, ref missing/*repair*/, ref missing/*direction*/,
                                                      ref missing/*NoEncodingDialog*/, ref missing/*XMLTransform*/);
                        System.Threading.Thread.Sleep(30000);
                        //Console.WriteLine(aDoc.Name);
                        wordApp.Documents.CheckOut(file);
                        System.Threading.Thread.Sleep(30000);
                        //aDoc.Activate();
                        attachTemplate(aDoc, template);

                        aDoc.SaveAs(file/*filename*/, ref missing/*file format*/, ref missing/*lock comments*/,
                                    ref missing/*pass*/, ref missing/*recent files*/, ref missing/*write pass*/,
                                    false/*read only suggest*/, ref missing/*embed fonts*/, ref missing/*native pic format*/,
                                    ref missing/*form data*/, ref missing/*AOCE letter*/, ref missing/*encoding*/,
                                    ref missing/*line breaks*/, ref missing/*substitutions*/, ref missing/*line endigng*/, ref missing/*BiDi marks*/);
                        if (aDoc.CanCheckin()) {
                            aDoc.CheckIn();
                        } else {
                            textBox2.Text += "Could not check in " + file + "\r\n";
                        }
                        aDoc.Close(Word.WdSaveOptions.wdSaveChanges, ref missing, ref missing);

                    } catch {
                        textBox2.Text += "Couldn't open " + file + "\r\n";
                    }
                    progressBar1.Value++;
                    progressBar1.Update();
                }
                Process[] wordClients = Process.GetProcessesByName("winword");
                foreach (Process p in wordClients) {
                    p.Kill();
                }
                if (textBox2.Text == "") {
                    MessageBox.Show("All completed successfully.");
                } else {
                    MessageBox.Show("Complete. Please verify the template of files listed above.");
                }
            }
        }

        private static void attachTemplate(Word.Document document, string templateFilePath) {
            object oTemplate = (object)templateFilePath;
            document.set_AttachedTemplate(ref oTemplate);
            document.UpdateStyles();
            MessageBox.Show("Successfully attached template.\r\n"); 
        }
    }
}
