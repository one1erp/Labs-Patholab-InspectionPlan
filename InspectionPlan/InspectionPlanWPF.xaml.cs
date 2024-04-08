using LSEXT;
using LSExtensionWindowLib;
using LSSERVICEPROVIDERLib;
using Patholab_Common;
using Patholab_DAL_V1;
using System.Data.OracleClient;
using System.Runtime.InteropServices;
using System;
using draw = System.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Data;
using MSXML;
using System.Diagnostics;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Configuration;
using oracleData = Oracle.DataAccess.Client;
using forms = System.Windows.Forms;
using System.Threading;
using Spire.Doc;
using wdUnits = Microsoft.Office.Interop.Word;


namespace InspectionPlan
{
    /// <summary>
    /// Interaction logic for InspectionPlanWPF.xaml
    /// </summary>
    public partial class InspectionPlanWPF : UserControl
    {
        #region Private members

        private INautilusProcessXML xmlProcessor;
        private INautilusUser _ntlsUser;
        private OracleConnection connection;
        private IExtensionWindowSite2 _ntlsSite;
        private INautilusServiceProvider _sp;
        private INautilusDBConnection _ntlsCon;
        private DataLayer dal;
        private SDG currentSdg = null;
        private inspectionPlan currentInspection;
        private long currentResultID = -1;
        private string currentResultName = string.Empty;
        private string currentRtfPath = string.Empty;
        private bool saveToDB = false;
        private List<long> currentResultIDs;
        private forms.UserControl parentUserControl;


        Word.Application wordFile = null;
        Word.Document document = null;
        private Object oMissing = System.Reflection.Missing.Value;
         

        ONE1_richTextCtrl.RichSpellCtrl richTextMacro;
        ONE1_richTextCtrl.RichSpellCtrl richTextDiagnos;
        ONE1_richTextCtrl.RichSpellCtrl richTextMicro;

        #endregion

        public InspectionPlanWPF(INautilusServiceProvider sp, INautilusProcessXML xmlProcessor, INautilusDBConnection ntlsCon, IExtensionWindowSite2 _ntlsSite, INautilusUser _ntlsUser)
        {
            try
            {
                InitializeComponent();

                this._sp = sp;
                this.xmlProcessor = xmlProcessor;
                this._ntlsCon = ntlsCon;
                this._ntlsSite = _ntlsSite;
                this._ntlsUser = _ntlsUser;
                currentResultIDs = new List<long>();

                textBoxEnterSDG.Focusable = true;
                textBoxEnterSDG.Focus();

                openConnection();

                this.dal = new DataLayer();
                dal.Connect(_ntlsCon);




                /////////////////////////////////////////////////////////////start handle rich texts///////////////////////////////////

                richTextMacro = new ONE1_richTextCtrl.RichSpellCtrl();
                richTextDiagnos = new ONE1_richTextCtrl.RichSpellCtrl();
                richTextMicro = new ONE1_richTextCtrl.RichSpellCtrl();

                richTextMicro.ExtraBtnClciked += this.extraBtn_Click;


                DockPanel.SetDock(winformsHostMacro, Dock.Right);
                DockPanel.SetDock(winformsHostDiagnos, Dock.Left);
                DockPanel.SetDock(winformsHostMicro, Dock.Bottom);


                winformsHostMacro.Child = richTextMacro;
                winformsHostDiagnos.Child = richTextDiagnos;
                winformsHostMicro.Child = richTextMicro;

                /////////////////////////////end handle rich texts////////////////////////////////////////////////////////////////////

                // multiline text for buttons
                buttonOpenExistWord.Content = "Open existing" + Environment.NewLine + "word document";
                buttonDocxToRtf.Content = "Convert docx" + Environment.NewLine + "     to RTF";
                buttonUpdateDB.Content = "Add RTF" + Environment.NewLine + "to database";


                // hint in textBoxEnterSDG
                textBoxEnterSDG.Foreground = Brushes.Gray;
                textBoxEnterSDG.Text = "Search SDG...";
                textBoxEnterSDG.GotKeyboardFocus += new KeyboardFocusChangedEventHandler(tb_GotKeyboardFocus);
                textBoxEnterSDG.LostKeyboardFocus += new KeyboardFocusChangedEventHandler(tb_LostKeyboardFocus);


                // hint in textBoxWordPath
                textBoxWordPath.Foreground = Brushes.Gray;
                textBoxWordPath.Text = "Enter word file path...";
                textBoxWordPath.GotKeyboardFocus += new KeyboardFocusChangedEventHandler(tb_GotKeyboardFocus);
                textBoxWordPath.LostKeyboardFocus += new KeyboardFocusChangedEventHandler(tb_LostKeyboardFocus);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }




        private void openConnection()
        {
            connection = new OracleConnection();
            List<string> conStringParts = _ntlsCon.GetADOConnectionString().Split(';').ToList();
            string[] conStringValues = new string[] { "data", "user", "password" };
            List<string> conString = new List<string>();

            foreach (string part in conStringParts)
            {
                foreach (string value in conStringValues)
                {
                    if (part.ToLower().Contains(value))
                    {
                        conString.Add(part);
                    }
                }
            }

            connection.ConnectionString = String.Join(";", conString);
            connection.Open();
        }

        private void buttonSearchSDG_Click(object sender, RoutedEventArgs e)
        {
            listBoxInspectionPlan.Items.Clear();
            listBoxResults.Items.Clear();
            if (!textBoxEnterSDG.Text.Equals(string.Empty) && !textBoxEnterSDG.Text.Equals("Search SDG..."))
            {
                try
                {
                    currentSdg = dal.FindBy<SDG>(s => s.NAME == textBoxEnterSDG.Text).FirstOrDefault();

                    if (currentSdg != null)
                    {
                        int inspectionID = Convert.ToInt32(currentSdg.INSPECTION_PLAN_ID);
                        string inspectionName = dal.FindBy<INSPECTION_PLAN>(iplan => iplan.INSPECTION_PLAN_ID == inspectionID).FirstOrDefault().NAME;
                        List<int> rolesIdToInspect = new List<int>(); // roles needed 
                        List<INSPECTION_ENTRY> rolesNeedForInspection = dal.FindBy<INSPECTION_ENTRY>(ie => ie.INSPECTION_PLAN_ID == inspectionID).ToList();
                        List<INSPECTION_LOG> signs = dal.FindBy<INSPECTION_LOG>(il => il.TABLE_KEY == currentSdg.SDG_ID).ToList();


                        foreach (INSPECTION_ENTRY ie in rolesNeedForInspection)
                        {
                            rolesIdToInspect.Add(Convert.ToInt32(ie.ROLE_ID));
                        }

                        currentInspection = new inspectionPlan(inspectionID, rolesIdToInspect);

                        foreach (INSPECTION_LOG il in signs)
                        {
                            if (currentInspection.rolesToSign.ContainsKey(Convert.ToInt32(il.ROLE_ID)))
                            {
                                currentInspection.rolesToSign[Convert.ToInt32(il.ROLE_ID)] = true;
                            }
                        }


                        listBoxInspectionPlan.Items.Add("Inspection Plan: " + inspectionName + Environment.NewLine);

                        foreach (int roleID in currentInspection.rolesToSign.Keys)
                        {
                            string roleName = dal.FindBy<INSPECTION_ENTRY>(ie => ie.ROLE_ID == roleID).FirstOrDefault().LIMS_ROLE.NAME;
                            listBoxInspectionPlan.Items.Add(roleName + ": " + (currentInspection.rolesToSign[roleID] ? "Sign" : "Didn't sign"));
                        }

                        isAuthorizedToSign();
                    }


                    //-----------------------------Load results of the entered SDG------------------------------------------

                    buildListBoxResults();

                    OpenRTF();

                }
                catch (Exception ex)
                {
                    MessageBox.Show("not a valid SDG." + Environment.NewLine + ex.Message);
                }
            }
        }

        private void buildListBoxResults()
        {
            try
            {
                foreach (SAMPLE sample in currentSdg.SAMPLEs)
                {
                    foreach (ALIQUOT aliquot in sample.ALIQUOTs)
                    {
                        foreach (TEST test in aliquot.TESTs)
                        {
                            foreach (RESULT result in test.RESULTs)
                            {
                                listBoxResults.Items.Add(result.NAME);
                                currentResultIDs.Add(result.RESULT_ID);
                            }
                        }
                    }
                }

                if (listBoxResults.Items.Count == 0)
                {
                    listBoxResults.Items.Add("SDG has no results");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void buttonAuthorize_Click(object sender, RoutedEventArgs e)
        {
            bool needToSign = false;

            try
            {
                if (currentSdg != null)
                {
                    needToSign = isAuthorizedToSign();

                    if (needToSign)
                    {
                        if (currentSdg.STATUS == "P")
                        {
                            currentSdg.STATUS = "C";
                            dal.SaveChanges();
                        }

                        dal.RefreshAll();
                        currentSdg.STATUS = "A";
                        dal.SaveChanges();
                        dal.RefreshAll();
                    }
                }

                MessageBox.Show(needToSign ? "Authorize." : "Role doesn't need to sign.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.InnerException.InnerException.Message);
            }
        }

        private bool isAuthorizedToSign()
        {
            bool needToSign = false;
            buttonAuthorize.IsEnabled = false;

            if (currentSdg != null)
            {
                foreach (int roleID in currentInspection.rolesToSign.Keys)
                {
                    if (_ntlsUser.GetRoleId() == (double)roleID)
                    {
                        if (!currentInspection.rolesToSign[roleID])
                        {
                            needToSign = true;
                            buttonAuthorize.IsEnabled = true;
                        }
                        break;
                    }
                }
            }

            return needToSign;
        }

        private void tb_GotKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            TextBox textBox = sender as TextBox;

            if (textBox != null)
            {
                //If nothing has been entered yet.
                if (textBox.Foreground == Brushes.Gray)
                {
                    textBox.Text = "";
                    textBox.Foreground = Brushes.Black;
                }
            }
        }

        private void tb_LostKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            TextBox textBox = sender as TextBox;

            //Make sure sender is the correct Control.
            if (textBox != null)
            {
                //If nothing was entered, reset default text.
                if (textBox.Text.Trim().Equals(""))
                {
                    textBox.Foreground = Brushes.Gray;
                    textBox.Text = textBox.Name == textBoxEnterSDG.Name ? "Search SDG..." : "Enter word file path...";
                }
            }
        }

        public string RtfToPlainText(string rtf)
        {
            var flowDocument = new FlowDocument();
            var textRange = new TextRange(flowDocument.ContentStart, flowDocument.ContentEnd);

            using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(rtf ?? string.Empty)))
            {
                textRange.Load(stream, DataFormats.Rtf);
            }

            return textRange.Text;
        }

        private void OpenRTF()
        {
            string item = listBoxResults.SelectedItem as string;
            RTF_RESULT rtfMicro = null;
            RTF_RESULT rtfMacro = null;
            RTF_RESULT rtfDiagnosis = null;

            if (true)
            {
                try
                {
                    dal.RefreshAll();
                    RESULT resultMicro = dal.FindBy<RESULT>(r => r.NAME.ToLower().Contains("micro") && currentResultIDs.Contains(r.RESULT_ID)).FirstOrDefault();
                    RESULT resultMacro = dal.FindBy<RESULT>(r => r.NAME.ToLower().Contains("macro") && currentResultIDs.Contains(r.RESULT_ID)).FirstOrDefault();
                    RESULT resultDiagnosis = dal.FindBy<RESULT>(r => r.NAME.ToLower().Contains("diagnosis") && currentResultIDs.Contains(r.RESULT_ID)).FirstOrDefault();

                    try
                    {
                        rtfMicro = dal.FindBy<RTF_RESULT>(r => r.RTF_RESULT_ID == resultMicro.RESULT_ID).FirstOrDefault();
                        rtfMacro = dal.FindBy<RTF_RESULT>(r => r.RTF_RESULT_ID == resultMacro.RESULT_ID).FirstOrDefault();
                        rtfDiagnosis = dal.FindBy<RTF_RESULT>(r => r.RTF_RESULT_ID == resultDiagnosis.RESULT_ID).FirstOrDefault();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("result not found.");
                    }

                    currentResultID = resultMicro.RESULT_ID;
                    currentResultName = resultMicro.NAME;


                    if (rtfMacro != null)
                    {
                        richTextMacro.SetRtf(rtfMacro.RTF_TEXT);
                    }
                    
                    if (rtfDiagnosis != null)
                    {
                        richTextDiagnos.SetRtf(rtfDiagnosis.RTF_TEXT);
                    }

                    if (rtfMicro != null)
                    {
                        currentRtfPath = (string.Format(@"C:\Users\orsh\Desktop\{0}_{1}.rtf", currentResultID, currentResultName)).Replace(" ", "_");


                        // need to set the document with content!!
                        //document.SaveAs(currentRtfPath, Word.WdSaveFormat.wdFormatRTF);

                        richTextMicro.SetRtf(rtfMicro.RTF_TEXT.Trim());
                        richTextMicro.AppendText(Environment.NewLine + Environment.NewLine);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }


        // button to open micro in word process
        private void extraBtn_Click()
        {
            saveToDB = true;

            if (!string.IsNullOrEmpty(richTextMicro.GetRtf()))
            {
                saveRtfTextAsWord(richTextMicro.GetRtf(), currentRtfPath);
                deleteFirstRow(currentRtfPath);
                openFile(currentRtfPath);
            }
        }

        private void buttonOpenExistWord_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string filePath = textBoxWordPath.Text;

                if (!string.IsNullOrEmpty(filePath) && !filePath.Equals("Enter word file path..."))
                {
                    saveToDB = false;
                    currentRtfPath = string.Empty;
                    openFile(filePath);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void openFile(string filePath, bool createNewFile = false)
        {
            try
            {
                // in case createNewFile = true --> the content of filePath will be the location to save the file
                if (!createNewFile)
                {
                    if (File.Exists(filePath))
                    {
                        Process wordProcess = Process.Start(filePath);
                        wordProcess.WaitForExit();
                        wordProcess.Dispose();
                        ProcessExited();
                    }
                    else
                    {
                        MessageBox.Show("Invalid path.");
                    }
                }
                else
                {
                    wordFile = new Word.Application();
                    document = wordFile.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                    currentRtfPath = (filePath + string.Format(@"\{0}_{1}.rtf", currentResultID, currentResultName)).Replace(" ", "_");
                    document.SaveAs(currentRtfPath, Word.WdSaveFormat.wdFormatRTF);

                    document.Close(ref oMissing, ref oMissing, ref oMissing);
                    wordFile.Quit(ref oMissing, ref oMissing, ref oMissing);

                    openFile(currentRtfPath);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ProcessExited()
        {
                try
                {
                    if (saveToDB)
                    {
                        updateResultRTF(currentResultID);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            saveToDB = false;
        }

        private void buttonDocxToRtf_Click(object sender, RoutedEventArgs e)
        {
            string filePath = textBoxWordPath.Text;

            try
            {
                if (!string.IsNullOrEmpty(filePath) && !filePath.Equals("Enter word file path..."))
                {
                    if (File.Exists(filePath))
                    {
                        if (filePath.Contains(".docx"))
                        {
                            wordFile = new Word.Application();
                            document = wordFile.Documents.Open(filePath);
                            filePath = filePath.Replace(".docx", ".rtf");

                            document.SaveAs(filePath, Word.WdSaveFormat.wdFormatRTF);

                            MessageBox.Show("File saved as RTF");
                        }
                        else
                        {
                            MessageBox.Show("Not docx file.");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Invalid path.");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                document.Close(ref oMissing, ref oMissing, ref oMissing);
                wordFile.Quit(ref oMissing, ref oMissing, ref oMissing);
            }

        }

        private string readRtfFile(string filePath)
        {
            try
            {
                string rtf = File.ReadAllText(filePath);
                return rtf;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return string.Empty;
            }
        }

        private string FormatTextAsRTF(string rtfText)
        {
            System.Windows.Forms.RichTextBox rtf = new System.Windows.Forms.RichTextBox();
            rtf.Text = rtfText;
            return rtf.Rtf;
        }

        private void updateResultRTF(long resultID)
        {
            try
            {
                using (OracleCommand cmd = new OracleCommand())
                {
                    cmd.Connection = connection;
                    cmd.CommandType = CommandType.Text;

                    RTF_RESULT hasRTF = dal.FindBy<RTF_RESULT>(rtf => rtf.RTF_RESULT_ID == resultID).FirstOrDefault();

                    string rtfString = readRtfFile(currentRtfPath);

                    saveRtfTextAsWord(rtfString, currentRtfPath);
                    deleteFirstRow(currentRtfPath);


                    string text = rtfFileToText(currentRtfPath);
                    text = text.Substring(0, text.Length > 4000 ? 4000 : text.Length);

                    int rowsAffected;

                    if (hasRTF != null)
                    {
                        cmd.CommandText = string.Format("UPDATE rtf_result SET rtf_text=:TEXT WHERE rtf_result_id={0}", resultID);
                        cmd.Parameters.Add("TEXT", OracleType.Clob).Value = rtfString;
                        rowsAffected = cmd.ExecuteNonQuery();

                        RESULT result = dal.FindBy<RESULT>(r => r.RESULT_ID == resultID).FirstOrDefault();
                        result.FORMATTED_RESULT = text;
                    }
                    else
                    {
                        cmd.CommandText = string.Format("INSERT INTO rtf_result (rtf_result_id, rtf_text) VALUES ('{0}', :TEXT)", resultID);
                        cmd.Parameters.Add("TEXT", OracleType.Clob).Value = rtfString;
                        rowsAffected = cmd.ExecuteNonQuery();

                        RESULT result = dal.FindBy<RESULT>(r => r.RESULT_ID == resultID).FirstOrDefault();
                        result.FORMATTED_RESULT = text;
                    }

                    dal.SaveChanges();
                    dal.RefreshAll();
                }
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("ORA-00942"))
                {
                    MessageBox.Show("Logged in user can't edit database");
                }
                else
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        //private void insertToResultRTF(long resultID, RTF_RESULT rtf, string text)
        //{
        //    //text = text.Substring(0, text.Length > 4000 ? 4000 : text.Length);
        //    rtf.RTF_RESULT_ID = resultID;
        //    rtf.RTF_TEXT = text;

        //    dal.SaveChanges();
        //}

        private string rtfFileToText(string RtfFilePath)
        {
            wordFile = new Word.Application();
            object encoding = Microsoft.Office.Core.MsoEncoding.msoEncodingUTF8;
            document = wordFile.Documents.Open(RtfFilePath, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref encoding, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);

            // Loop through all words in the document.
            int count = document.Words.Count;
            string text = string.Empty;

            for (int i = 1; i <= count; i++)
            {
                text += document.Words[i].Text;
            }

            document.Close(ref oMissing, ref oMissing, ref oMissing);
            wordFile.Quit(ref oMissing, ref oMissing, ref oMissing);

            return text;
        }

        private void buttonUpdateDB_Click(object sender, RoutedEventArgs e)
        {
            string filePath = textBoxWordPath.Text;

            if (!string.IsNullOrEmpty(filePath) && !filePath.Equals("Enter word file path..."))
            {
                //updateResultRTF(filePath)
            }
        }

        internal static void saveRtfTextAsWord(string rtf, string filePathAndName)
        {
            Document doc= null;

            try
            {
                doc = new Document();

                TextReader tr = new StringReader(rtf);

                doc.LoadRtf(tr);
                doc.SaveToFile(filePathAndName, FileFormat.Rtf);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally 
            {
                if(doc != null)
                    doc.Close();
            }

        }

        private void deleteFirstRow(string filePath)
        {
            try
            {
                wordFile = new Word.Application();
                document = wordFile.Documents.Open(filePath);
                var range = document.Content;
                if (range.Find.Execute("Evaluation Warning: The document was created with Spire.Doc for .NET."))
                {
                    range.Expand(wdUnits.WdUnits.wdSentence); // or change to .wdLine or .wdSentence or .wdParagraph
                    range.Delete();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if(document != null)
                    document.Close(ref oMissing, ref oMissing, ref oMissing);
                if(wordFile != null)
                    wordFile.Quit(ref oMissing, ref oMissing, ref oMissing);
            }


        }

        public bool CloseQuery()
        {
            if (connection != null)
            {
                connection.Close();
                connection = null;
            }

            if (dal != null)
                dal.Close();

            return true;
        }

        private void textBoxEnterSDG_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter || e.Key == Key.Tab)
            {
                buttonSearchSDG_Click(null, null);
            }
        }
    }

    class inspectionPlan
    {
        public int InspectionID { get; set; }
        public Dictionary<int, bool> rolesToSign { get; set; }

        public inspectionPlan(int i_InspectionID, List<int> i_rolesID)
        {
            this.InspectionID = i_InspectionID;
            rolesToSign = new Dictionary<int, bool>();
            initRoles(i_rolesID);
        }

        private void initRoles(List<int> i_rolesID) 
        {
            foreach (int id in i_rolesID)
            {
                rolesToSign.Add(id, false);
            }
        }
    }

    //class role
    //{
    //    public int roleID { get; set; }

    //    public role(int i_roleID)
    //    {
    //        this.roleID = i_roleID;
    //    }
    //}
}
