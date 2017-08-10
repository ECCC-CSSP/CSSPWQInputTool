using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Collections.Specialized;
using System.Data.Entity.Validation;
using System.Diagnostics;
using CSSPFCFormWriterDLL.Services;
using CSSPLabSheetParserDLL.Services;
using CSSPModelsDLL.Models;
using CSSPEnumsDLL.Enums;

namespace CSSPWQInputTool
{
    public partial class CSSPWQInputToolForm : Form
    {
        #region Variables
        private string ApprovalSupervisorInitials = "";
        private bool IsSaving = false;
        private string SpaceStr = "                 ";
        private bool IsTest = false;
        private string TestFile = @"C:\Users\leblancc\Documents\CSSP\SamplingPlan_Testing.txt";
        private string r = "OgW2S3EHhQ(6!Z$odV7eAGnim/#YIClk9vF&1@5xDUa)wPLu*BN.t,c8%JRMbK^yqzXpfTj4sr0:d";
        private bool InternetConnection = false;
        private string FormTitle = "";
        private List<CSSPWQInputParam> csspWQInputParamList = new List<CSSPWQInputParam>();
        private CSSPWQInputApp csspWQInputApp = new CSSPWQInputApp();
        private Color ButBackColor = Color.Black;
        private CSSPWQInputTypeEnum csspWQInputTypeCurrent = CSSPWQInputTypeEnum.Subsector;
        private CSSPWQInputSheetTypeEnum csspWQInputSheetType = CSSPWQInputSheetTypeEnum.A1;
        private string RootCurrentPath = @"C:\CSSPLabSheets\";
        private string CurrentPath = "";
        private string NameCurrent = "";
        private int TVItemIDCurrent = 0;
        private string YearMonthDayCurrent = "";
        private string RunNumberCurrent = "";
        private CSSPWQInputParam CSSPWQInputParamCurrent = new CSSPWQInputParam();
        private DataGridViewCellStyle dataGridViewCellStyleDefault = new DataGridViewCellStyle();
        private DataGridViewCellStyle dataGridViewCellStyleEdit = new DataGridViewCellStyle();
        private DataGridViewCellStyle dataGridViewCellStyleEditRowCell = new DataGridViewCellStyle();
        private DataGridViewCellStyle dataGridViewCellStyleEditError = new DataGridViewCellStyle();
        private List<CSSPMPNTable> csspMPNTableList = new List<CSSPMPNTable>();
        private bool InLoadingFile = false;
        private string SamplingPlanName = "";
        private bool NoUpdate = false;
        private int TideToTryIndex = 0;
        private bool panelAppInputIsVisible = false;
        private Color ControlBackColor;
        private Color TextBoxBackColor;
        private Color DataGridViewCSSPBackgroundColor;
        private int VersionOfSamplingPlanFile = 0;
        private string LabSheetType = "";
        private string SamplingPlanType = "";
        private string SampleType = "";
        private int VersionOfResultFile = 1;
        private Panel CurrentPanel = null;
        private string Initials = "";
        private bool IsOnDailyDuplicate = false;
        private bool AppIsWide = false;
        private bool FileListOnlyChangedAndRejected = false;
        private bool FileListViewTotalColiformLabSheets = false;
        private List<string> AllowableTideString = new List<string>()
        {
            "/", "--", "HR", "HT", "HF", "MR", "MT", "MF", "LR", "LT", "LF",
        };
        private StringBuilder sbLog = new StringBuilder();
        private List<string> PossibleLabSheetFileNamesTxt = new List<string>() { "_C.txt", "_S.txt", "_E.txt", "_R.txt", "_A.txt", "_F.txt" };
        private List<string> PossibleLabSheetFileNamesDocx = new List<string>() { "_C.docx", "_S.docx", "_E.docx", "_R.docx", "_A.docx", "_F.docx" };
        private List<string> LogHistory = new List<string>();
        private string Server = "(Server)";
        private string Local = "(Local)";
        private bool WaitingForUserAction = true;
        private bool UserActionFileArchiveCopy = false;
        private bool UserActionFileArchiveSkip = false;
        private bool UserActionFileArchiveCancel = false;
        #endregion Variables

        #region Properties
        public LabSheetA1Sheet labSheetA1Sheet { get; set; }
        public CSSPFCFormWriter csspFCFormWriter { get; set; }
        public CSSPLabSheetParser csspLabSheetParser { get; set; }
        public CultureInfo currentCulture { get; set; }
        public CultureInfo currentUICulture { get; set; }
        #endregion Properties

        #region Contructors
        public CSSPWQInputToolForm()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-CA");
            Thread.CurrentThread.CurrentUICulture = new CultureInfo("en-CA");

            currentCulture = Thread.CurrentThread.CurrentCulture;
            currentUICulture = Thread.CurrentThread.CurrentUICulture;

            InitializeComponent();
            FormTitle = this.Text;
            ControlBackColor = lblSampleCrewInitials.BackColor;
            TextBoxBackColor = textBoxSampleCrewInitials.BackColor;
            DataGridViewCSSPBackgroundColor = dataGridViewCSSP.BackgroundColor;
            lblSamplingPlanFileName.Text = "";
            panelPassword.Dock = DockStyle.Fill;
            panelApp.Dock = DockStyle.Fill;
            panelButtonBar.Visible = false;
            panelPassword.BringToFront();
            CurrentPanel = panelPassword;
            panelPasswordCenter.Location = new Point(panelPassword.Width / 2 - panelPasswordCenter.Size.Width / 2, panelPassword.Height / 2 - panelPasswordCenter.Size.Height / 2);
            panelPasswordCenter.Anchor = AnchorStyles.None;
            panelAppInput.Dock = DockStyle.Fill;
            panelAppInputFiles.Dock = DockStyle.Fill;
            panelSendToServerCompare.Dock = DockStyle.Fill;
            FillCSSPMPNTable();
            webBrowserCSSP.ScriptErrorsSuppressed = true;

            csspFCFormWriter = new CSSPFCFormWriter(LanguageEnum.en, "Empty for now");
            csspLabSheetParser = new CSSPLabSheetParser();
            dateTimePickerArchiveFilterFrom.Value = new DateTime(DateTime.Now.Year, 1, 1);
            dateTimePickerArchiveFilterTo.Value = new DateTime(DateTime.Now.Year, 12, 31);
            RadioButtonBathNumberChanged();
        }
        #endregion Contructors

        #region Events
        #region Events Form
        private void CSSPWQInputToolForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            while (IsSaving == true)
            {
                Application.DoEvents();
            }
        }
        private void CSSPWQInputToolForm_SizeChanged(object sender, EventArgs e)
        {
            ResetBottomPanels();
        }
        #endregion Events Form
        #region Events Buttons
        private void butApprove_Click(object sender, EventArgs e)
        {
            Approve();
        }
        private void butArchive_Click(object sender, EventArgs e)
        {
            SetupAppInputFiles();
        }
        private void butBrowseSamplingPlanFile_Click(object sender, EventArgs e)
        {
            OpenSamplingPlanFile(IsTest);
        }
        private void butCancelSendToServer_Click(object sender, EventArgs e)
        {
            CancelSendToServer();
        }
        private void butContinueSendToServer_Click(object sender, EventArgs e)
        {
            ContinueSendToServer();
        }
        private void butChangeDate_Click(object sender, EventArgs e)
        {
            ChangeDate();
        }
        private void butChangeDateCancel_Click(object sender, EventArgs e)
        {
            lblChangeDateError.Text = "";
            panelAppInput.BringToFront();
            CurrentPanel = panelAppInput;
        }
        private void butCreateFile_Click(object sender, EventArgs e)
        {
            checkBoxViewTotalColiformLabSheets.Checked = false;
            FileListViewTotalColiformLabSheets = false;
            FileInfo fi = new FileInfo(CurrentPath);
            if (NameCurrent.Contains(" "))
            {
                fi = new FileInfo(CurrentPath + YearMonthDayCurrent.Substring(0, 4) + @"\" + NameCurrent.Substring(0, NameCurrent.IndexOf(" ")) + "_" + YearMonthDayCurrent + "_" + csspWQInputSheetType.ToString() + "_R" + RunNumberCurrent + "_S.txt");
            }
            else
            {
                fi = new FileInfo(CurrentPath + YearMonthDayCurrent.Substring(0, 4) + @"\" + NameCurrent + YearMonthDayCurrent + "_" + csspWQInputSheetType.ToString() + "_R" + RunNumberCurrent + "_S.txt");
            }

            if (fi.Exists)
            {
                lblStatus.Text = "File already exist ... " + fi.FullName;
                return;
            }
            else
            {
                fi = new FileInfo(fi.FullName.Replace("_S.txt", "_C.txt"));
                if (fi.Exists)
                {
                    lblStatus.Text = "File already exist ... " + fi.FullName;
                }
                else
                {
                    if (!fi.Directory.Exists)
                    {
                        try
                        {
                            fi.Directory.Create();
                        }
                        catch (Exception ex)
                        {
                            lblStatus.Text = "Error: " + ex.Message + (ex.InnerException != null ? " --- " + ex.InnerException.Message : "");
                            return;
                        }
                    }
                    StreamWriter sw = fi.CreateText();
                    sw.Close();
                }
            }
            sbLog = new StringBuilder();
            sbLog.AppendLine("________________________________");
            sbLog.AppendLine("Log");
            textBoxTides.Text = "-- / --";
            dateTimePickerSalinitiesReadDate.Value = dateTimePickerRun.Value.AddDays(1);
            dateTimePickerResultsReadDate.Value = dateTimePickerRun.Value.AddDays(1);
            dateTimePickerResultsRecordedDate.Value = dateTimePickerRun.Value.AddDays(1);
            textBoxDailyDuplicatePrecisionCriteria.Text = csspWQInputApp.DailyDuplicatePrecisionCriteria.ToString();
            textBoxIntertechDuplicatePrecisionCriteria.Text = csspWQInputApp.IntertechDuplicatePrecisionCriteria.ToString();
            UpdatePanelApp();
            lblSupervisorInitials.Text = "";
            lblApprovalDate.Text = "";
            SaveInfoOnLocalMachine(false);
            ReadFileFromLocalMachine();
            butCreateFile.Visible = false;
            NoUpdate = false;
            UpdatePanelApp();
        }
        private void butGetTides_Click(object sender, EventArgs e)
        {
            TideToTryIndex = 0;
            textBoxTides.BackColor = TextBoxBackColor;
            textBoxTides.Text = "Loading ...";
            timerGetTides.Enabled = true;
        }
        private void butFail_Click(object sender, EventArgs e)
        {
            ToggleFailFileName();
            UpdatePanelApp();
        }
        private void butFileArchiveCancel_Click(object sender, EventArgs e)
        {
            UserActionFileArchiveCancel = true;
        }
        private void butFileArchiveCopy_Click(object sender, EventArgs e)
        {
            UserActionFileArchiveCopy = true;
        }
        private void butFileArchiveSkip_Click(object sender, EventArgs e)
        {
            UserActionFileArchiveSkip = true;
        }
        private void butLogoff_Click(object sender, EventArgs e)
        {
            panelAccessCode.Visible = false;
            SamplingPlanName = "";
            lblSamplingPlanFileName.Text = "";
            panelPassword.BringToFront();
            CurrentPanel = panelPassword;
            panelButtonBar.Visible = false;
            textBoxAccessCode.Text = "";
            textBoxApprovalCode.Text = "";
            textBoxInitials.Text = "";
            textBoxInitials.Focus();
        }
        private void butOpen_Click(object sender, EventArgs e)
        {
            OpenFileName();
        }
        private void butSalinitySameDay_Click(object sender, EventArgs e)
        {
            if (dateTimePickerSalinitiesReadDate.Value == dateTimePickerRun.Value)
            {
                dateTimePickerSalinitiesReadDate.Value = dateTimePickerRun.Value.AddDays(1);
                butSalinitySameDay.Text = "Same Day";
            }
            else
            {
                dateTimePickerSalinitiesReadDate.Value = dateTimePickerRun.Value;
                butSalinitySameDay.Text = "Next Day";
            }
        }
        private void butSendToServer_Click(object sender, EventArgs e)
        {
            if (lblFilePath.Text.EndsWith("_S.txt"))
            {
                MessageBox.Show("Can't post lab sheet that has already been sent or has the status of sent [ends with _S.txt].", "Error");
                return;
            }

            if (!EverythingEntered())
            {
                return;
            }

            TrySendingToServer();
        }
        private void butSyncArchives_Click(object sender, EventArgs e)
        {
            TryToSyncArchive();
        }
        private void butGetLabSheetsStatus_Click(object sender, EventArgs e)
        {
            butGetLabSheetsStatus.Text = "Working ...";
            GetLabSheetsStatus();
            butGetLabSheetsStatus.Text = "Get lab sheets status";
            SetupAppInputFiles();
        }
        private void butViewFCForm_Click(object sender, EventArgs e)
        {
            CreateWordDoc();
            processCSSP.StartInfo.FileName = lblFilePath.Text.Replace(".txt", ".docx");
            processCSSP.Start();
            if (butViewFCForm.ForeColor == Color.Black)
            {
                butViewFCForm.Text = "View FC Form";
                lblStatus.Text = "Created and loaded [" + processCSSP.StartInfo.FileName + "]";
            }
        }
        #endregion Events Buttons
        #region Events checkBox
        private void checkBox2Coolers_CheckedChanged(object sender, EventArgs e)
        {
            Modifying();
            if (checkBox2Coolers.Checked)
            {
                checkBox2Coolers.ForeColor = Color.Green;
                textBoxTCField2.Visible = true;
                textBoxTCLab2.Visible = true;
                AddLog("Two Coolers", true.ToString());
            }
            else
            {
                checkBox2Coolers.ForeColor = Color.Black;
                textBoxTCField2.Text = "";
                textBoxTCLab2.Text = "";
                textBoxTCField2.Visible = false;
                textBoxTCLab2.Visible = false;
                AddLog("Two Coolers", false.ToString());
            }
        }
        private void checkBoxIncubationStartSameDay_CheckedChanged(object sender, EventArgs e)
        {
            Modifying();
            if (checkBoxIncubationStartSameDay.Checked)
            {
                AddLog("Incubation Start Same Day", true.ToString());
                dateTimePickerResultsReadDate.Value = dateTimePickerRun.Value.AddDays(1);
                dateTimePickerResultsRecordedDate.Value = dateTimePickerRun.Value.AddDays(1);
            }
            else
            {
                DialogResult dialogResult = MessageBox.Show("Sample hold time exceeds recommended 8hrs. \r\n\r\n" +
                    "Have you obtained supervisor approval?\r\n\r\n" +
                    "If yes: Make sure you indicate the name of the supervisor who gave approval in the Run Comment section", "Supervisor permission required", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    AddLog("Incubation Start Same Day", false.ToString());
                    dateTimePickerResultsReadDate.Value = dateTimePickerRun.Value.AddDays(2);
                    dateTimePickerResultsRecordedDate.Value = dateTimePickerRun.Value.AddDays(2);
                }
                else
                {
                    checkBoxIncubationStartSameDay.Checked = true;
                }
            }
        }
        private void checkBoxOnlyChangedAndRejected_CheckedChanged(object sender, EventArgs e)
        {
            OnlyChangedAndRejected();
        }
        private void checkBoxViewTotalColiformLabSheets_CheckedChanged(object sender, EventArgs e)
        {
            ViewTotalColiformLabSheets();
        }

        private void ViewTotalColiformLabSheets()
        {
            if (checkBoxViewTotalColiformLabSheets.Checked == true)
            {
                FileListViewTotalColiformLabSheets = true;
            }
            else
            {
                FileListViewTotalColiformLabSheets = false;
            }
            LoadFileList();
        }
        #endregion checkBox
        #region Events comboBoxSubsectorNames
        private void CheckIfFileExist()
        {
            bool ShouldUpdatePanelApp = true;
            FileItem fileItem = (FileItem)comboBoxSubsectorNames.SelectedItem;
            if (fileItem.TVItemID == 0)
            {
                ShouldUpdatePanelApp = false;
            }
            int RunNumber = 0;
            if (!int.TryParse((string)comboBoxRunNumber.SelectedItem, out RunNumber))
            {
                ShouldUpdatePanelApp = false;
            }
            if (RunNumber == 0)
            {
                ShouldUpdatePanelApp = false;
            }
            if (NoUpdate)
            {
                ShouldUpdatePanelApp = false;
            }

            if (ShouldUpdatePanelApp)
            {
                UpdatePanelApp();
            }
            else
            {
                lblFilePath.Text = "";
                butCreateFile.Visible = false;
                SetupAppInputFiles();
            }
        }
        private void comboBoxSubsectorNames_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckIfFileExist();
            //if (comboBoxSubsectorNames.SelectedIndex == 0)
            //{
            //    butGetLabSheetsStatus.Enabled = false;
            //}
            //else
            //{
            //    butGetLabSheetsStatus.Enabled = true;
            //}
        }
        #endregion Events comboBoxSubsectorNames
        #region Events comboBoxFile
        private void comboBoxFileFilter_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadFileList();
        }
        private void comboBoxFileSubsector_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadFileList();
        }
        private void comboBoxFileYear_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadFileList();
        }
        #endregion Events comboBoxFile
        #region comboBoxRunNumber
        private void comboBoxRunNumber_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckIfFileExist();
        }
        #endregion comboBoxRunNumber
        #region Events dataGridViewCSSP
        private void dataGridViewCSSP_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            dataGridViewCSSP.BackgroundColor = DataGridViewCSSPBackgroundColor;
            if (csspWQInputSheetType == CSSPWQInputSheetTypeEnum.LTB)
            {
                ValidateCellLTB(e);
            }
            else if (csspWQInputSheetType == CSSPWQInputSheetTypeEnum.EC)
            {
                ValidateCellEC(e);
            }
            else
            {
                ValidateCellA1(e.ColumnIndex, e.RowIndex);
            }
            CalculateDuplicate();
            Modifying();
        }
        private void dataGridViewCSSP_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            int SiteColumn = 1;
            int TimeColumn = 2;
            int SalinityColumn = 7;
            int TemperatureColumn = 8;
            int ProcessByColumn = 9;
            int SampleTypeColumn = 10;
            if (dataGridViewCSSP.CurrentCell != null)
            {
                if (csspWQInputSheetType == CSSPWQInputSheetTypeEnum.A1)
                {
                    if (e.RowIndex > 0 && (e.ColumnIndex == TimeColumn || e.ColumnIndex == SalinityColumn || e.ColumnIndex == TemperatureColumn || e.ColumnIndex == ProcessByColumn))
                    {
                        string cellStr = dataGridViewCSSP.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                        if (string.IsNullOrWhiteSpace(cellStr))
                        {
                            if (e.ColumnIndex == TimeColumn || e.ColumnIndex == SalinityColumn || e.ColumnIndex == TemperatureColumn)
                            {
                                for (int i = e.RowIndex - 1; i >= 0; i--)
                                {
                                    string siteCurrentStr = dataGridViewCSSP.Rows[e.RowIndex].Cells[SiteColumn].Value.ToString();
                                    string siteParentStr = dataGridViewCSSP.Rows[i].Cells[SiteColumn].Value.ToString();

                                    if (siteCurrentStr == siteParentStr)
                                    {
                                        if (dataGridViewCSSP.Rows[e.RowIndex].Cells[SampleTypeColumn].Value.ToString() == "DailyDuplicate"
                                            || dataGridViewCSSP.Rows[e.RowIndex].Cells[SampleTypeColumn].Value.ToString() == "IntertechDuplicate"
                                            || dataGridViewCSSP.Rows[e.RowIndex].Cells[SampleTypeColumn].Value.ToString() == "IntertechRead")
                                        {
                                            dataGridViewCSSP[e.ColumnIndex, e.RowIndex].Style.ForeColor = Color.Black;
                                            dataGridViewCSSP.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = dataGridViewCSSP.Rows[i].Cells[e.ColumnIndex].Value;
                                            Modifying();
                                            break;
                                        }
                                    }
                                }
                            }
                            else if (e.ColumnIndex == ProcessByColumn)
                            {
                                dataGridViewCSSP.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = dataGridViewCSSP.Rows[(e.RowIndex - 1)].Cells[e.ColumnIndex].Value;
                                Modifying();
                            }
                        }
                    }
                    switch (GetSampleType(dataGridViewCSSP.Rows[e.RowIndex].Cells[SampleTypeColumn].Value.ToString().Trim()))
                    {
                        case SampleTypeEnum.DailyDuplicate:
                            {
                                IsOnDailyDuplicate = true;
                                ResetBottomPanels();
                            }
                            break;
                        case SampleTypeEnum.Infrastructure:
                            {
                            }
                            break;
                        case SampleTypeEnum.IntertechDuplicate:
                            {
                                IsOnDailyDuplicate = true;
                                ResetBottomPanels();
                            }
                            break;
                        case SampleTypeEnum.IntertechRead:
                            {
                                IsOnDailyDuplicate = true;
                                ResetBottomPanels();
                            }
                            break;
                        case SampleTypeEnum.RainCMPRoutine:
                            {
                            }
                            break;
                        case SampleTypeEnum.RainRun:
                            {
                            }
                            break;
                        case SampleTypeEnum.ReopeningEmergencyRain:
                            {
                            }
                            break;
                        case SampleTypeEnum.ReopeningSpill:
                            {
                            }
                            break;
                        case SampleTypeEnum.Routine:
                            {
                                IsOnDailyDuplicate = false;
                                ResetBottomPanels();
                            }
                            break;
                        case SampleTypeEnum.Sanitary:
                            {
                            }
                            break;
                        case SampleTypeEnum.Study:
                            {
                            }
                            break;
                        default:
                            break;
                    }
                }
            }
        }
        private SampleTypeEnum GetSampleType(string SampleTypeText)
        {
            switch (SampleTypeText)
            {
                case "DailyDuplicate":
                    return SampleTypeEnum.DailyDuplicate;
                case "Infrastructure":
                    return SampleTypeEnum.Infrastructure;
                case "IntertechDuplicate":
                    return SampleTypeEnum.IntertechDuplicate;
                case "IntertechRead":
                    return SampleTypeEnum.IntertechRead;
                case "RainCMPRoutine":
                    return SampleTypeEnum.RainCMPRoutine;
                case "RainRun":
                    return SampleTypeEnum.RainRun;
                case "ReopeningEmergencyRain":
                    return SampleTypeEnum.ReopeningEmergencyRain;
                case "ReopeningSpill":
                    return SampleTypeEnum.ReopeningSpill;
                case "Routine":
                    return SampleTypeEnum.Routine;
                case "Sanitary":
                    return SampleTypeEnum.Sanitary;
                case "Study":
                    return SampleTypeEnum.Study;
                default:
                    return SampleTypeEnum.Error;
            }
        }
        #endregion dataGridViewCSSP
        #region Events dateTimePicker
        private void dateTimePickerChangeDate_ValueChanged(object sender, EventArgs e)
        {
            FileInfo fi = CanChangeDate();
        }
        private void dateTimePickerDuplicateDataEntryDate_ValueChanged(object sender, EventArgs e)
        {
            Modifying();
        }
        private void dateTimePickerResultsRecordedDate_ValueChanged(object sender, EventArgs e)
        {
        }
        private void dateTimePickerRun_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F2)
            {
                if (panelAppInput != CurrentPanel)
                {
                    lblStatus.Text = "You have to open a document before being able to change it's date.";
                    return;
                }
                if (!butSendToServer.Enabled)
                {
                    lblStatus.Text = "Please wait for the document to finished saving before changing the document date.";
                    return;
                }
                lblOldDateText.Text = dateTimePickerRun.Value.ToString("yyyy MMMM dd");
                lblChangeDateError.Text = "";
                panelChangeDateOfCurrentDoc.BringToFront();
                FileInfo fi = CanChangeDate();
            }
        }
        private void dateTimePickerRun_ValueChanged(object sender, EventArgs e)
        {
            CheckIfFileExist();
        }
        private void dateTimePickerSalinitiesReadDate_ValueChanged(object sender, EventArgs e)
        {
            Modifying();
        }
        private void dateTimePickerArchiveFilterFrom_ValueChanged(object sender, EventArgs e)
        {
            LoadFileList();
        }
        private void dateTimePickerArchiveFilterTo_ValueChanged(object sender, EventArgs e)
        {
            LoadFileList();
        }
        #endregion Events dateTimePickerRun
        #region Events Focus
        private void checkBox2Coolers_Leave(object sender, EventArgs e)
        {
            string CheckBoxText = (checkBox2Coolers.Checked ? "true" : "false");
            if (labSheetA1Sheet.TCHas2Coolers != CheckBoxText)
            {
                labSheetA1Sheet.TCHas2Coolers = CheckBoxText;
                AddLog("2 Coolers", CheckBoxText);
            }
        }
        private void dateTimePickerResultsReadDate_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.ResultsReadYear != dateTimePickerResultsReadDate.Value.Year.ToString()
                || labSheetA1Sheet.ResultsReadMonth != dateTimePickerResultsReadDate.Value.Month.ToString()
                || labSheetA1Sheet.ResultsReadDay != dateTimePickerResultsReadDate.Value.Day.ToString())
            {
                labSheetA1Sheet.ResultsReadYear = dateTimePickerResultsReadDate.Value.Year.ToString();
                labSheetA1Sheet.ResultsReadMonth = dateTimePickerResultsReadDate.Value.Month.ToString();
                labSheetA1Sheet.ResultsReadDay = dateTimePickerResultsReadDate.Value.Day.ToString();
                AddLog("Results Read Date", labSheetA1Sheet.ResultsReadYear + "\t" + labSheetA1Sheet.ResultsReadMonth + "\t" + labSheetA1Sheet.ResultsReadDay);
            }
        }
        private void dateTimePickerResultsRecordedDate_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.ResultsRecordedYear != dateTimePickerResultsRecordedDate.Value.Year.ToString()
                || labSheetA1Sheet.ResultsRecordedMonth != dateTimePickerResultsRecordedDate.Value.Month.ToString()
                || labSheetA1Sheet.ResultsRecordedDay != dateTimePickerResultsRecordedDate.Value.Day.ToString())
            {
                labSheetA1Sheet.ResultsRecordedYear = dateTimePickerResultsRecordedDate.Value.Year.ToString();
                labSheetA1Sheet.ResultsRecordedMonth = dateTimePickerResultsRecordedDate.Value.Month.ToString();
                labSheetA1Sheet.ResultsRecordedDay = dateTimePickerResultsRecordedDate.Value.Day.ToString();
                AddLog("Results Recorded Date", labSheetA1Sheet.ResultsRecordedYear + "\t" + labSheetA1Sheet.ResultsRecordedMonth + "\t" + labSheetA1Sheet.ResultsRecordedDay);
            }
        }
        private void dateTimePickerSalinitiesReadDate_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.SalinitiesReadYear != dateTimePickerSalinitiesReadDate.Value.Year.ToString()
               || labSheetA1Sheet.SalinitiesReadMonth != dateTimePickerSalinitiesReadDate.Value.Month.ToString()
               || labSheetA1Sheet.SalinitiesReadDay != dateTimePickerSalinitiesReadDate.Value.Day.ToString())
            {
                labSheetA1Sheet.SalinitiesReadYear = dateTimePickerSalinitiesReadDate.Value.Year.ToString();
                labSheetA1Sheet.SalinitiesReadMonth = dateTimePickerSalinitiesReadDate.Value.Month.ToString();
                labSheetA1Sheet.SalinitiesReadDay = dateTimePickerSalinitiesReadDate.Value.Day.ToString();
                AddLog("Results Salinities Date", labSheetA1Sheet.SalinitiesReadYear + "\t" + labSheetA1Sheet.SalinitiesReadMonth + "\t" + labSheetA1Sheet.SalinitiesReadDay);
            }
        }
        private void richTextBoxRunWeatherComment_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.RunWeatherComment != richTextBoxRunWeatherComment.Text)
            {
                labSheetA1Sheet.RunWeatherComment = richTextBoxRunWeatherComment.Text;
                AddLog("Run Weather Comment", richTextBoxRunWeatherComment.Text);
            }
        }
        private void richTextBoxRunComment_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.RunComment != richTextBoxRunComment.Text)
            {
                labSheetA1Sheet.RunWeatherComment = richTextBoxRunComment.Text;
                AddLog("Run Comment", richTextBoxRunComment.Text);
            }
        }
        private void textBoxControlBlank35_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.Blank35 != textBoxControlBlank35.Text)
            {
                labSheetA1Sheet.Blank35 = textBoxControlBlank35.Text;
                AddLog("Control Blank 35", textBoxControlBlank35.Text);
            }
        }
        private void textBoxControlBath1Blank44_5_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.Bath1Blank44_5 != textBoxControlBath1Blank44_5.Text)
            {
                labSheetA1Sheet.Bath1Blank44_5 = textBoxControlBath1Blank44_5.Text;
                AddLog("Control Bath 1 Blank 44.5", textBoxControlBath1Blank44_5.Text);
            }
        }
        private void textBoxControlBath2Blank44_5_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.Bath2Blank44_5 != textBoxControlBath2Blank44_5.Text)
            {
                labSheetA1Sheet.Bath2Blank44_5 = textBoxControlBath2Blank44_5.Text;
                AddLog("Control Bath 2 Blank 44.5", textBoxControlBath2Blank44_5.Text);
            }
        }
        private void textBoxControlBath3Blank44_5_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.Bath3Blank44_5 != textBoxControlBath3Blank44_5.Text)
            {
                labSheetA1Sheet.Bath3Blank44_5 = textBoxControlBath3Blank44_5.Text;
                AddLog("Control Bath 3 Blank 44.5", textBoxControlBath3Blank44_5.Text);
            }
        }
        private void textBoxControlLot_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.ControlLot != textBoxControlLot.Text)
            {
                labSheetA1Sheet.ControlLot = textBoxControlLot.Text;
                AddLog("Control Lot", textBoxControlLot.Text);
            }
        }
        private void textBoxControlNegative35_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.Negative35 != textBoxControlNegative35.Text)
            {
                labSheetA1Sheet.Negative35 = textBoxControlNegative35.Text;
                AddLog("Negative 35", textBoxControlNegative35.Text);
            }
        }
        private void textBoxControlBath1Negative44_5_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.Bath1Negative44_5 != textBoxControlBath1Negative44_5.Text)
            {
                labSheetA1Sheet.Bath1Negative44_5 = textBoxControlBath1Negative44_5.Text;
                AddLog("Bath 1 Negative 44.5", textBoxControlBath1Negative44_5.Text);
            }
        }
        private void textBoxControlBath2Negative44_5_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.Bath2Negative44_5 != textBoxControlBath2Negative44_5.Text)
            {
                labSheetA1Sheet.Bath2Negative44_5 = textBoxControlBath2Negative44_5.Text;
                AddLog("Bath 2 Negative 44.5", textBoxControlBath2Negative44_5.Text);
            }
        }
        private void textBoxControlBath3Negative44_5_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.Bath3Negative44_5 != textBoxControlBath3Negative44_5.Text)
            {
                labSheetA1Sheet.Bath3Negative44_5 = textBoxControlBath3Negative44_5.Text;
                AddLog("Bath 3 Negative 44.5", textBoxControlBath3Negative44_5.Text);
            }
        }
        private void textBoxControlNonTarget35_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.NonTarget35 != textBoxControlNonTarget35.Text)
            {
                labSheetA1Sheet.NonTarget35 = textBoxControlNonTarget35.Text;
                AddLog("Non Target 35", textBoxControlNonTarget35.Text);
            }
        }
        private void textBoxControlBath1NonTarget44_5_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.Bath1NonTarget44_5 != textBoxControlBath1NonTarget44_5.Text)
            {
                labSheetA1Sheet.Bath1NonTarget44_5 = textBoxControlBath1NonTarget44_5.Text;
                AddLog("Bath 1 Non Target 44.5", textBoxControlBath1NonTarget44_5.Text);
            }
        }
        private void textBoxControlBath2NonTarget44_5_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.Bath2NonTarget44_5 != textBoxControlBath2NonTarget44_5.Text)
            {
                labSheetA1Sheet.Bath2NonTarget44_5 = textBoxControlBath2NonTarget44_5.Text;
                AddLog("Bath 2 Non Target 44.5", textBoxControlBath2NonTarget44_5.Text);
            }
        }
        private void textBoxControlBath3NonTarget44_5_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.Bath3NonTarget44_5 != textBoxControlBath3NonTarget44_5.Text)
            {
                labSheetA1Sheet.Bath3NonTarget44_5 = textBoxControlBath3NonTarget44_5.Text;
                AddLog("Bath 3 Non Target 44.5", textBoxControlBath3NonTarget44_5.Text);
            }
        }
        private void textBoxControlPositive35_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.Positive35 != textBoxControlPositive35.Text)
            {
                labSheetA1Sheet.Positive35 = textBoxControlPositive35.Text;
                AddLog("Positive 35", textBoxControlPositive35.Text);
            }
        }
        private void textBoxControlBath1Positive44_5_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.Bath1Positive44_5 != textBoxControlBath1Positive44_5.Text)
            {
                labSheetA1Sheet.Bath1Positive44_5 = textBoxControlBath1Positive44_5.Text;
                AddLog("Bath 1 Positive 44.5", textBoxControlBath1Positive44_5.Text);
            }
        }
        private void textBoxControlBath2Positive44_5_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.Bath2Positive44_5 != textBoxControlBath2Positive44_5.Text)
            {
                labSheetA1Sheet.Bath2Positive44_5 = textBoxControlBath2Positive44_5.Text;
                AddLog("Bath 2 Positive 44.5", textBoxControlBath2Positive44_5.Text);
            }
        }
        private void textBoxControlBath3Positive44_5_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.Bath3Positive44_5 != textBoxControlBath3Positive44_5.Text)
            {
                labSheetA1Sheet.Bath3Positive44_5 = textBoxControlBath3Positive44_5.Text;
                AddLog("Bath 3 Positive 44.5", textBoxControlBath3Positive44_5.Text);
            }
        }
        private void textBoxIncubationBath1StartTime_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.IncubationBath1StartTime != textBoxIncubationBath1StartTime.Text)
            {
                labSheetA1Sheet.IncubationBath1StartTime = textBoxIncubationBath1StartTime.Text;
                AddLog("Incubation Bath 1 Start Time", textBoxIncubationBath1StartTime.Text);
            }
        }
        private void textBoxIncubationBath2StartTime_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.IncubationBath2StartTime != textBoxIncubationBath2StartTime.Text)
            {
                labSheetA1Sheet.IncubationBath2StartTime = textBoxIncubationBath2StartTime.Text;
                AddLog("Incubation Bath 2 Start Time", textBoxIncubationBath2StartTime.Text);
            }
        }
        private void textBoxIncubationBath3StartTime_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.IncubationBath3StartTime != textBoxIncubationBath3StartTime.Text)
            {
                labSheetA1Sheet.IncubationBath3StartTime = textBoxIncubationBath3StartTime.Text;
                AddLog("Incubation Bath 3 Start Time", textBoxIncubationBath3StartTime.Text);
            }
        }
        private void textBoxIncubationBath1EndTime_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.IncubationBath1EndTime != textBoxIncubationBath1EndTime.Text)
            {
                labSheetA1Sheet.IncubationBath1EndTime = textBoxIncubationBath1EndTime.Text;
                AddLog("Incubation Bath 1 End Time", textBoxIncubationBath1EndTime.Text);
            }
        }
        private void textBoxIncubationBath2EndTime_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.IncubationBath2EndTime != textBoxIncubationBath2EndTime.Text)
            {
                labSheetA1Sheet.IncubationBath2EndTime = textBoxIncubationBath2EndTime.Text;
                AddLog("Incubation Bath 2 End Time", textBoxIncubationBath2EndTime.Text);
            }
        }
        private void textBoxIncubationBath3EndTime_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.IncubationBath3EndTime != textBoxIncubationBath3EndTime.Text)
            {
                labSheetA1Sheet.IncubationBath3EndTime = textBoxIncubationBath3EndTime.Text;
                AddLog("Incubation Bath 3 End Time", textBoxIncubationBath3EndTime.Text);
            }
        }
        private void textBoxInitials_Leave(object sender, EventArgs e)
        {
            textBoxInitials.Text = textBoxInitials.Text.ToUpper();
            Initials = textBoxInitials.Text;
        }
        private void textBoxLot35_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.Lot35 != textBoxLot35.Text)
            {
                labSheetA1Sheet.Lot35 = textBoxLot35.Text;
                AddLog("Lot 35", textBoxLot35.Text);
            }

            textBoxLot35.ForeColor = Color.Black;
            textBoxLot44_5.ForeColor = Color.Black;

            if (textBoxLot35.Text.Trim().ToUpper() == textBoxLot44_5.Text.Trim().ToUpper())
            {
                textBoxLot35.ForeColor = Color.Red;
                textBoxLot44_5.ForeColor = Color.Red;
            }

        }
        private void textBoxLot44_5_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.Lot44_5 != textBoxLot44_5.Text)
            {
                labSheetA1Sheet.Lot44_5 = textBoxLot44_5.Text;
                AddLog("Lot 44.5", textBoxLot44_5.Text);
            }

            textBoxLot35.ForeColor = Color.Black;
            textBoxLot44_5.ForeColor = Color.Black;

            if (textBoxLot35.Text.Trim().ToUpper() == textBoxLot44_5.Text.Trim().ToUpper())
            {
                textBoxLot35.ForeColor = Color.Red;
                textBoxLot44_5.ForeColor = Color.Red;
            }
        }
        private void textBoxResultsReadBy_Leave(object sender, EventArgs e)
        {
            textBoxResultsReadBy.Text = textBoxResultsReadBy.Text.ToUpper();

            if (labSheetA1Sheet.ResultsReadBy != textBoxResultsReadBy.Text)
            {
                labSheetA1Sheet.ResultsReadBy = textBoxResultsReadBy.Text;
                AddLog("Results Read By", textBoxResultsReadBy.Text);
            }
        }
        private void textBoxResultsRecordedBy_Leave(object sender, EventArgs e)
        {
            textBoxResultsRecordedBy.Text = textBoxResultsRecordedBy.Text.ToUpper();

            if (labSheetA1Sheet.ResultsRecordedBy != textBoxResultsRecordedBy.Text)
            {
                labSheetA1Sheet.ResultsRecordedBy = textBoxResultsRecordedBy.Text;
                AddLog("Results Recorded By", textBoxResultsRecordedBy.Text);
            }
        }
        private void textBoxSalinitiesReadBy_Leave(object sender, EventArgs e)
        {
            textBoxSalinitiesReadBy.Text = textBoxSalinitiesReadBy.Text.ToUpper();

            if (labSheetA1Sheet.SalinitiesReadBy != textBoxSalinitiesReadBy.Text)
            {
                labSheetA1Sheet.SalinitiesReadBy = textBoxSalinitiesReadBy.Text;
                AddLog("Salinities Read By", textBoxSalinitiesReadBy.Text);
            }
        }
        private void textBoxSampleBottleLotNumber_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.SampleBottleLotNumber != textBoxSampleBottleLotNumber.Text)
            {
                labSheetA1Sheet.SampleBottleLotNumber = textBoxSampleBottleLotNumber.Text;
                AddLog("Sample Bottle Lot Number", textBoxSampleBottleLotNumber.Text);
            }
        }
        private void textBoxSampleCrewInitials_Leave(object sender, EventArgs e)
        {
            textBoxSampleCrewInitials.Text = textBoxSampleCrewInitials.Text.ToUpper();

            if (labSheetA1Sheet.SampleCrewInitials != textBoxSampleCrewInitials.Text)
            {
                labSheetA1Sheet.SampleCrewInitials = textBoxSampleCrewInitials.Text;
                AddLog("Sample Crew Initials", textBoxSampleCrewInitials.Text);
            }
        }
        private void textBoxTCField1_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.TCField1 != textBoxTCField1.Text)
            {
                labSheetA1Sheet.TCField1 = textBoxTCField1.Text;
                AddLog("TC Field #1", textBoxTCField1.Text);
            }
        }
        private void textBoxTCField2_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.TCField2 != textBoxTCField2.Text)
            {
                labSheetA1Sheet.TCField2 = textBoxTCField2.Text;
                AddLog("TC Field #2", textBoxTCField2.Text);
            }
        }
        private void textBoxTCLab1_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.TCLab1 != textBoxTCLab1.Text)
            {
                labSheetA1Sheet.TCLab1 = textBoxTCLab1.Text;
                AddLog("TC Lab #1", textBoxTCLab1.Text);
            }
        }
        private void textBoxTCLab2_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.TCLab2 != textBoxTCLab2.Text)
            {
                labSheetA1Sheet.TCLab2 = textBoxTCLab2.Text;
                AddLog("TC Lab #2", textBoxTCLab2.Text);
            }
        }
        private void textBoxTides_Leave(object sender, EventArgs e)
        {
            textBoxTides.ForeColor = Color.Black;
            textBoxTides.Text = textBoxTides.Text.ToUpper();
            if (string.IsNullOrWhiteSpace(textBoxTides.Text))
            {
                textBoxTides.Text = "-- / --";
                return;
            }
            if (textBoxTides.Text.Length < 7 || textBoxTides.Text.Length > 7)
            {
                textBoxTides.ForeColor = Color.Red;
                lblStatus.Text = "Tides text should contain exactly 7 characters. Ex: [HT / HT].";
                return;
            }
            List<string> strList = textBoxTides.Text.Split(" ".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();
            foreach (string s in strList)
            {
                if (!AllowableTideString.Contains(s))
                {
                    textBoxTides.ForeColor = Color.Red;
                    lblStatus.Text = "Text for tides has to be of the form [HT / HT]. With allowables [HR, HT, HF, MR, MT, MF, LR, LT, LF]";
                    return;
                }
            }

            if (labSheetA1Sheet.Tides != textBoxTides.Text)
            {
                labSheetA1Sheet.Tides = textBoxTides.Text;
                AddLog("Tides", textBoxTides.Text);
            }
        }
        private void textBoxWaterBath1Number_Leave(object sender, EventArgs e)
        {
            textBoxWaterBath1Number.Text = textBoxWaterBath1Number.Text.ToUpper();

            if (labSheetA1Sheet.WaterBath1 != textBoxWaterBath1Number.Text)
            {
                labSheetA1Sheet.WaterBath1 = textBoxWaterBath1Number.Text;
                AddLog("Water Bath 1", textBoxWaterBath1Number.Text);
            }
        }
        private void textBoxWaterBath2Number_Leave(object sender, EventArgs e)
        {
            textBoxWaterBath2Number.Text = textBoxWaterBath2Number.Text.ToUpper();

            if (labSheetA1Sheet.WaterBath2 != textBoxWaterBath2Number.Text)
            {
                labSheetA1Sheet.WaterBath2 = textBoxWaterBath2Number.Text;
                AddLog("Water Bath 2", textBoxWaterBath2Number.Text);
            }
        }
        private void textBoxWaterBath3Number_Leave(object sender, EventArgs e)
        {
            textBoxWaterBath3Number.Text = textBoxWaterBath3Number.Text.ToUpper();

            if (labSheetA1Sheet.WaterBath3 != textBoxWaterBath3Number.Text)
            {
                labSheetA1Sheet.WaterBath3 = textBoxWaterBath3Number.Text;
                AddLog("Water Bath 3", textBoxWaterBath3Number.Text);
            }
        }
        #endregion Events Focus
        #region Events KeyDown
        private void dataGridViewCSSP_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F1)
            {
                if (csspWQInputSheetType == CSSPWQInputSheetTypeEnum.A1)
                {
                    switch (dataGridViewCSSP.CurrentCell.ColumnIndex)
                    {
                        case 0:
                            {
                                lblStatus.Text = "Read only. Provided by sampling plan file. Pressing F2 will also set daily duplicate. F3 will also set intertech duplicate. F4 will also set intertech read. F5 add another sample time for the site.";
                            }
                            break;
                        case 1:
                            {
                                lblStatus.Text = "Read only. MWQM Site name";
                            }
                            break;
                        case 2:
                            {
                                lblStatus.Text = "Sampling Time. Time is entered with 4 digits. 1234 will be converted to 12:34";
                            }
                            break;
                        case 3:
                            {
                                lblStatus.Text = "Read only. MPN is calculated from the 3 positive tubes columns";
                            }
                            break;
                        case 4:
                            {
                                lblStatus.Text = "Allowable values are 0 to 5";
                            }
                            break;
                        case 5:
                            {
                                lblStatus.Text = "Allowable values are 0 to 5";
                            }
                            break;
                        case 6:
                            {
                                lblStatus.Text = "Allowable values are 0 to 5";
                            }
                            break;
                        case 7:
                            {
                                lblStatus.Text = "Salinity (PPT). Allowable number from 0 to 36";
                            }
                            break;
                        case 8:
                            {
                                lblStatus.Text = "Temperature (degree Celcius). Allowable number from 0 to 35";
                            }
                            break;
                        case 9:
                            {
                                lblStatus.Text = "Initial of person. Lowercase will automatically convert to uppercase";
                            }
                            break;
                        case 10:
                            {
                                lblStatus.Text = "Read only. Provided by the sampling plan file. Example of allowable values Normal, Duplicate, After Rain, Study, Infrastructure etc...";
                            }
                            break;
                        case 11:
                            {
                                lblStatus.Text = "Read only. TVItemID of the MWQM site.";
                            }
                            break;
                        case 12:
                            {
                                lblStatus.Text = "Comment associated to the MWQM site for the particular run.";
                            }
                            break;
                        default:
                            break;
                    }
                }
            }
            else if (e.KeyCode == Keys.F2)
            {
                int TextIndex = 0;
                int StationIndex = 1;
                int TVItemIDIndex = 11;
                int SampleTypeIndex = 10;
                if (csspWQInputSheetType == CSSPWQInputSheetTypeEnum.A1)
                {
                    switch (dataGridViewCSSP.CurrentCell.ColumnIndex)
                    {
                        case 0:
                            {
                                if (!(dataGridViewCSSP.CurrentCell.Value.ToString().Contains(SampleTypeEnum.DailyDuplicate.ToString())
                                    || dataGridViewCSSP.CurrentCell.Value.ToString().Contains(SampleTypeEnum.IntertechDuplicate.ToString())
                                    || dataGridViewCSSP.CurrentCell.Value.ToString().Contains(SampleTypeEnum.IntertechRead.ToString())))
                                {
                                    int RowOfDuplicate = 0;
                                    for (int i = 0, count = dataGridViewCSSP.Rows.Count; i < count; i++)
                                    {
                                        if (dataGridViewCSSP[0, i].Value.ToString().Contains(SampleTypeEnum.DailyDuplicate.ToString()))
                                        {
                                            RowOfDuplicate = i;
                                            break;
                                        }
                                    }
                                    if (RowOfDuplicate == 0)
                                    {
                                        dataGridViewCSSP.Rows.AddCopy(dataGridViewCSSP.CurrentCell.RowIndex);
                                        RowOfDuplicate = dataGridViewCSSP.Rows.Count - 1;
                                        for (int col = 0, count = dataGridViewCSSP.Rows[RowOfDuplicate].Cells.Count; col < count; col++)
                                        {
                                            if (col == SampleTypeIndex)
                                            {
                                                dataGridViewCSSP.Rows[RowOfDuplicate].Cells[col].Value = SampleTypeEnum.DailyDuplicate.ToString();
                                            }
                                            else
                                            {
                                                if (col == 0)
                                                {
                                                    dataGridViewCSSP.Rows[RowOfDuplicate].Cells[col].Value = ((string)dataGridViewCSSP.Rows[dataGridViewCSSP.CurrentCell.RowIndex].Cells[col].Value)
                                                        .Replace(SampleTypeEnum.Infrastructure.ToString(), SampleTypeEnum.DailyDuplicate.ToString())
                                                        .Replace(SampleTypeEnum.IntertechDuplicate.ToString(), SampleTypeEnum.DailyDuplicate.ToString())
                                                        .Replace(SampleTypeEnum.IntertechRead.ToString(), SampleTypeEnum.DailyDuplicate.ToString())
                                                        .Replace(SampleTypeEnum.RainCMPRoutine.ToString(), SampleTypeEnum.DailyDuplicate.ToString())
                                                        .Replace(SampleTypeEnum.RainRun.ToString(), SampleTypeEnum.DailyDuplicate.ToString())
                                                        .Replace(SampleTypeEnum.ReopeningEmergencyRain.ToString(), SampleTypeEnum.DailyDuplicate.ToString())
                                                        .Replace(SampleTypeEnum.ReopeningSpill.ToString(), SampleTypeEnum.DailyDuplicate.ToString())
                                                        .Replace(SampleTypeEnum.Routine.ToString(), SampleTypeEnum.DailyDuplicate.ToString())
                                                        .Replace(SampleTypeEnum.Sanitary.ToString(), SampleTypeEnum.DailyDuplicate.ToString())
                                                        .Replace(SampleTypeEnum.Study.ToString(), SampleTypeEnum.DailyDuplicate.ToString());


                                                    while (((string)dataGridViewCSSP.Rows[RowOfDuplicate].Cells[col].Value).Contains("  "))
                                                    {
                                                        dataGridViewCSSP.Rows[RowOfDuplicate].Cells[col].Value = ((string)dataGridViewCSSP.Rows[RowOfDuplicate].Cells[col].Value).Replace("  ", " ");
                                                    }

                                                    dataGridViewCSSP.Rows[RowOfDuplicate].Cells[col].Value = ((string)dataGridViewCSSP.Rows[RowOfDuplicate].Cells[col].Value).Replace(SampleTypeEnum.DailyDuplicate.ToString(), SampleTypeEnum.DailyDuplicate.ToString() + "             ");

                                                }
                                                else
                                                {
                                                    dataGridViewCSSP.Rows[RowOfDuplicate].Cells[col].Value = dataGridViewCSSP.Rows[dataGridViewCSSP.CurrentCell.RowIndex].Cells[col].Value;
                                                }
                                            }
                                        }
                                        dataGridViewCSSP.Rows[RowOfDuplicate].Cells[2].Value = "";
                                        dataGridViewCSSP.Rows[RowOfDuplicate].Cells[3].Value = "";
                                        dataGridViewCSSP.Rows[RowOfDuplicate].Cells[4].Value = "";
                                        dataGridViewCSSP.Rows[RowOfDuplicate].Cells[5].Value = "";
                                        dataGridViewCSSP.Rows[RowOfDuplicate].Cells[6].Value = "";
                                        dataGridViewCSSP.Rows[RowOfDuplicate].Cells[7].Value = "";
                                        dataGridViewCSSP.Rows[RowOfDuplicate].Cells[8].Value = "";
                                        dataGridViewCSSP.Rows[RowOfDuplicate].Cells[9].Value = "";
                                        dataGridViewCSSP.Rows[RowOfDuplicate].Cells[12].Value = "";
                                        DoSave();
                                    }
                                    else
                                    {
                                        string MWQMSiteOld = dataGridViewCSSP[StationIndex, RowOfDuplicate].Value.ToString();
                                        string MWQMSite = dataGridViewCSSP[StationIndex, dataGridViewCSSP.CurrentCell.RowIndex].Value.ToString().Trim();
                                        if (DialogResult.OK == MessageBox.Show("Create Daily Duplicate with " + MWQMSite, "Setting Daily Duplicate Site", MessageBoxButtons.OKCancel))
                                        {
                                            dataGridViewCSSP[StationIndex, RowOfDuplicate].Value = MWQMSite;
                                            dataGridViewCSSP[TextIndex, RowOfDuplicate].Value = dataGridViewCSSP[0, RowOfDuplicate].Value.ToString().Replace(MWQMSiteOld, MWQMSite);
                                            dataGridViewCSSP[TVItemIDIndex, RowOfDuplicate].Value = dataGridViewCSSP[TVItemIDIndex, dataGridViewCSSP.CurrentCell.RowIndex].Value.ToString().Trim();
                                            AddLog("Change Daily Duplicate [" + MWQMSiteOld + "]", MWQMSite);
                                            DoSave();
                                        }
                                    }
                                }
                            }
                            break;
                        default:
                            break;
                    }
                }
            }
            else if (e.KeyCode == Keys.F3)
            {
                int StationIndex = 1;
                int TVItemIDIndex = 11;
                if (csspWQInputSheetType == CSSPWQInputSheetTypeEnum.A1)
                {
                    switch (dataGridViewCSSP.CurrentCell.ColumnIndex)
                    {
                        case 0:
                            {
                                if (!(dataGridViewCSSP.CurrentCell.Value.ToString().Contains(SampleTypeEnum.DailyDuplicate.ToString())
                                    || dataGridViewCSSP.CurrentCell.Value.ToString().Contains(SampleTypeEnum.IntertechDuplicate.ToString())
                                    || dataGridViewCSSP.CurrentCell.Value.ToString().Contains(SampleTypeEnum.IntertechRead.ToString())))
                                {
                                    int RowOfIntertechDuplicate = 0;
                                    for (int i = 0, count = dataGridViewCSSP.Rows.Count; i < count; i++)
                                    {
                                        if (dataGridViewCSSP[0, i].Value.ToString().Contains(SampleTypeEnum.IntertechDuplicate.ToString()))
                                        {
                                            RowOfIntertechDuplicate = i;
                                            break;
                                        }
                                    }
                                    if (RowOfIntertechDuplicate == 0)
                                    {
                                        string MWQMSite = dataGridViewCSSP[StationIndex, dataGridViewCSSP.CurrentCell.RowIndex].Value.ToString().Trim();
                                        if (DialogResult.OK == MessageBox.Show("Create Intertech Duplicate with " + MWQMSite, "Setting Intertech Duplicate Site", MessageBoxButtons.OKCancel))
                                        {
                                            object[] row = { MWQMSite + " - " + SampleTypeEnum.IntertechDuplicate.ToString() + "    " +
                                            SpaceStr.Substring(0, SpaceStr.Length - 0) + "",
                                            MWQMSite, "", "", "", "", "", "", "", "", SampleTypeEnum.IntertechDuplicate.ToString(),
                                            dataGridViewCSSP[TVItemIDIndex, dataGridViewCSSP.CurrentCell.RowIndex].Value.ToString(), ""};
                                            dataGridViewCSSP.Rows.Add(row);
                                            DoSave();
                                            ReadFileFromLocalMachine();
                                        }
                                    }
                                    else
                                    {
                                        if (DialogResult.OK == MessageBox.Show("Remove Intertech Duplicate", "Setting Intertech Duplicate Site", MessageBoxButtons.OKCancel))
                                        {
                                            dataGridViewCSSP.Rows.RemoveAt(RowOfIntertechDuplicate);
                                            SaveInfoOnLocalMachine(false);
                                            DoSave();
                                            ReadFileFromLocalMachine();
                                        }
                                    }
                                }
                            }
                            break;
                        default:
                            break;
                    }
                }
            }
            else if (e.KeyCode == Keys.F4)
            {
                int StationIndex = 1;
                int TVItemIDIndex = 11;
                if (csspWQInputSheetType == CSSPWQInputSheetTypeEnum.A1)
                {
                    switch (dataGridViewCSSP.CurrentCell.ColumnIndex)
                    {
                        case 0:
                            {
                                if (!(dataGridViewCSSP.CurrentCell.Value.ToString().Contains(SampleTypeEnum.DailyDuplicate.ToString())
                                    || dataGridViewCSSP.CurrentCell.Value.ToString().Contains(SampleTypeEnum.IntertechDuplicate.ToString())
                                    || dataGridViewCSSP.CurrentCell.Value.ToString().Contains(SampleTypeEnum.IntertechRead.ToString())))
                                {
                                    int RowOfIntertechRead = 0;
                                    for (int i = 0, count = dataGridViewCSSP.Rows.Count; i < count; i++)
                                    {
                                        if (dataGridViewCSSP[0, i].Value.ToString().Contains(SampleTypeEnum.IntertechRead.ToString()))
                                        {
                                            RowOfIntertechRead = i;
                                            break;
                                        }
                                    }
                                    if (RowOfIntertechRead == 0)
                                    {
                                        string MWQMSite = dataGridViewCSSP[StationIndex, dataGridViewCSSP.CurrentCell.RowIndex].Value.ToString().Trim();
                                        if (DialogResult.OK == MessageBox.Show("Create Intertech Read with " + MWQMSite, "Setting Intertech Read Site", MessageBoxButtons.OKCancel))
                                        {
                                            object[] row = { MWQMSite + " - " + SampleTypeEnum.IntertechRead.ToString() + "    " +
                                            SpaceStr.Substring(0, SpaceStr.Length - 0) + "",
                                            MWQMSite, "", "", "", "", "", "", "", "", SampleTypeEnum.IntertechRead.ToString(),
                                            dataGridViewCSSP[TVItemIDIndex, dataGridViewCSSP.CurrentCell.RowIndex].Value.ToString(), ""};
                                            dataGridViewCSSP.Rows.Add(row);
                                            DoSave();
                                            ReadFileFromLocalMachine();
                                        }
                                    }
                                    else
                                    {
                                        if (DialogResult.OK == MessageBox.Show("Remove Intertech Read", "Setting Intertech Read Site", MessageBoxButtons.OKCancel))
                                        {
                                            dataGridViewCSSP.Rows.RemoveAt(RowOfIntertechRead);
                                            DoSave();
                                            ReadFileFromLocalMachine();
                                        }
                                    }
                                }
                            }
                            break;
                        default:
                            break;
                    }
                }
            }
            else if (e.KeyCode == Keys.F5)
            {
                int StationIndex = 1;
                int TVItemIDIndex = 11;
                int SampleType = 9;
                if (csspWQInputSheetType == CSSPWQInputSheetTypeEnum.A1)
                {
                    switch (dataGridViewCSSP.CurrentCell.ColumnIndex)
                    {
                        case 0:
                            {
                                if (!(dataGridViewCSSP.CurrentCell.Value.ToString().Contains(SampleTypeEnum.DailyDuplicate.ToString())
                                    || dataGridViewCSSP.CurrentCell.Value.ToString().Contains(SampleTypeEnum.IntertechDuplicate.ToString())
                                    || dataGridViewCSSP.CurrentCell.Value.ToString().Contains(SampleTypeEnum.IntertechRead.ToString())))
                                {
                                    int RowOfAnotherSample = 0;
                                    string MWQMSite = dataGridViewCSSP[StationIndex, dataGridViewCSSP.CurrentCell.RowIndex].Value.ToString().Trim();
                                    for (int i = dataGridViewCSSP.Rows.Count - 1; i > dataGridViewCSSP.CurrentCell.RowIndex; i--)
                                    {
                                        if (MWQMSite == dataGridViewCSSP[StationIndex, i].Value.ToString().Trim() && dataGridViewCSSP[SampleType, i].Value.ToString().Trim() == labSheetA1Sheet.SampleType.ToString())
                                        {
                                            RowOfAnotherSample = i;
                                            break;
                                        }
                                    }
                                    if (RowOfAnotherSample == 0)
                                    {
                                        DialogResult dialogResult = MessageBox.Show("Add another sampling time for site " + MWQMSite, "Same day sampling setup", MessageBoxButtons.YesNo);
                                        if (DialogResult.Yes == dialogResult)
                                        {
                                            object[] row = { MWQMSite + " - " + labSheetA1Sheet.SampleType.ToString() + "    " +
                                            SpaceStr.Substring(0, SpaceStr.Length - 0) + "",
                                            MWQMSite, "", "", "", "", "", "", "", "", labSheetA1Sheet.SampleType.ToString(),
                                            dataGridViewCSSP[TVItemIDIndex, dataGridViewCSSP.CurrentCell.RowIndex].Value.ToString(), ""};
                                            dataGridViewCSSP.Rows.Add(row);
                                            DoSave();
                                            ReadFileFromLocalMachine();
                                        }
                                    }
                                    else
                                    {
                                        DialogResult dialogResult = MessageBox.Show("Add (Yes) another sampling time for site " + MWQMSite + ". \r\n" + " or \r\n" +
                                            "Remove last (No) sampling time for site " + MWQMSite, "Same day sampling setup", MessageBoxButtons.YesNoCancel);
                                        if (DialogResult.Yes == dialogResult)
                                        {
                                            object[] row = { MWQMSite + " - " + labSheetA1Sheet.SampleType.ToString() + "    " +
                                            SpaceStr.Substring(0, SpaceStr.Length - 0) + "",
                                            MWQMSite, "", "", "", "", "", "", "", "", labSheetA1Sheet.SampleType.ToString(),
                                            dataGridViewCSSP[TVItemIDIndex, dataGridViewCSSP.CurrentCell.RowIndex].Value.ToString(), ""};
                                            dataGridViewCSSP.Rows.Add(row);
                                            DoSave();
                                            ReadFileFromLocalMachine();
                                        }
                                        else if (DialogResult.No == dialogResult)
                                        {
                                            dataGridViewCSSP.Rows.RemoveAt(RowOfAnotherSample);
                                            DoSave();
                                            ReadFileFromLocalMachine();
                                        }
                                    }
                                }
                            }
                            break;
                        default:
                            break;
                    }
                }
            }
        }
        private void richTextBoxRunWeatherComment_KeyDown(object sender, KeyEventArgs e)
        {
            richTextBoxRunWeatherComment.BackColor = TextBoxBackColor;

            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "Anything related to the RunWeatherComment during the sampling";
            }
        }
        private void richTextBoxRunComment_KeyDown(object sender, KeyEventArgs e)
        {
            richTextBoxRunComment.BackColor = TextBoxBackColor;

            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "Anything observed during the field trip";
            }
        }
        private void textBoxControlBlank35_KeyDown(object sender, KeyEventArgs e)
        {
            textBoxControlBlank35.BackColor = TextBoxBackColor;

            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    textBoxControlNegative35.Focus();
                }
                else
                {
                    textBoxControlBath1Positive44_5.Focus();
                }
            }
            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "+, - or N";
            }
        }
        private void textBoxControlBath1Blank44_5_KeyDown(object sender, KeyEventArgs e)
        {
            textBoxControlBath1Blank44_5.BackColor = TextBoxBackColor;

            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    textBoxControlBath1Negative44_5.Focus();
                }
                else
                {
                    if (radioButton1Baths.Checked)
                    {
                        textBoxLot35.Focus();
                    }
                    else
                    {
                        textBoxControlBath2Positive44_5.Focus();
                    }
                }
            }
            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "+, - or N";
            }
        }
        private void textBoxControlBath2Blank44_5_KeyDown(object sender, KeyEventArgs e)
        {
            textBoxControlBath2Blank44_5.BackColor = TextBoxBackColor;

            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    textBoxControlBath2Negative44_5.Focus();
                }
                else
                {
                    if (radioButton3Baths.Checked)
                    {
                        textBoxControlBath3Positive44_5.Focus();
                    }
                    else
                    {
                        textBoxLot35.Focus();
                    }
                }
            }
            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "+, - or N";
            }
        }
        private void textBoxControlBath3Blank44_5_KeyDown(object sender, KeyEventArgs e)
        {
            textBoxControlBath3Blank44_5.BackColor = TextBoxBackColor;

            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    textBoxControlBath3Negative44_5.Focus();
                }
                else
                {
                    textBoxLot35.Focus();
                }
            }
            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "+, - or N";
            }
        }
        private void textBoxControlLot_KeyDown(object sender, KeyEventArgs e)
        {
            textBoxControlLot.BackColor = TextBoxBackColor;

            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    if (radioButton3Baths.Checked)
                    {
                        textBoxWaterBath3Number.Focus();
                    }
                    else if (radioButton2Baths.Checked)
                    {
                        textBoxWaterBath2Number.Focus();
                    }
                    else
                    {
                        textBoxWaterBath1Number.Focus();
                    }
                }
                else
                {
                    textBoxControlPositive35.Focus();
                }
            }
            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "Anything";
            }
        }
        private void textBoxControlNegative35_KeyDown(object sender, KeyEventArgs e)
        {
            textBoxControlNegative35.BackColor = TextBoxBackColor;

            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    textBoxControlNonTarget35.Focus();
                }
                else
                {
                    textBoxControlBlank35.Focus();
                }
            }

            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "Only allowable characters are '+' and '-'";
            }
        }
        private void textBoxControlBath1Negative44_5_KeyDown(object sender, KeyEventArgs e)
        {
            textBoxControlBath1Negative44_5.BackColor = TextBoxBackColor;

            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    textBoxControlBath1NonTarget44_5.Focus();
                }
                else
                {
                    textBoxControlBath1Blank44_5.Focus();
                }
            }

            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "Only allowable characters are '+' and '-'";
            }
        }
        private void textBoxControlBath2Negative44_5_KeyDown(object sender, KeyEventArgs e)
        {
            textBoxControlBath2Negative44_5.BackColor = TextBoxBackColor;

            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    textBoxControlBath2NonTarget44_5.Focus();
                }
                else
                {
                    textBoxControlBath2Blank44_5.Focus();
                }
            }

            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "Only allowable characters are '+' and '-'";
            }
        }
        private void textBoxControlBath3Negative44_5_KeyDown(object sender, KeyEventArgs e)
        {
            textBoxControlBath3Negative44_5.BackColor = TextBoxBackColor;

            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    textBoxControlBath3NonTarget44_5.Focus();
                }
                else
                {
                    textBoxControlBath3Blank44_5.Focus();
                }
            }

            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "Only allowable characters are '+' and '-'";
            }
        }
        private void textBoxControlNonTarget35_KeyDown(object sender, KeyEventArgs e)
        {
            textBoxControlNonTarget35.BackColor = TextBoxBackColor;

            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    textBoxControlPositive35.Focus();
                }
                else
                {
                    textBoxControlNegative35.Focus();
                }
            }

            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "Only allowable characters are '+' , '-' or 'N' for not applicable";
            }
        }
        private void textBoxControlBath1NonTarget44_5_KeyDown(object sender, KeyEventArgs e)
        {
            textBoxControlBath1NonTarget44_5.BackColor = TextBoxBackColor;

            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    textBoxControlBath1Positive44_5.Focus();
                }
                else
                {
                    textBoxControlBath1Negative44_5.Focus();
                }
            }

            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "Only allowable characters are '+' , '-' or 'N' for not applicable";
            }
        }
        private void textBoxControlBath2NonTarget44_5_KeyDown(object sender, KeyEventArgs e)
        {
            textBoxControlBath2NonTarget44_5.BackColor = TextBoxBackColor;

            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    textBoxControlBath2Positive44_5.Focus();
                }
                else
                {
                    textBoxControlBath2Negative44_5.Focus();
                }
            }

            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "Only allowable characters are '+' , '-' or 'N' for not applicable";
            }
        }
        private void textBoxControlBath3NonTarget44_5_KeyDown(object sender, KeyEventArgs e)
        {
            textBoxControlBath3NonTarget44_5.BackColor = TextBoxBackColor;

            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    textBoxControlBath3Positive44_5.Focus();
                }
                else
                {
                    textBoxControlBath3Negative44_5.Focus();
                }
            }

            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "Only allowable characters are '+' , '-' or 'N' for not applicable";
            }
        }
        private void textBoxControlPositive35_KeyDown(object sender, KeyEventArgs e)
        {
            textBoxControlPositive35.BackColor = TextBoxBackColor;

            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    textBoxControlLot.Focus();
                }
                else
                {
                    textBoxControlNonTarget35.Focus();
                }
            }

            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "Only allowable characters are '+' and '-'";
            }
        }
        private void textBoxControlBath1Positive44_5_KeyDown(object sender, KeyEventArgs e)
        {
            textBoxControlBath1Positive44_5.BackColor = TextBoxBackColor;

            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    textBoxControlBlank35.Focus();
                }
                else
                {
                    textBoxControlBath1NonTarget44_5.Focus();
                }
            }

            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "Only allowable characters are '+' and '-'";
            }
        }
        private void textBoxControlBath2Positive44_5_KeyDown(object sender, KeyEventArgs e)
        {
            textBoxControlBath2Positive44_5.BackColor = TextBoxBackColor;

            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    textBoxControlBath1Blank44_5.Focus();
                }
                else
                {
                    textBoxControlBath2NonTarget44_5.Focus();
                }
            }

            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "Only allowable characters are '+' and '-'";
            }
        }
        private void textBoxControlBath3Positive44_5_KeyDown(object sender, KeyEventArgs e)
        {
            textBoxControlBath3Positive44_5.BackColor = TextBoxBackColor;

            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    textBoxControlBath2Blank44_5.Focus();
                }
                else
                {
                    textBoxControlBath3NonTarget44_5.Focus();
                }
            }

            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "Only allowable characters are '+' and '-'";
            }
        }
        private void textBoxDuplicatePrecisionCriteria_KeyDown(object sender, KeyEventArgs e)
        {
            textBoxDailyDuplicatePrecisionCriteria.BackColor = TextBoxBackColor;

            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    textBoxLot44_5.Focus();
                }
                else
                {
                    textBoxSampleBottleLotNumber.Focus();
                }
            }
            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "Lab specific precision criteria";
            }
        }
        private void textBoxIncubationBath1StartTime_KeyDown(object sender, KeyEventArgs e)
        {
            textBoxIncubationBath1StartTime.BackColor = TextBoxBackColor;

            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    if (radioButton3Baths.Checked)
                    {
                        radioButton3Baths.Focus();
                    }
                    else if (radioButton2Baths.Checked)
                    {
                        radioButton2Baths.Focus();
                    }
                    else
                    {
                        radioButton1Baths.Focus();
                    }
                }
                else
                {
                    textBoxIncubationBath1EndTime.Focus();
                }
            }
            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "All time should be entered with 4 digits. 1234 for 12:34. ':' will be added automatically";
            }
        }
        private void textBoxIncubationBath2StartTime_KeyDown(object sender, KeyEventArgs e)
        {
            textBoxIncubationBath2StartTime.BackColor = TextBoxBackColor;

            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    textBoxWaterBath1Number.Focus();
                }
                else
                {
                    textBoxIncubationBath2EndTime.Focus();
                }
            }
            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "All time should be entered with 4 digits. 1234 for 12:34. ':' will be added automatically";
            }
        }
        private void textBoxIncubationBath3StartTime_KeyDown(object sender, KeyEventArgs e)
        {
            textBoxIncubationBath3StartTime.BackColor = TextBoxBackColor;

            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    textBoxWaterBath2Number.Focus();
                }
                else
                {
                    textBoxIncubationBath3EndTime.Focus();
                }
            }
            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "All time should be entered with 4 digits. 1234 for 12:34. ':' will be added automatically";
            }
        }
        private void textBoxIncubationBath1EndTime_KeyDown(object sender, KeyEventArgs e)
        {
            textBoxIncubationBath1EndTime.BackColor = TextBoxBackColor;

            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    textBoxIncubationBath1StartTime.Focus();
                }
                else
                {
                    textBoxWaterBath1Number.Focus();
                }
            }
            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "All time should be entered with 4 digits. 1234 for 12:34. ':' will be added automatically";
            }
        }
        private void textBoxIncubationBath2EndTime_KeyDown(object sender, KeyEventArgs e)
        {
            textBoxIncubationBath2EndTime.BackColor = TextBoxBackColor;

            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    textBoxIncubationBath2StartTime.Focus();
                }
                else
                {
                    textBoxWaterBath2Number.Focus();
                }
            }
            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "All time should be entered with 4 digits. 1234 for 12:34. ':' will be added automatically";
            }
        }
        private void textBoxIncubationBath3EndTime_KeyDown(object sender, KeyEventArgs e)
        {
            textBoxIncubationBath3EndTime.BackColor = TextBoxBackColor;

            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    textBoxIncubationBath3StartTime.Focus();
                }
                else
                {
                    textBoxWaterBath3Number.Focus();
                }
            }
            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "All time should be entered with 4 digits. 1234 for 12:34. ':' will be added automatically";
            }
        }
        private void textBoxLot35_KeyDown(object sender, KeyEventArgs e)
        {
            textBoxLot35.BackColor = TextBoxBackColor;
            textBoxLot44_5.BackColor = TextBoxBackColor;

            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    if (radioButton3Baths.Checked)
                    {
                        textBoxControlBath3Blank44_5.Focus();
                    }
                    else if (radioButton2Baths.Checked)
                    {
                        textBoxControlBath2Blank44_5.Focus();
                    }
                    else
                    {
                        textBoxControlBath1Blank44_5.Focus();
                    }
                }
                else
                {
                    textBoxLot44_5.Focus();
                }
            }
            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "Anything";
            }
        }
        private void textBoxLot44_5_KeyDown(object sender, KeyEventArgs e)
        {
            textBoxLot35.BackColor = TextBoxBackColor;
            textBoxLot44_5.BackColor = TextBoxBackColor;

            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    textBoxLot35.Focus();
                }
                else
                {
                    richTextBoxRunWeatherComment.Focus();
                }
            }
            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "Anything";
            }
        }
        private void textBoxResultsReadBy_KeyDown(object sender, KeyEventArgs e)
        {
            textBoxResultsReadBy.BackColor = TextBoxBackColor;

            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    textBoxSalinitiesReadBy.Focus();
                }
                else
                {
                    textBoxResultsRecordedBy.Focus();
                }
            }
            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "Initials of person who read the results";
            }
        }
        private void textBoxResultsRecordedBy_KeyDown(object sender, KeyEventArgs e)
        {
            textBoxResultsRecordedBy.BackColor = TextBoxBackColor;

            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    textBoxResultsRecordedBy.Focus();
                }
                else
                {
                    dataGridViewCSSP.Focus();
                }
            }
            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "Initials of person who recorded the results";
            }
        }
        private void textBoxSampleBottleLotNumber_KeyDown(object sender, KeyEventArgs e)
        {
            textBoxSampleBottleLotNumber.BackColor = TextBoxBackColor;

            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    richTextBoxRunComment.Focus();
                }
                else
                {
                    textBoxSalinitiesReadBy.Focus();
                }
            }
            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "Anything representing sample bottle lot number";
            }
        }
        private void textBoxSampleCrewInitials_KeyDown(object sender, KeyEventArgs e)
        {
            textBoxSampleCrewInitials.BackColor = TextBoxBackColor;

            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    textBoxTides.Focus();
                }
                else
                {
                    textBoxTCField1.Focus();
                }
            }
            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "Initials of Sampling Crew. Lowercase is ok. It will be set to uppercase automatically";
            }
        }
        private void textBoxSalinitiesReadBy_KeyDown(object sender, KeyEventArgs e)
        {
            textBoxSalinitiesReadBy.BackColor = TextBoxBackColor;

            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    textBoxSampleBottleLotNumber.Focus();
                }
                else
                {
                    textBoxResultsRecordedBy.Focus();
                }
            }
            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "Initials of person who measured the salinities in the lab";
            }
        }
        private void textBoxTCField1_KeyDown(object sender, KeyEventArgs e)
        {
            textBoxTCField1.BackColor = TextBoxBackColor;

            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    checkBox2Coolers.Focus();
                }
                else
                {
                    if (checkBox2Coolers.Checked)
                    {
                        textBoxTCField2.Focus();
                    }
                    else
                    {
                        textBoxTCLab1.Focus();
                    }
                }
            }
            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "Enter temperature. Only digits and '.' are accepted";
            }
        }
        private void textBoxTCField2_KeyDown(object sender, KeyEventArgs e)
        {
            textBoxTCField2.BackColor = TextBoxBackColor;

            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    textBoxTCField1.Focus();
                }
                else
                {
                    textBoxTCLab1.Focus();
                }
            }
            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "Enter temperature. Only digits and '.' are accepted";
            }
        }
        private void textBoxTCLab1_KeyDown(object sender, KeyEventArgs e)
        {
            textBoxTCLab1.BackColor = TextBoxBackColor;

            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    textBoxTCField1.Focus();
                }
                else
                {
                    if (checkBox2Coolers.Checked)
                    {
                        textBoxTCField2.Focus();
                    }
                    else
                    {
                        checkBoxIncubationStartSameDay.Focus();
                    }
                }
            }
            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "Enter temperature. Only digits and '.' are accepted";
            }
        }
        private void textBoxTCLab2_KeyDown(object sender, KeyEventArgs e)
        {
            textBoxTCLab2.BackColor = TextBoxBackColor;

            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    textBoxTCLab1.Focus();
                }
                else
                {
                    checkBoxIncubationStartSameDay.Focus();
                }
            }
            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "Enter temperature. Only digits and '.' are accepted";
            }
        }
        private void textBoxTide_KeyDown(object sender, KeyEventArgs e)
        {
            textBoxTides.BackColor = TextBoxBackColor;

            if (e.KeyCode == Keys.Enter)
            {
                textBoxSampleCrewInitials.Focus();
            }
            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "Allowables are [HR, HT, HF, MR, MT, MF, LR, LT, LF]";
            }
        }
        private void textBoxWaterBath1Number_KeyDown(object sender, KeyEventArgs e)
        {
            textBoxWaterBath1Number.BackColor = TextBoxBackColor;

            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    textBoxIncubationBath1EndTime.Focus();
                }
                else
                {
                    if (!radioButton1Baths.Checked)
                    {
                        textBoxIncubationBath2StartTime.Focus();
                    }
                    else
                    {
                        textBoxControlLot.Focus();
                    }
                }
            }
            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "Anything. Will be added automatically converted to uppercase";
            }
        }
        private void textBoxWaterBath2Number_KeyDown(object sender, KeyEventArgs e)
        {
            textBoxWaterBath2Number.BackColor = TextBoxBackColor;

            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    textBoxIncubationBath2EndTime.Focus();
                }
                else
                {
                    if (radioButton3Baths.Checked)
                    {
                        textBoxIncubationBath3StartTime.Focus();
                    }
                    else
                    {
                        textBoxControlLot.Focus();
                    }
                }
            }
            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "Anything. Will be added automatically converted to uppercase";
            }
        }
        private void textBoxWaterBath3Number_KeyDown(object sender, KeyEventArgs e)
        {
            textBoxWaterBath3Number.BackColor = TextBoxBackColor;

            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    textBoxIncubationBath3EndTime.Focus();
                }
                else
                {
                    textBoxControlLot.Focus();
                }
            }
            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "Anything. Will be added automatically converted to uppercase";
            }
        }
        #endregion Events KeyDown
        #region Events KeyPress
        #endregion Events KeyPress
        #region Events KeyUp
        private void listBoxFiles_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && butOpen.Enabled)
            {
                OpenFileName();
            }
        }
        #endregion Events KeyUp
        #region Events listBoxFiles
        private void listBoxFiles_DoubleClick(object sender, EventArgs e)
        {
            OpenFileName();
        }
        private void listBoxFiles_SelectedIndexChanged(object sender, EventArgs e)
        {
            lblFilePath.Text = "";
            butOpen.Enabled = false;
            butViewFCForm.Enabled = false;

            string FileName = ((FileItemList)listBoxFiles.SelectedItem).FileName;
            if (string.IsNullOrWhiteSpace(FileName))
            {
                richTextBoxFile.Text = "Select a date";
                return;
            }
            else
            {
                butCreateFile.Visible = false;
            }

            butOpen.Enabled = true;
            lblFilePath.Text = "";
            lblFilePath.Text = FileName;

            FileInfo fi = new FileInfo(FileName);

            if (fi.Exists)
            {
                StreamReader sr = fi.OpenText();
                string FileContent = sr.ReadToEnd();
                sr.Close();

                richTextBoxFile.Text = (FileContent.Length > 0 ? FileContent : "File empty");
            }

            string Rest = FileName.Replace(CurrentPath, "");
            int Pos = Rest.IndexOf("_");
            string Subsector = Rest.Substring(0, Pos);
            string Year = Rest.Substring(Pos + 1, 4);
            string Month = Rest.Substring(Pos + 6, 2);
            string Day = Rest.Substring(Pos + 9, 2);

            //NoUpdate = true;

            //foreach (FileItem item in comboBoxSubsectorNames.Items)
            //{
            //    if (item.Name.Contains(Subsector))
            //    {
            //        comboBoxSubsectorNames.SelectedItem = item;
            //        break;
            //    }
            //}

            //DateTime dateTimeRun = new DateTime(int.Parse(Year), int.Parse(Month), int.Parse(Day));

            //dateTimePickerRun.Value = dateTimeRun;
            //NoUpdate = false;

        }
        #endregion Events listBoxFiles
        #region Events radioButtons
        private void radioButton1Baths_CheckedChanged(object sender, EventArgs e)
        {
            RadioButtonBathNumberChanged();
        }
        private void radioButton2Baths_CheckedChanged(object sender, EventArgs e)
        {
            RadioButtonBathNumberChanged();
        }
        private void radioButton3Baths_CheckedChanged(object sender, EventArgs e)
        {
            RadioButtonBathNumberChanged();
        }
        #endregion Events radioButtons
        #region Events TextChanged
        private void lblFilePath_TextChanged(object sender, EventArgs e)
        {
            if (lblFilePath.Text.Length == 0)
            {
                if (csspWQInputApp.IncludeLaboratoryQAQC)
                {
                    butViewFCForm.Visible = false;
                    butFail.Enabled = false;
                    butSendToServer.Enabled = false;
                }
            }
            else
            {
                if (csspWQInputApp.IncludeLaboratoryQAQC)
                {
                    butViewFCForm.Visible = true;
                    butFail.Enabled = true;
                    if (lblFilePath.Text.Substring(lblFilePath.Text.Length - 6) == "_F.txt")
                    {
                        butFail.BackColor = Color.LightSalmon;
                    }
                    else
                    {
                        butFail.BackColor = Color.LightGray;
                    }

                    if (!(lblFilePath.Text.Substring(lblFilePath.Text.Length - 6) == "_F.txt" || lblFilePath.Text.Substring(lblFilePath.Text.Length - 6) == "_C.txt"))
                    {
                        butFail.Enabled = false;
                    }
                    else
                    {
                        butFail.Enabled = true;
                    }
                }
                if (lblFilePath.Text.Substring(lblFilePath.Text.Length - 6) == "_C.txt")
                {
                    butSendToServer.Enabled = true;
                }
                else
                {
                    butSendToServer.Enabled = false;
                }
            }
        }
        private void richTextBoxRunWeatherComment_TextChanged(object sender, EventArgs e)
        {
            if (!InLoadingFile)
                Modifying();
        }
        private void richTextBoxRunComment_TextChanged(object sender, EventArgs e)
        {
            if (!InLoadingFile)
                Modifying();
        }
        private void textBoxControlBlank35_TextChanged(object sender, EventArgs e)
        {
            textBoxControlBlank35.ForeColor = Color.Black;
            if (textBoxControlBlank35.Text == "-" || textBoxControlBlank35.Text == "+")
            {
                if (!InLoadingFile)
                    textBoxControlBath1Positive44_5.Focus();

                textBoxControlBlank35.ForeColor = Color.Black;
                if (textBoxControlBlank35.Text == "+")
                {
                    textBoxControlBlank35.ForeColor = Color.Red;
                }
            }
            else
            {
                textBoxControlBlank35.Text = "";
                lblStatus.Text = "Only allowable characters are '+' and '-'";
            }

            if (!InLoadingFile)
                Modifying();
        }
        private void textBoxControlBath1Blank44_5_TextChanged(object sender, EventArgs e)
        {
            textBoxControlBath1Blank44_5.ForeColor = Color.Black;
            if (textBoxControlBath1Blank44_5.Text == "-" || textBoxControlBath1Blank44_5.Text == "+")
            {
                if (radioButton1Baths.Checked)
                {
                    if (!InLoadingFile)
                        textBoxLot35.Focus();
                }
                else
                {
                    if (!InLoadingFile)
                        textBoxControlBath2Positive44_5.Focus();
                }
                textBoxControlBath1Blank44_5.ForeColor = Color.Black;
                if (textBoxControlBath1Blank44_5.Text == "+")
                {
                    textBoxControlBath1Blank44_5.ForeColor = Color.Red;
                }
            }
            else
            {
                textBoxControlBath1Blank44_5.Text = "";
                lblStatus.Text = "Only allowable characters are '+' and '-'";
            }

            if (!InLoadingFile)
                Modifying();
        }
        private void textBoxControlBath2Blank44_5_TextChanged(object sender, EventArgs e)
        {
            textBoxControlBath2Blank44_5.ForeColor = Color.Black;
            if (textBoxControlBath2Blank44_5.Text == "-" || textBoxControlBath2Blank44_5.Text == "+")
            {
                if (radioButton3Baths.Checked)
                {
                    if (!InLoadingFile)
                        textBoxControlBath3Positive44_5.Focus();
                }
                else
                {
                    if (!InLoadingFile)
                        textBoxLot35.Focus();
                }
                textBoxControlBath2Blank44_5.ForeColor = Color.Black;
                if (textBoxControlBath2Blank44_5.Text == "+")
                {
                    textBoxControlBath2Blank44_5.ForeColor = Color.Red;
                }
            }
            else
            {
                textBoxControlBath2Blank44_5.Text = "";
                lblStatus.Text = "Only allowable characters are '+' and '-'";
            }

            if (!InLoadingFile)
                Modifying();
        }
        private void textBoxControlBath3Blank44_5_TextChanged(object sender, EventArgs e)
        {
            textBoxControlBath3Blank44_5.ForeColor = Color.Black;
            if (textBoxControlBath3Blank44_5.Text == "-" || textBoxControlBath3Blank44_5.Text == "+")
            {
                if (!InLoadingFile)
                    textBoxLot35.Focus();

                textBoxControlBath3Blank44_5.ForeColor = Color.Black;
                if (textBoxControlBath3Blank44_5.Text == "+")
                {
                    textBoxControlBath3Blank44_5.ForeColor = Color.Red;
                }
            }
            else
            {
                textBoxControlBath3Blank44_5.Text = "";
                lblStatus.Text = "Only allowable characters are '+' and '-'";
            }

            if (!InLoadingFile)
                Modifying();
        }
        private void textBoxControlLot_TextChanged(object sender, EventArgs e)
        {
            if (!InLoadingFile)
                Modifying();
        }
        private void textBoxControlNegative35_TextChanged(object sender, EventArgs e)
        {
            textBoxControlNegative35.ForeColor = Color.Black;
            if (textBoxControlNegative35.Text == "-" || textBoxControlNegative35.Text == "+")
            {
                if (!InLoadingFile)
                    textBoxControlBlank35.Focus();

                if (textBoxControlNegative35.Text == "+")
                {
                    textBoxControlNegative35.ForeColor = Color.Red;
                }
            }
            else
            {
                textBoxControlNegative35.Text = "";
                lblStatus.Text = "Only allowable characters are '+' and '-'";
            }

            if (!InLoadingFile)
                Modifying();
        }
        private void textBoxControlBath1Negative44_5_TextChanged(object sender, EventArgs e)
        {
            textBoxControlBath1Negative44_5.ForeColor = Color.Black;
            if (textBoxControlBath1Negative44_5.Text == "-" || textBoxControlBath1Negative44_5.Text == "+")
            {
                if (!InLoadingFile)
                    textBoxControlBath1Blank44_5.Focus();

                if (textBoxControlBath1Negative44_5.Text == "+")
                {
                    textBoxControlBath1Negative44_5.ForeColor = Color.Red;
                }
            }
            else
            {
                textBoxControlBath1Negative44_5.Text = "";
                lblStatus.Text = "Only allowable characters are '+' and '-'";
            }

            if (!InLoadingFile)
                Modifying();
        }
        private void textBoxControlBath2Negative44_5_TextChanged(object sender, EventArgs e)
        {
            textBoxControlBath2Negative44_5.ForeColor = Color.Black;
            if (textBoxControlBath2Negative44_5.Text == "-" || textBoxControlBath2Negative44_5.Text == "+")
            {
                if (!InLoadingFile)
                    textBoxControlBath2Blank44_5.Focus();

                if (textBoxControlBath2Negative44_5.Text == "+")
                {
                    textBoxControlBath2Negative44_5.ForeColor = Color.Red;
                }
            }
            else
            {
                textBoxControlBath2Negative44_5.Text = "";
                lblStatus.Text = "Only allowable characters are '+' and '-'";
            }

            if (!InLoadingFile)
                Modifying();
        }
        private void textBoxControlBath3Negative44_5_TextChanged(object sender, EventArgs e)
        {
            textBoxControlBath3Negative44_5.ForeColor = Color.Black;
            if (textBoxControlBath3Negative44_5.Text == "-" || textBoxControlBath3Negative44_5.Text == "+")
            {
                if (!InLoadingFile)
                    textBoxControlBath3Blank44_5.Focus();

                if (textBoxControlBath3Negative44_5.Text == "+")
                {
                    textBoxControlBath3Negative44_5.ForeColor = Color.Red;
                }
            }
            else
            {
                textBoxControlBath3Negative44_5.Text = "";
                lblStatus.Text = "Only allowable characters are '+' and '-'";
            }

            if (!InLoadingFile)
                Modifying();
        }
        private void textBoxControlNonTarget35_TextChanged(object sender, EventArgs e)
        {
            textBoxControlNonTarget35.ForeColor = Color.Black;
            if (textBoxControlNonTarget35.Text == "+" || textBoxControlNonTarget35.Text == "-" || textBoxControlNonTarget35.Text.ToUpper() == "N")
            {
                if (!InLoadingFile)
                    textBoxControlNegative35.Focus();

                if (textBoxControlNonTarget35.Text == "-")
                {
                    textBoxControlNonTarget35.ForeColor = Color.Red;
                }
                if (textBoxControlNonTarget35.Text.ToUpper() == "N")
                {
                    textBoxControlNonTarget35.Text = textBoxControlNonTarget35.Text.ToUpper();
                }

            }
            else
            {
                textBoxControlNonTarget35.Text = "";
                lblStatus.Text = "Only allowable characters are '+' , '-' or 'N' for not applicable";
            }

            if (!InLoadingFile)
                Modifying();
        }
        private void textBoxControlBath1NonTarget44_5_TextChanged(object sender, EventArgs e)
        {
            textBoxControlBath1NonTarget44_5.ForeColor = Color.Black;
            if (textBoxControlBath1NonTarget44_5.Text == "-" || textBoxControlBath1NonTarget44_5.Text == "+" || textBoxControlBath1NonTarget44_5.Text.ToUpper() == "N")
            {
                if (!InLoadingFile)
                    textBoxControlBath1Negative44_5.Focus();

                if (textBoxControlBath1NonTarget44_5.Text == "+")
                {
                    textBoxControlBath1NonTarget44_5.ForeColor = Color.Red;
                }
                if (textBoxControlBath1NonTarget44_5.Text.ToUpper() == "N")
                {
                    textBoxControlBath1NonTarget44_5.Text = textBoxControlBath1NonTarget44_5.Text.ToUpper();
                }
            }
            else
            {
                textBoxControlBath1NonTarget44_5.Text = "";
                lblStatus.Text = "Only allowable characters are '+' , '-' or 'N' for not applicable";
            }

            if (!InLoadingFile)
                Modifying();
        }
        private void textBoxControlBath2NonTarget44_5_TextChanged(object sender, EventArgs e)
        {
            textBoxControlBath2NonTarget44_5.ForeColor = Color.Black;
            if (textBoxControlBath2NonTarget44_5.Text == "-" || textBoxControlBath2NonTarget44_5.Text == "+" || textBoxControlBath2NonTarget44_5.Text.ToUpper() == "N")
            {
                if (!InLoadingFile)
                    textBoxControlBath2Negative44_5.Focus();

                if (textBoxControlBath2NonTarget44_5.Text == "+")
                {
                    textBoxControlBath2NonTarget44_5.ForeColor = Color.Red;
                }
                if (textBoxControlBath2NonTarget44_5.Text.ToUpper() == "N")
                {
                    textBoxControlBath2NonTarget44_5.Text = textBoxControlBath2NonTarget44_5.Text.ToUpper();
                }
            }
            else
            {
                textBoxControlBath2NonTarget44_5.Text = "";
                lblStatus.Text = "Only allowable characters are '+' , '-' or 'N' for not applicable";
            }

            if (!InLoadingFile)
                Modifying();
        }
        private void textBoxControlBath3NonTarget44_5_TextChanged(object sender, EventArgs e)
        {
            textBoxControlBath3NonTarget44_5.ForeColor = Color.Black;
            if (textBoxControlBath3NonTarget44_5.Text == "-" || textBoxControlBath3NonTarget44_5.Text == "+" || textBoxControlBath3NonTarget44_5.Text.ToUpper() == "N")
            {
                if (!InLoadingFile)
                    textBoxControlBath3Negative44_5.Focus();

                if (textBoxControlBath3NonTarget44_5.Text == "+")
                {
                    textBoxControlBath3NonTarget44_5.ForeColor = Color.Red;
                }
                if (textBoxControlBath3NonTarget44_5.Text.ToUpper() == "N")
                {
                    textBoxControlBath3NonTarget44_5.Text = textBoxControlBath3NonTarget44_5.Text.ToUpper();
                }
            }
            else
            {
                textBoxControlBath3NonTarget44_5.Text = "";
                lblStatus.Text = "Only allowable characters are '+' , '-' or 'N' for not applicable";
            }

            if (!InLoadingFile)
                Modifying();
        }
        private void textBoxControlPositive35_TextChanged(object sender, EventArgs e)
        {
            textBoxControlPositive35.ForeColor = Color.Black;
            if (textBoxControlPositive35.Text == "+" || textBoxControlPositive35.Text == "-")
            {
                if (!InLoadingFile)
                    textBoxControlNonTarget35.Focus();

                if (textBoxControlPositive35.Text == "-")
                {
                    textBoxControlPositive35.ForeColor = Color.Red;
                }
            }
            else
            {
                textBoxControlPositive35.Text = "";
                lblStatus.Text = "Only allowable characters are '+' and '-'";
            }

            if (!InLoadingFile)
                Modifying();
        }
        private void textBoxControlBath1Positive44_5_TextChanged(object sender, EventArgs e)
        {
            textBoxControlBath1Positive44_5.ForeColor = Color.Black;
            if (textBoxControlBath1Positive44_5.Text == "+" || textBoxControlBath1Positive44_5.Text == "-")
            {
                if (!InLoadingFile)
                    textBoxControlBath1NonTarget44_5.Focus();

                if (textBoxControlBath1Positive44_5.Text == "-")
                {
                    textBoxControlBath1Positive44_5.ForeColor = Color.Red;
                }
            }
            else
            {
                textBoxControlBath1Positive44_5.Text = "";
                lblStatus.Text = "Only allowable characters are '+' and '-'";
            }

            if (!InLoadingFile)
                Modifying();
        }
        private void textBoxControlBath2Positive44_5_TextChanged(object sender, EventArgs e)
        {
            textBoxControlBath2Positive44_5.ForeColor = Color.Black;
            if (textBoxControlBath2Positive44_5.Text == "+" || textBoxControlBath2Positive44_5.Text == "-")
            {
                if (!InLoadingFile)
                    textBoxControlBath2NonTarget44_5.Focus();

                if (textBoxControlBath2Positive44_5.Text == "-")
                {
                    textBoxControlBath2Positive44_5.ForeColor = Color.Red;
                }
            }
            else
            {
                textBoxControlBath2Positive44_5.Text = "";
                lblStatus.Text = "Only allowable characters are '+' and '-'";
            }

            if (!InLoadingFile)
                Modifying();
        }
        private void textBoxControlBath3Positive44_5_TextChanged(object sender, EventArgs e)
        {
            textBoxControlBath3Positive44_5.ForeColor = Color.Black;
            if (textBoxControlBath3Positive44_5.Text == "+" || textBoxControlBath3Positive44_5.Text == "-")
            {
                if (!InLoadingFile)
                    textBoxControlBath3NonTarget44_5.Focus();

                if (textBoxControlBath3Positive44_5.Text == "-")
                {
                    textBoxControlBath3Positive44_5.ForeColor = Color.Red;
                }
            }
            else
            {
                textBoxControlBath3Positive44_5.Text = "";
                lblStatus.Text = "Only allowable characters are '+' and '-'";
            }

            if (!InLoadingFile)
                Modifying();
        }
        private void textBoxDailyDuplicatePrecisionCriteria_TextChanged(object sender, EventArgs e)
        {
            CalculateDuplicate();

            if (!InLoadingFile)
                Modifying();
        }
        private void textBoxIncubationBath1StartTime_TextChanged(object sender, EventArgs e)
        {
            textBoxIncubationBath1StartTime.ForeColor = Color.Black;
            if (!CheckTimeInTextBox(textBoxIncubationBath1StartTime))
            {
                textBoxIncubationBath1StartTime.ForeColor = Color.Red;
            }
            else
            {
                if (textBoxIncubationBath1StartTime.Text.Length == 5)
                {
                    if (!InLoadingFile)
                        textBoxIncubationBath1EndTime.Focus();

                    TryToCalculateIncubationTimeSpan();
                }
                else
                {
                    textBoxIncubationBath1StartTime.ForeColor = Color.Red;
                }
            }

            if (!InLoadingFile)
                Modifying();
        }
        private void textBoxIncubationBath2StartTime_TextChanged(object sender, EventArgs e)
        {
            textBoxIncubationBath2StartTime.ForeColor = Color.Black;
            if (!CheckTimeInTextBox(textBoxIncubationBath2StartTime))
            {
                textBoxIncubationBath2StartTime.ForeColor = Color.Red;
            }
            else
            {
                if (textBoxIncubationBath2StartTime.Text.Length == 5)
                {
                    if (!InLoadingFile)
                        textBoxIncubationBath2EndTime.Focus();

                    TryToCalculateIncubationTimeSpan();
                }
                else
                {
                    textBoxIncubationBath2StartTime.ForeColor = Color.Red;
                }
            }

            if (!InLoadingFile)
                Modifying();
        }
        private void textBoxIncubationBath3StartTime_TextChanged(object sender, EventArgs e)
        {
            textBoxIncubationBath3StartTime.ForeColor = Color.Black;
            if (!CheckTimeInTextBox(textBoxIncubationBath3StartTime))
            {
                textBoxIncubationBath3StartTime.ForeColor = Color.Red;
            }
            else
            {
                if (textBoxIncubationBath3StartTime.Text.Length == 5)
                {
                    if (!InLoadingFile)
                        textBoxIncubationBath3EndTime.Focus();

                    TryToCalculateIncubationTimeSpan();
                }
                else
                {
                    textBoxIncubationBath3StartTime.ForeColor = Color.Red;
                }
            }

            if (!InLoadingFile)
                Modifying();
        }
        private void textBoxIncubationBath1EndTime_TextChanged(object sender, EventArgs e)
        {
            textBoxIncubationBath1EndTime.ForeColor = Color.Black;
            if (!CheckTimeInTextBox(textBoxIncubationBath1EndTime))
            {
                textBoxIncubationBath1EndTime.ForeColor = Color.Red;
            }
            else
            {
                if (textBoxIncubationBath1EndTime.Text.Length == 5)
                {
                    if (!InLoadingFile)
                        textBoxWaterBath1Number.Focus();

                    TryToCalculateIncubationTimeSpan();
                }
                else
                {
                    textBoxIncubationBath1EndTime.ForeColor = Color.Red;
                }
            }

            if (!InLoadingFile)
                Modifying();
        }
        private void textBoxIncubationBath2EndTime_TextChanged(object sender, EventArgs e)
        {
            textBoxIncubationBath2EndTime.ForeColor = Color.Black;
            if (!CheckTimeInTextBox(textBoxIncubationBath2EndTime))
            {
                textBoxIncubationBath2EndTime.ForeColor = Color.Red;
            }
            else
            {
                if (textBoxIncubationBath2EndTime.Text.Length == 5)
                {
                    if (!InLoadingFile)
                        textBoxWaterBath2Number.Focus();

                    TryToCalculateIncubationTimeSpan();
                }
                else
                {
                    textBoxIncubationBath2EndTime.ForeColor = Color.Red;
                }
            }

            if (!InLoadingFile)
                Modifying();
        }
        private void textBoxIncubationBath3EndTime_TextChanged(object sender, EventArgs e)
        {
            textBoxIncubationBath3EndTime.ForeColor = Color.Black;
            if (!CheckTimeInTextBox(textBoxIncubationBath3EndTime))
            {
                textBoxIncubationBath3EndTime.ForeColor = Color.Red;
            }
            else
            {
                if (textBoxIncubationBath3EndTime.Text.Length == 5)
                {
                    if (!InLoadingFile)
                        textBoxWaterBath3Number.Focus();

                    TryToCalculateIncubationTimeSpan();
                }
                else
                {
                    textBoxIncubationBath3EndTime.ForeColor = Color.Red;
                }
            }

            if (!InLoadingFile)
                Modifying();
        }
        private void textBoxIntertechDuplicatePrecisionCriteria_TextChanged(object sender, EventArgs e)
        {
            CalculateDuplicate();

            if (!InLoadingFile)
                Modifying();
        }
        private void textBoxLot35_TextChanged(object sender, EventArgs e)
        {
            if (!InLoadingFile)
                Modifying();
        }
        private void textBoxLot44_5_TextChanged(object sender, EventArgs e)
        {
            if (!InLoadingFile)
                Modifying();
        }
        private void textBoxResultsReadBy_TextChanged(object sender, EventArgs e)
        {
            if (!InLoadingFile)
                Modifying();
        }
        private void textBoxResultsRecordedBy_TextChanged(object sender, EventArgs e)
        {
            if (!InLoadingFile)
                Modifying();
        }
        private void textBoxSalinitiesReadBy_TextChanged(object sender, EventArgs e)
        {
            if (!InLoadingFile)
                Modifying();
        }
        private void textBoxSampleBottleLotNumber_TextChanged(object sender, EventArgs e)
        {
            if (!InLoadingFile)
                Modifying();
        }
        private void textBoxSampleCrewInitials_TextChanged(object sender, EventArgs e)
        {
            textBoxSampleCrewInitials.ForeColor = Color.Black;
            foreach (char c in textBoxSampleCrewInitials.Text)
            {
                if (char.IsLetter(c) || char.IsPunctuation(c) || char.IsWhiteSpace(c))
                {
                }
                else
                {
                    textBoxSampleCrewInitials.ForeColor = Color.Red;
                }
            }

            if (!InLoadingFile)
                Modifying();
        }
        private void textBoxAccessCode_TextChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBoxInitials.Text))
            {
                MessageBox.Show("Please enter initial before code.");
                return;
            }

            if (textBoxAccessCode.Text == csspWQInputApp.AccessCode)
            {

                if (csspWQInputApp.ApprovalCode.Length > 0 && csspWQInputApp.ApprovalCode == textBoxApprovalCode.Text)
                {
                    butApprove.Enabled = true;
                }
                else
                {
                    butApprove.Enabled = false;
                }

                SetupCSSPWQInputTool();
                dateTimePickerRun.Value = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
            }
        }
        private void textBoxTCField1_TextChanged(object sender, EventArgs e)
        {
            textBoxTCField1.ForeColor = Color.Black;
            foreach (char c in textBoxTCField1.Text)
            {
                if (char.IsNumber(c) || c.ToString() == ".")
                {
                }
                else
                {
                    textBoxTCField1.ForeColor = Color.Red;
                    return;
                }
            }
            float TCField = -99.0f;
            float.TryParse(textBoxTCField1.Text, out TCField);

            textBoxTCField1.ForeColor = Color.Black;
            if (TCField == -99.0f)
            {
                textBoxTCField1.ForeColor = Color.Red;
                return;
            }
            if (TCField < 0 || TCField > 35)
            {
                textBoxTCField1.ForeColor = Color.Red;
                return;
            }

            if (!InLoadingFile)
                Modifying();
        }
        private void textBoxTCField2_TextChanged(object sender, EventArgs e)
        {
            textBoxTCField2.ForeColor = Color.Black;
            foreach (char c in textBoxTCField2.Text)
            {
                if (char.IsNumber(c) || c.ToString() == ".")
                {
                }
                else
                {
                    textBoxTCField2.ForeColor = Color.Red;
                    return;
                }
            }
            float TCField = -99.0f;
            float.TryParse(textBoxTCField2.Text, out TCField);

            textBoxTCField2.ForeColor = Color.Black;
            if (TCField == -99.0f)
            {
                textBoxTCField2.ForeColor = Color.Red;
                return;
            }
            if (TCField < 0 || TCField > 35)
            {
                textBoxTCField2.ForeColor = Color.Red;
                return;
            }

            if (!InLoadingFile)
                Modifying();
        }
        private void textBoxTCLab1_TextChanged(object sender, EventArgs e)
        {
            textBoxTCLab1.ForeColor = Color.Black;
            foreach (char c in textBoxTCLab1.Text)
            {
                if (char.IsNumber(c) || c.ToString() == ".")
                {
                }
                else
                {
                    textBoxTCLab1.ForeColor = Color.Red;
                    return;
                }
            }
            float TCLab = -99.0f;
            float.TryParse(textBoxTCLab1.Text, out TCLab);

            textBoxTCLab1.ForeColor = Color.Black;
            if (TCLab == -99.0f)
            {
                textBoxTCLab1.ForeColor = Color.Red;
                return;
            }
            if (TCLab < 0 || TCLab > 8.50f)
            {
                textBoxTCLab1.ForeColor = Color.Red;
                return;
            }

            if (!InLoadingFile)
                Modifying();
        }
        private void textBoxTCLab2_TextChanged(object sender, EventArgs e)
        {
            textBoxTCLab2.ForeColor = Color.Black;
            foreach (char c in textBoxTCLab2.Text)
            {
                if (char.IsNumber(c) || c.ToString() == ".")
                {
                }
                else
                {
                    textBoxTCLab2.ForeColor = Color.Red;
                    return;
                }
            }
            float TCLab = -99.0f;
            float.TryParse(textBoxTCLab2.Text, out TCLab);

            textBoxTCLab2.ForeColor = Color.Black;
            if (TCLab == -99.0f)
            {
                textBoxTCLab2.ForeColor = Color.Red;
                return;
            }
            if (TCLab < 0 || TCLab > 8.50f)
            {
                textBoxTCLab2.ForeColor = Color.Red;
                return;
            }

            if (!InLoadingFile)
                Modifying();
        }
        private void textBoxTides_TextChanged(object sender, EventArgs e)
        {
            textBoxTides.ForeColor = Color.Black;
            foreach (char c in textBoxTides.Text)
            {
                if (char.IsLetter(c) || char.IsPunctuation(c) || char.IsWhiteSpace(c))
                {
                }
                else
                {
                    textBoxTides.ForeColor = Color.Red;
                    return;
                }
            }

            if (!InLoadingFile)
                Modifying();
        }
        private void textBoxWaterBath1Number_TextChanged(object sender, EventArgs e)
        {
            if (!InLoadingFile)
                Modifying();
        }
        private void textBoxWaterBath2Number_TextChanged(object sender, EventArgs e)
        {
            if (!InLoadingFile)
                Modifying();
        }
        private void textBoxWaterBath3Number_TextChanged(object sender, EventArgs e)
        {
            if (!InLoadingFile)
                Modifying();
        }
        #endregion Events TextChanged
        #region Events timerGetTides
        private void timerGetTides_Tick(object sender, EventArgs e)
        {
            GetTides();
        }
        #endregion Events timerGetTides
        #region Events timerSave
        private void timerSave_Tick(object sender, EventArgs e)
        {
            DoSave();
        }
        #endregion timerDataGridViewEditCheck
        #region Events WebBrowserCSSP
        private void webBrowserCSSP_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (e.Url.AbsolutePath != (sender as WebBrowser).Url.AbsolutePath)
                return;

            if (webBrowserCSSP.Url.ToString().Contains(@"http://www.tides.gc.ca/eng"))
            {
                textBoxTides.Text = GetTideText();
                if (textBoxTides.Text == "-- / --")
                {
                    TideToTryIndex += 1;
                    if (TideToTryIndex > 2)
                    {
                        Modifying();
                        if (labSheetA1Sheet.Tides != textBoxTides.Text)
                        {
                            labSheetA1Sheet.Tides = textBoxTides.Text;
                            AddLog("Tides", textBoxTides.Text);
                        }
                        return;
                    }
                    timerGetTides.Enabled = true;
                }
                Modifying();
                if (labSheetA1Sheet.Tides != textBoxTides.Text)
                {
                    labSheetA1Sheet.Tides = textBoxTides.Text;
                    AddLog("Tides", textBoxTides.Text);
                }
            }
        }
        #endregion Events WebBrowserCSSP
        #endregion Events

        #region Functions
        #region form Functions
        private void AddLog(string Element, string NewValue)
        {
            sbLog.AppendLine(DateTime.Now + "\t" + Initials + "\t|||" + Element + "\t|||" + NewValue);
        }
        private void AdjustVisualForIncludeLaboratoryQAQC()
        {
            if (csspWQInputApp.IncludeLaboratoryQAQC)
            {
                int PanelHeight = 169;
                panelAppInputTopTideCrew.Width = 130;
                panelAppInputTopTideCrew.Height = 160;
                textBoxTides.Top = 30;
                textBoxTides.Left = 16;
                lblSampleCrewInitials.Visible = true;
                textBoxSampleCrewInitials.Visible = true;
                butViewFCForm.Visible = true;
                butFail.Visible = true;
                panelTC.Visible = true;
                panelAppInputTopIncubation.Visible = true;
                panelControl.Visible = true;
                panelAddInputBottomLeftDuplicate.Visible = true;
                panelAddInputBottomRight.Visible = true;
                panelTC.Height = PanelHeight;
                panelAppInputTopIncubation.Height = PanelHeight;
                panelControl.Height = PanelHeight;
                panelAppInputTop.Height = PanelHeight;
                lblApprovalCode.Visible = true;
                textBoxApprovalCode.Visible = true;
                lblSupervisorOnly.Visible = true;
            }
            else
            {
                int PanelHeight = 55;
                panelAppInputTopTideCrew.Width = panelAppInputTop.Width - 30;
                textBoxTides.Top = butGetTides.Top + 3;
                textBoxTides.Left = butGetTides.Right + 4;
                lblSampleCrewInitials.Visible = false;
                textBoxSampleCrewInitials.Visible = false;
                butViewFCForm.Visible = false;
                butFail.Visible = false;
                panelAppInputTopTideCrew.Height = PanelHeight - 10;
                panelTC.Visible = false;
                panelAppInputTopIncubation.Visible = false;
                panelControl.Visible = false;
                panelAddInputBottomLeftDuplicate.Visible = false;
                panelAddInputBottomRight.Visible = false;
                panelTC.Height = PanelHeight;
                panelAppInputTopIncubation.Height = PanelHeight;
                panelControl.Height = PanelHeight;
                panelAppInputTop.Height = PanelHeight;
                lblApprovalCode.Visible = false;
                textBoxApprovalCode.Visible = false;
                lblSupervisorOnly.Visible = false;
            }
        }
        private void Approve()
        {
            if (string.IsNullOrWhiteSpace(csspWQInputApp.ApprovalCode))
            {
                MessageBox.Show("Can not approve lab sheet", "Approval Code is empty", MessageBoxButtons.OK);
                return;
            }

            if (csspWQInputApp.ApprovalCode != textBoxApprovalCode.Text)
            {
                MessageBox.Show("Could not approve the lab sheet", "Approval Code error", MessageBoxButtons.OK);
                return;
            }

            lblSupervisorInitials.Text = "";
            lblSupervisorInitials.Text = textBoxInitials.Text;
            ApprovalSupervisorInitials = textBoxInitials.Text;

            csspWQInputApp.ApprovalDate = DateTime.Now;
            lblApprovalDate.Text = DateTime.Now.ToString("yyyy MMMM dd");
            Modifying();
        }
        private void CalculateDuplicate()
        {
            int SiteColumn = 1;
            int MPNColumn = 3;
            int SampleTypeColumn = 10;
            switch (csspWQInputSheetType)
            {
                case CSSPWQInputSheetTypeEnum.A1:
                    {
                        MPNColumn = 3;
                    }
                    break;
                case CSSPWQInputSheetTypeEnum.LTB:
                    {
                        MPNColumn = 99;
                    }
                    break;
                case CSSPWQInputSheetTypeEnum.EC:
                    {
                        MPNColumn = 99;
                    }
                    break;
                default:
                    {
                        MPNColumn = 99;
                    }
                    break;
            }

            // Calculating DailyDuplicate
            int DailyDuplicateRow = -1;
            int DailyDuplicateMPN = 0;
            for (int i = 0, count = dataGridViewCSSP.Rows.Count; i < count; i++)
            {
                if (dataGridViewCSSP[SampleTypeColumn, i].Value.ToString() == SampleTypeEnum.DailyDuplicate.ToString())
                {
                    DailyDuplicateRow = i;
                    break;
                }
            }
            if (DailyDuplicateRow != -1)
            {
                if (!string.IsNullOrWhiteSpace(dataGridViewCSSP[MPNColumn, DailyDuplicateRow].Value.ToString()))
                {
                    int.TryParse(dataGridViewCSSP[MPNColumn, DailyDuplicateRow].Value.ToString(), out DailyDuplicateMPN);

                    if (DailyDuplicateMPN != 0)
                    {
                        if (dataGridViewCSSP[SampleTypeColumn, DailyDuplicateRow].Value.ToString() == SampleTypeEnum.DailyDuplicate.ToString())
                        {
                            for (int i = 0, count = dataGridViewCSSP.Rows.Count; i < count; i++)
                            {
                                if (DailyDuplicateRow != i)
                                {
                                    if (dataGridViewCSSP[SiteColumn, i].Value.ToString() == dataGridViewCSSP[SiteColumn, DailyDuplicateRow].Value.ToString()
                                        && !(dataGridViewCSSP[SampleTypeColumn, i].Value.ToString() == SampleTypeEnum.DailyDuplicate.ToString()
                                        || dataGridViewCSSP[SampleTypeColumn, i].Value.ToString() == SampleTypeEnum.IntertechDuplicate.ToString()
                                        || dataGridViewCSSP[SampleTypeColumn, i].Value.ToString() == SampleTypeEnum.IntertechRead.ToString()))
                                    {
                                        int OtherMPN = 0;
                                        int.TryParse(dataGridViewCSSP[MPNColumn, i].Value.ToString(), out OtherMPN);
                                        if (OtherMPN != 0)
                                        {
                                            float OtherMPNLog = (float)Math.Log10((double)OtherMPN);
                                            float MPNLog = (float)Math.Log10((double)DailyDuplicateMPN);

                                            float MPNLogDiff = Math.Abs(OtherMPNLog - MPNLog);

                                            lblDailyDuplicateRLogValue.Text = MPNLogDiff.ToString("F5");

                                            if (textBoxDailyDuplicatePrecisionCriteria.Text.Length > 0)
                                            {
                                                float MPNLogCriteria;
                                                float.TryParse(textBoxDailyDuplicatePrecisionCriteria.Text, out MPNLogCriteria);

                                                if (MPNLogCriteria > 0)
                                                {
                                                    if (MPNLogDiff > MPNLogCriteria)
                                                    {
                                                        lblDailyDuplicateAcceptableOrUnacceptable.ForeColor = Color.Red;
                                                        lblDailyDuplicateAcceptableOrUnacceptable.Text = "Unacceptable";
                                                    }
                                                    else
                                                    {
                                                        lblDailyDuplicateAcceptableOrUnacceptable.ForeColor = Color.Green;
                                                        lblDailyDuplicateAcceptableOrUnacceptable.Text = "Acceptable";
                                                    }
                                                }
                                                else
                                                {
                                                    lblDailyDuplicateAcceptableOrUnacceptable.ForeColor = Color.Blue;
                                                    lblDailyDuplicateAcceptableOrUnacceptable.Text = "ERROR";

                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            // Calculating IntertechDuplicate
            int IntertechDuplicateRow = -1;
            int IntertechDuplicateMPN = 0;
            for (int i = 0, count = dataGridViewCSSP.Rows.Count; i < count; i++)
            {
                if (dataGridViewCSSP[SampleTypeColumn, i].Value.ToString() == SampleTypeEnum.IntertechDuplicate.ToString())
                {
                    IntertechDuplicateRow = i;
                    break;
                }
            }
            if (IntertechDuplicateRow != -1)
            {
                if (!string.IsNullOrWhiteSpace(dataGridViewCSSP[MPNColumn, IntertechDuplicateRow].Value.ToString()))
                {
                    int.TryParse(dataGridViewCSSP[MPNColumn, IntertechDuplicateRow].Value.ToString(), out IntertechDuplicateMPN);

                    if (IntertechDuplicateMPN != 0)
                    {
                        if (dataGridViewCSSP[SampleTypeColumn, IntertechDuplicateRow].Value.ToString() == SampleTypeEnum.IntertechDuplicate.ToString())
                        {
                            for (int i = 0, count = dataGridViewCSSP.Rows.Count; i < count; i++)
                            {
                                if (IntertechDuplicateRow != i)
                                {
                                    if (dataGridViewCSSP[SiteColumn, i].Value.ToString() == dataGridViewCSSP[SiteColumn, IntertechDuplicateRow].Value.ToString()
                                        && !(dataGridViewCSSP[SampleTypeColumn, i].Value.ToString() == SampleTypeEnum.DailyDuplicate.ToString()
                                        || dataGridViewCSSP[SampleTypeColumn, i].Value.ToString() == SampleTypeEnum.IntertechDuplicate.ToString()
                                        || dataGridViewCSSP[SampleTypeColumn, i].Value.ToString() == SampleTypeEnum.IntertechRead.ToString()))
                                    {
                                        int OtherMPN = 0;
                                        int.TryParse(dataGridViewCSSP[MPNColumn, i].Value.ToString(), out OtherMPN);
                                        if (OtherMPN != 0)
                                        {
                                            float OtherMPNLog = (float)Math.Log10((double)OtherMPN);
                                            float MPNLog = (float)Math.Log10((double)IntertechDuplicateMPN);

                                            float MPNLogDiff = Math.Abs(OtherMPNLog - MPNLog);

                                            lblIntertechDuplicateRLogValue.Text = MPNLogDiff.ToString("F5");

                                            if (textBoxIntertechDuplicatePrecisionCriteria.Text.Length > 0)
                                            {
                                                float MPNLogCriteria;
                                                float.TryParse(textBoxIntertechDuplicatePrecisionCriteria.Text, out MPNLogCriteria);

                                                if (MPNLogCriteria > 0)
                                                {
                                                    if (MPNLogDiff > MPNLogCriteria)
                                                    {
                                                        lblIntertechDuplicateAcceptableOrUnacceptable.ForeColor = Color.Red;
                                                        lblIntertechDuplicateAcceptableOrUnacceptable.Text = "Unacceptable";
                                                    }
                                                    else
                                                    {
                                                        lblIntertechDuplicateAcceptableOrUnacceptable.ForeColor = Color.Green;
                                                        lblIntertechDuplicateAcceptableOrUnacceptable.Text = "Acceptable";
                                                    }
                                                }
                                                else
                                                {
                                                    lblIntertechDuplicateAcceptableOrUnacceptable.ForeColor = Color.Blue;
                                                    lblIntertechDuplicateAcceptableOrUnacceptable.Text = "ERROR";

                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            // Calculating IntertechRead
            int IntertechReadRow = -1;
            int IntertechReadMPN = 0;
            for (int i = 0, count = dataGridViewCSSP.Rows.Count; i < count; i++)
            {
                if (dataGridViewCSSP[SampleTypeColumn, i].Value.ToString() == SampleTypeEnum.IntertechRead.ToString())
                {
                    IntertechReadRow = i;
                    break;
                }
            }
            if (IntertechReadRow != -1)
            {
                if (!string.IsNullOrWhiteSpace(dataGridViewCSSP[MPNColumn, IntertechReadRow].Value.ToString()))
                {
                    int.TryParse(dataGridViewCSSP[MPNColumn, IntertechReadRow].Value.ToString(), out IntertechReadMPN);

                    if (IntertechReadMPN != 0)
                    {
                        if (dataGridViewCSSP[SampleTypeColumn, IntertechReadRow].Value.ToString() == SampleTypeEnum.IntertechRead.ToString())
                        {
                            for (int i = 0, count = dataGridViewCSSP.Rows.Count; i < count; i++)
                            {
                                if (IntertechReadRow != i)
                                {
                                    if (dataGridViewCSSP[SiteColumn, i].Value.ToString() == dataGridViewCSSP[SiteColumn, IntertechReadRow].Value.ToString()
                                        && !(dataGridViewCSSP[SampleTypeColumn, i].Value.ToString() == SampleTypeEnum.DailyDuplicate.ToString()
                                        || dataGridViewCSSP[SampleTypeColumn, i].Value.ToString() == SampleTypeEnum.IntertechDuplicate.ToString()
                                        || dataGridViewCSSP[SampleTypeColumn, i].Value.ToString() == SampleTypeEnum.IntertechRead.ToString()))
                                    {
                                        int OtherMPN = 0;
                                        int.TryParse(dataGridViewCSSP[MPNColumn, i].Value.ToString(), out OtherMPN);
                                        if (OtherMPN != 0)
                                        {
                                            if (OtherMPN == IntertechReadMPN)
                                            {
                                                lblIntertechReadAcceptableOrUnacceptable.ForeColor = Color.Green;
                                                lblIntertechReadAcceptableOrUnacceptable.Text = "Acceptable";
                                            }
                                            else
                                            {

                                                lblIntertechReadAcceptableOrUnacceptable.ForeColor = Color.Red;
                                                lblIntertechReadAcceptableOrUnacceptable.Text = "Unacceptable";
                                            }
                                        }
                                    }

                                }
                            }
                        }
                    }
                }
            }
        }
        private void CalculateTCFirstAndAverage()
        {
            int TimeColumn = -1;
            int TempColumn = -1;
            int RowWithSmallestTime = -1;
            List<float> TempList = new List<float>();
            switch (csspWQInputSheetType)
            {
                case CSSPWQInputSheetTypeEnum.A1:
                    {
                        TimeColumn = 2;
                        TempColumn = 7;
                    }
                    break;
                case CSSPWQInputSheetTypeEnum.LTB:
                    {
                        TimeColumn = 99;
                        TempColumn = 99;
                    }
                    break;
                case CSSPWQInputSheetTypeEnum.EC:
                    {
                        TimeColumn = 99;
                        TempColumn = 99;
                    }
                    break;
                default:
                    break;
            }
            int MinHour = 24;
            int MinMinute = 60;
            for (int i = 0, count = dataGridViewCSSP.Rows.Count; i < count; i++)
            {
                if (dataGridViewCSSP[TempColumn, i].Value == null || string.IsNullOrWhiteSpace(dataGridViewCSSP[TempColumn, i].Value.ToString()))
                {
                    // nothing
                }
                else
                {
                    TempList.Add(float.Parse(dataGridViewCSSP[TempColumn, i].Value.ToString()));
                }
                if (dataGridViewCSSP[TimeColumn, i].Value == null)
                    continue;

                if (!string.IsNullOrWhiteSpace(dataGridViewCSSP[TimeColumn, i].Value.ToString()) && !string.IsNullOrWhiteSpace(dataGridViewCSSP[TimeColumn, i].Value.ToString()))
                {
                    string TimeStr = dataGridViewCSSP[TimeColumn, i].Value.ToString();
                    if (TimeStr.Length == 5)
                    {
                        int Hour = int.Parse(TimeStr.Substring(0, 2));
                        int Minute = int.Parse(TimeStr.Substring(3, 2));

                        if (Hour <= MinHour)
                        {
                            if (Hour == MinHour)
                            {
                                if (Minute < MinMinute)
                                {
                                    MinMinute = Minute;
                                    RowWithSmallestTime = i;
                                }
                            }
                            else
                            {
                                MinHour = Hour;
                                MinMinute = 60;
                                if (Minute < MinMinute)
                                {
                                    MinMinute = Minute;
                                    RowWithSmallestTime = i;
                                }
                            }
                        }
                    }
                }
            }
            if (RowWithSmallestTime < 0)
            {
                lblTCFirst.Text = "---";
                lblTCAverage.Text = "---";
                return;
            }
            if (dataGridViewCSSP[TempColumn, RowWithSmallestTime].Value != null)
            {
                string FirstTemp = dataGridViewCSSP[TempColumn, RowWithSmallestTime].Value.ToString();
                if (FirstTemp.Length > 0)
                {
                    lblTCFirst.Text = FirstTemp;
                }
                else
                {
                    lblTCFirst.Text = "---";
                }
            }
            if (TempList.Count > 0)
            {
                lblTCAverage.Text = (from c in TempList
                                     select c).Average().ToString("F1");
            }
        }
        private void CancelSendToServer()
        {
            panelSendToServerCompare.SendToBack();
            butSendToServer.Enabled = true;
        }
        private void ContinueSendToServer()
        {
            SendToServer();
            panelSendToServerCompare.SendToBack();
        }
        private FileInfo CanChangeDate()
        {
            butChangeDate.Enabled = false;
            lblChangeDateError.Text = "";
            int year = dateTimePickerChangeDate.Value.Year;
            int month = dateTimePickerChangeDate.Value.Month;
            int day = dateTimePickerChangeDate.Value.Day;
            string YearMonthDayCurrent2 = dateTimePickerRun.Value.Year + "_" + (month > 9 ? month.ToString() : "0" + month) + "_" + (day > 9 ? day.ToString() : "0" + day);

            if (dateTimePickerRun.Value == dateTimePickerChangeDate.Value)
            {
                lblChangeDateError.Text = "Same Date";
                butChangeDate.Enabled = false;
                return null;
            }

            FileInfo fi = new FileInfo(CurrentPath);
            if (NameCurrent.Contains(" "))
            {
                fi = new FileInfo(CurrentPath + YearMonthDayCurrent2.Substring(0, 4) + @"\" + NameCurrent.Substring(0, NameCurrent.IndexOf(" ")) + "_" + YearMonthDayCurrent2 + "_" + csspWQInputSheetType.ToString() + "_R" + RunNumberCurrent + "_C.txt");
            }
            else
            {
                fi = new FileInfo(CurrentPath + YearMonthDayCurrent2.Substring(0, 4) + @"\" + NameCurrent + YearMonthDayCurrent2 + "_" + csspWQInputSheetType.ToString() + "_R" + RunNumberCurrent + "_C.txt");
            }

            if (fi.Exists)
            {
                lblChangeDateError.Text = "File already exist";
                butChangeDate.Enabled = false;
                return null;
            }

            butChangeDate.Enabled = true;
            return fi;
        }
        private void ChangeDate()
        {
            FileInfo fi = CanChangeDate();
            if (fi == null)
            {
                butChangeDate.Enabled = false;
                return;
            }

            FileInfo fiOld = new FileInfo(lblFilePath.Text);

            if (!fiOld.Exists)
            {
                lblChangeDateError.Text = "Could not find file to change date";
                butChangeDate.Enabled = false;
                return;
            }

            try
            {
                File.Copy(fiOld.FullName, fi.FullName, false);
            }
            catch (Exception ex)
            {
                lblChangeDateError.Text = "Could not copy file." + ex.Message;
                butChangeDate.Enabled = false;
                return;
            }

            try
            {
                DirectoryInfo di = new DirectoryInfo(CurrentPath + YearMonthDayCurrent.Substring(0, 4) + @"\" + @"NotUsed\");
                if (!di.Exists)
                {
                    di.Create();
                }
            }
            catch (Exception ex)
            {
                lblChangeDateError.Text = "Could not create NotUsed directory." + ex.Message;
                butChangeDate.Enabled = false;
                return;
            }

            FileInfo fiNotUsed = new FileInfo(CurrentPath + YearMonthDayCurrent.Substring(0, 4) + @"\" + @"NotUsed\" + fiOld.Name);

            int count = 0;
            while (fiNotUsed.Exists)
            {
                count += 1;
                fiNotUsed = new FileInfo(fiNotUsed.FullName.Replace(".txt", "_" + count.ToString() + ".txt"));
            }

            try
            {
                File.Copy(fiOld.FullName, fiNotUsed.FullName, false);
                fiOld.Delete();
            }
            catch (Exception ex)
            {
                lblChangeDateError.Text = "Could not copy old file to NotUsed Directory." + ex.Message;
                butChangeDate.Enabled = false;
                return;
            }

            try
            {
                fiOld.Delete();
            }
            catch (Exception ex)
            {
                lblChangeDateError.Text = "Could not remove old file." + ex.Message;
                butChangeDate.Enabled = false;
                return;
            }

            lblFilePath.Text = fi.FullName;
            SaveInfoOnLocalMachine(true);
            dateTimePickerRun.Value = dateTimePickerChangeDate.Value;
            dateTimePickerSalinitiesReadDate.Value = dateTimePickerRun.Value;
            dateTimePickerResultsReadDate.Value = dateTimePickerRun.Value.AddDays(1);
            dateTimePickerResultsRecordedDate.Value = dateTimePickerRun.Value.AddDays(1);
            panelAppInput.BringToFront();
            CurrentPanel = panelAppInput;
        }
        private string CheckDestinationFilesDocx(FileInfo fi, FileInfo fiTo)
        {
            List<FileInfo> fiToList = (from c in fiTo.Directory.GetFiles()
                                       from p in PossibleLabSheetFileNamesDocx
                                       where fiTo.FullName.StartsWith(c.FullName.Substring(0, c.FullName.Length - 7))
                                       && c.FullName.Substring(c.FullName.Length - 7) == p
                                       && c.FullName.EndsWith(".docx")
                                       select c).ToList();

            if (fiToList.Count == 0)
                return "";

            if (fi.FullName.Substring(fi.FullName.Length - 6) != fiToList[0].FullName.Substring(fiToList[0].FullName.Length - 6))
            {
                if (fi.LastWriteTimeUtc >= fiToList[0].LastWriteTimeUtc)
                {
                    if (fiToList[0].Exists)
                    {
                        File.Copy(fiToList[0].FullName, fiToList[0].FullName.Replace(fiToList[0].FullName.Substring(fiToList[0].FullName.Length - 6), fi.FullName.Substring(fi.FullName.Length - 6)));
                        fiToList[0].Delete();
                        return "";
                    }
                }
                else
                {
                    return "Don't copy";
                }
            }

            return "";
        }
        private string CheckDestinationFilesTxt(FileInfo fi, FileInfo fiTo)
        {
            List<FileInfo> fiToList = (from c in fiTo.Directory.GetFiles()
                                       from p in PossibleLabSheetFileNamesTxt
                                       where fiTo.FullName.StartsWith(c.FullName.Substring(0, c.FullName.Length - 6))
                                       && c.FullName.Substring(c.FullName.Length - 6) == p
                                       && c.FullName.EndsWith(".txt")
                                       select c).ToList();

            if (fiToList.Count == 0)
                return "";

            if (fi.FullName.Substring(fi.FullName.Length - 6) != fiToList[0].FullName.Substring(fiToList[0].FullName.Length - 6))
            {
                if (fi.LastWriteTimeUtc >= fiToList[0].LastWriteTimeUtc)
                {
                    if (fiToList[0].Exists)
                    {
                        File.Copy(fiToList[0].FullName, fiToList[0].FullName.Replace(fiToList[0].FullName.Substring(fiToList[0].FullName.Length - 6), fi.FullName.Substring(fi.FullName.Length - 6)));
                        fiToList[0].Delete();
                        return "";
                    }
                }
                else
                {
                    return "Don't copy";
                }
            }

            return "";
        }
        private bool CheckFollowingAndCount(StreamReader sr, int LineNumber, string OldFirstObj, List<string> ValueArr, string ToFollow, int count)
        {
            if (OldFirstObj != ToFollow)
            {
                PostErrorWhileReadingFile(sr, LineNumber, ValueArr[0] + " has to be following " + ToFollow + " in the file.");
                return false;
            }
            if (ValueArr.Count != count)
            {
                PostErrorWhileReadingFile(sr, LineNumber, ValueArr[0] + " line does not have " + count + " items.");
                return false;
            }

            return true;
        }
        private bool CheckTimeInTextBox(TextBox textBoxTemp)
        {
            foreach (char c in textBoxTemp.Text)
            {
                if (char.IsNumber(c) || c.ToString() == ":")
                {
                }
                else
                {
                    return false;
                }
            }
            if (textBoxTemp.Text.Length < 4)
            {
                return true;
            }
            if (textBoxTemp.Text.Length == 4)
            {
                if (!textBoxTemp.Text.Contains(":"))
                {
                    textBoxTemp.Text = textBoxTemp.Text.Substring(0, 2) + ":" + textBoxTemp.Text.Substring(2, 2);
                }
            }
            if (textBoxTemp.Text.Length == 5 && textBoxTemp.Text.Substring(2, 1) == ":")
            {
                if (!(int.Parse(textBoxTemp.Text.Substring(0, 2)) >= 0) || !(int.Parse(textBoxTemp.Text.Substring(0, 2)) <= 23))
                {
                    textBoxTemp.ForeColor = Color.Red;
                    return false;
                }
                if (!(int.Parse(textBoxTemp.Text.Substring(3, 2)) >= 0) || !(int.Parse(textBoxTemp.Text.Substring(3, 2)) <= 59))
                {
                    textBoxTemp.ForeColor = Color.Red;
                    return false;
                }
            }
            else
            {
                textBoxTemp.ForeColor = Color.Red;
                return false;
            }

            return true;
        }
        private void CleanAppInputPanel()
        {
            textBoxSampleCrewInitials.Text = "";
            textBoxIncubationBath1StartTime.Text = "";
            textBoxIncubationBath2StartTime.Text = "";
            textBoxIncubationBath3StartTime.Text = "";
            textBoxIncubationBath1EndTime.Text = "";
            textBoxIncubationBath2EndTime.Text = "";
            textBoxIncubationBath3EndTime.Text = "";
            textBoxWaterBath1Number.Text = "";
            textBoxWaterBath2Number.Text = "";
            textBoxWaterBath3Number.Text = "";
            textBoxTCField1.Text = "";
            textBoxTCLab1.Text = "";
            textBoxTCField2.Text = "";
            textBoxTCLab2.Text = "";
            checkBox2Coolers.Checked = false;
            radioButton1Baths.Checked = true;
            textBoxControlLot.Text = "";
            textBoxControlPositive35.Text = "";
            textBoxControlNonTarget35.Text = "";
            textBoxControlNegative35.Text = "";
            textBoxControlBath1Positive44_5.Text = "";
            textBoxControlBath2Positive44_5.Text = "";
            textBoxControlBath3Positive44_5.Text = "";
            textBoxControlBath1NonTarget44_5.Text = "";
            textBoxControlBath2NonTarget44_5.Text = "";
            textBoxControlBath3NonTarget44_5.Text = "";
            textBoxControlBath1Negative44_5.Text = "";
            textBoxControlBath2Negative44_5.Text = "";
            textBoxControlBath3Negative44_5.Text = "";
            textBoxControlBlank35.Text = "";
            textBoxControlBath1Blank44_5.Text = "";
            textBoxControlBath2Blank44_5.Text = "";
            textBoxControlBath3Blank44_5.Text = "";
            textBoxLot35.Text = "";
            textBoxLot44_5.Text = "";
            textBoxDailyDuplicatePrecisionCriteria.Text = csspWQInputApp.DailyDuplicatePrecisionCriteria.ToString();
            textBoxIntertechDuplicatePrecisionCriteria.Text = csspWQInputApp.IntertechDuplicatePrecisionCriteria.ToString();
            richTextBoxRunWeatherComment.Text = "";
            richTextBoxRunComment.Text = "";
            textBoxSampleBottleLotNumber.Text = "";
            textBoxSalinitiesReadBy.Text = "";
            textBoxResultsReadBy.Text = "";
            textBoxResultsRecordedBy.Text = "";
            lblIncubationBath1TimeCalculated.Text = "--:--";
            lblTCFirst.Text = "-.-";
            lblTCAverage.Text = "-.-";
            lblDailyDuplicateRLogValue.Text = "Not Calculated";
            lblDailyDuplicateAcceptableOrUnacceptable.Text = "Unknown";
        }
        private void CompareTheDocs()
        {
            bool Found = true;
            int ServerStartPos = 0;
            int LocalStartPos = 0;
            int ServerEndPos = 0;
            int LocalEndPos = 0;
            string ServerLine = "";
            string LocalLine = "";
            string ServerText = richTextBoxLabSheetReceiver.Text;
            string LocalText = richTextBoxLabSheetSender.Text;
            richTextBoxLabSheetReceiver.SelectAll();
            richTextBoxLabSheetReceiver.SelectionColor = Color.Black;
            richTextBoxLabSheetSender.SelectAll();
            richTextBoxLabSheetSender.SelectionColor = Color.Black;
            while (Found)
            {

                ServerEndPos = ServerText.IndexOf("\n", ServerStartPos);
                LocalEndPos = LocalText.IndexOf("\n", LocalStartPos);
                if (ServerEndPos > 0 && LocalEndPos > 0)
                {
                    richTextBoxLabSheetReceiver.Select(ServerStartPos, ServerEndPos - ServerStartPos);
                    ServerLine = richTextBoxLabSheetReceiver.SelectedText;
                    richTextBoxLabSheetSender.Select(LocalStartPos, LocalEndPos - LocalStartPos);
                    LocalLine = richTextBoxLabSheetSender.SelectedText;
                    if (ServerLine != LocalLine)
                    {
                        richTextBoxLabSheetSender.SelectionColor = Color.Red;
                    }

                    if (ServerEndPos < ServerStartPos)
                        break;

                    ServerStartPos = ServerEndPos + 1;
                    LocalStartPos = LocalEndPos + 1;
                }
                else
                {
                    break;
                }
            }
        }
        private string CreateCode(string textToCode)
        {
            List<int> intList = new List<int>();
            Random rd = new Random();
            string str = textToCode;
            foreach (char c in str)
            {
                int pos = r.IndexOf(c);
                int first = rd.Next(pos + 1, pos + 9);
                int second = rd.Next(2, 9);
                int tot = (first * second) + pos;
                intList.Add(tot);
                intList.Add(first);
            }

            StringBuilder sb = new StringBuilder();
            foreach (int i in intList)
            {
                sb.Append(i.ToString() + ",");
            }

            return sb.ToString();
        }
        private void CreateCSSPSamplingPlanFilePath()
        {
            FileInfo fiConf = new FileInfo(SamplingPlanName);
            CurrentPath = RootCurrentPath + fiConf.Name.Replace(".txt", "") + @"\";

            DirectoryInfo di = new DirectoryInfo(CurrentPath);
            if (!di.Exists)
            {
                try
                {
                    di.Create();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message + (ex.InnerException != null ? ex.InnerException.Message : ""), "Error while trying to create a directory [" + RootCurrentPath + @"] under C:\");
                    return;
                }
            }
        }
        private void CreateWordDoc()
        {
            butViewFCForm.Enabled = false;
            UpdatePanelApp();
            string FileName = lblFilePath.Text.Replace(".txt", ".docx");

            FileInfo fi = new FileInfo(lblFilePath.Text);
            StreamReader sr = fi.OpenText();
            csspFCFormWriter.FullLabSheet = sr.ReadToEnd();
            sr.Close();

            DirectoryInfo directoryInfo = fi.Directory;
            List<FileInfo> fileInfoList = directoryInfo.GetFiles().Where(c => c.FullName.EndsWith(".docx") && c.FullName.StartsWith(fi.FullName.Substring(0, fi.FullName.Length - 7))).ToList();
            foreach (FileInfo fiToDelete in fileInfoList)
            {
                try
                {
                    fiToDelete.Delete();
                }
                catch (Exception ex)
                {
                    lblStatusTxt.Text = ex.Message;
                    MessageBox.Show(ex.Message, "Old FC Form might be open. Please close FC Form.");
                    return;
                }
            }

            string retStr = csspFCFormWriter.CreateFCForm(FileName);
            if (!string.IsNullOrWhiteSpace(retStr))
            {
                lblStatus.Text = retStr;
            }
            butViewFCForm.Enabled = true;
        }
        private void DoLogCheckColorField(string WithinBars)
        {
            if (WithinBars == "Tides	")
            {
                textBoxTides.BackColor = Color.LightBlue;
            }
            if (WithinBars == "Sample Crew Initials	")
            {
                textBoxSampleCrewInitials.BackColor = Color.LightBlue;
            }
            if (WithinBars == "Incubation Start Same Day	")
            {
                checkBoxIncubationStartSameDay.BackColor = Color.LightBlue;
            }
            if (WithinBars == "Two Coolers	")
            {
                checkBox2Coolers.BackColor = Color.LightBlue;
            }
            if (WithinBars == "Incubation Bath 1 Start Time	")
            {
                textBoxIncubationBath1StartTime.BackColor = Color.LightBlue;
            }
            if (WithinBars == "Incubation Bath 2 Start Time	")
            {
                textBoxIncubationBath2StartTime.BackColor = Color.LightBlue;
            }
            if (WithinBars == "Incubation Bath 3 Start Time	")
            {
                textBoxIncubationBath3StartTime.BackColor = Color.LightBlue;
            }
            if (WithinBars == "Incubation Bath 1 End Time	")
            {
                textBoxIncubationBath1EndTime.BackColor = Color.LightBlue;
            }
            if (WithinBars == "Incubation Bath 2 End Time	")
            {
                textBoxIncubationBath2EndTime.BackColor = Color.LightBlue;
            }
            if (WithinBars == "Incubation Bath 3 End Time	")
            {
                textBoxIncubationBath3EndTime.BackColor = Color.LightBlue;
            }
            if (WithinBars == "Water Bath 1	")
            {
                textBoxWaterBath1Number.BackColor = Color.LightBlue;
            }
            if (WithinBars == "Water Bath 2	")
            {
                textBoxWaterBath2Number.BackColor = Color.LightBlue;
            }
            if (WithinBars == "Water Bath 3	")
            {
                textBoxWaterBath3Number.BackColor = Color.LightBlue;
            }
            if (WithinBars == "TC Field #1	")
            {
                textBoxTCField1.BackColor = Color.LightBlue;
            }
            if (WithinBars == "TC Lab #1	")
            {
                textBoxTCLab1.BackColor = Color.LightBlue;
            }
            if (WithinBars == "TC Field #2	")
            {
                textBoxTCField2.BackColor = Color.LightBlue;
            }
            if (WithinBars == "TC Lab #2	")
            {
                textBoxTCLab2.BackColor = Color.LightBlue;
            }
            if (WithinBars == "Control Lot	")
            {
                textBoxControlLot.BackColor = Color.LightBlue;
            }
            if (WithinBars == "Positive 35	")
            {
                textBoxControlPositive35.BackColor = Color.LightBlue;
            }
            if (WithinBars == "Non Target 35	")
            {
                textBoxControlNonTarget35.BackColor = Color.LightBlue;
            }
            if (WithinBars == "Negative 35	")
            {
                textBoxControlNegative35.BackColor = Color.LightBlue;
            }
            if (WithinBars == "Bath 1 Positive 44.5	")
            {
                textBoxControlBath1Positive44_5.BackColor = Color.LightBlue;
            }
            if (WithinBars == "Bath 2 Positive 44.5	")
            {
                textBoxControlBath2Positive44_5.BackColor = Color.LightBlue;
            }
            if (WithinBars == "Bath 3 Positive 44.5	")
            {
                textBoxControlBath3Positive44_5.BackColor = Color.LightBlue;
            }
            if (WithinBars == "Bath 1 Non Target 44.5	")
            {
                textBoxControlBath1NonTarget44_5.BackColor = Color.LightBlue;
            }
            if (WithinBars == "Bath 2 Non Target 44.5	")
            {
                textBoxControlBath2NonTarget44_5.BackColor = Color.LightBlue;
            }
            if (WithinBars == "Bath 3 Non Target 44.5	")
            {
                textBoxControlBath3NonTarget44_5.BackColor = Color.LightBlue;
            }
            if (WithinBars == "Bath 1 Negative 44.5	")
            {
                textBoxControlBath1Negative44_5.BackColor = Color.LightBlue;
            }
            if (WithinBars == "Bath 2 Negative 44.5	")
            {
                textBoxControlBath2Negative44_5.BackColor = Color.LightBlue;
            }
            if (WithinBars == "Bath 3 Negative 44.5	")
            {
                textBoxControlBath3Negative44_5.BackColor = Color.LightBlue;
            }
            if (WithinBars == "Control Blank 35	")
            {
                textBoxControlBlank35.BackColor = Color.LightBlue;
            }
            if (WithinBars == "Control Bath 1 Blank 44.5	")
            {
                textBoxControlBath1Blank44_5.BackColor = Color.LightBlue;
            }
            if (WithinBars == "Control Bath 2 Blank 44.5	")
            {
                textBoxControlBath2Blank44_5.BackColor = Color.LightBlue;
            }
            if (WithinBars == "Control Bath 3 Blank 44.5	")
            {
                textBoxControlBath3Blank44_5.BackColor = Color.LightBlue;
            }
            if (WithinBars == "Lot 35	")
            {
                textBoxLot35.BackColor = Color.LightBlue;
            }
            if (WithinBars == "Lot 44.5	")
            {
                textBoxLot44_5.BackColor = Color.LightBlue;
            }
            if (WithinBars == "Sample Bottle Lot Number	")
            {
                textBoxSampleBottleLotNumber.BackColor = Color.LightBlue;
            }
            if (WithinBars == "Salinities Read By	")
            {
                textBoxSalinitiesReadBy.BackColor = Color.LightBlue;
            }
            if (WithinBars == "Results Read By	")
            {
                textBoxResultsReadBy.BackColor = Color.LightBlue;
            }
            if (WithinBars == "Results Recorded By	")
            {
                textBoxResultsRecordedBy.BackColor = Color.LightBlue;
            }
            if (WithinBars == "Run Weather Comment	")
            {
                richTextBoxRunWeatherComment.BackColor = Color.LightBlue;
            }
            if (WithinBars == "Run Comment	")
            {
                richTextBoxRunComment.BackColor = Color.LightBlue;
            }
            if (WithinBars.StartsWith("CSSP Grid("))
            {
                if (WithinBars.Contains(")"))
                {
                    string tempText = WithinBars.Substring(0, WithinBars.IndexOf(")")).Replace("CSSP Grid(", "").Replace(")", "").Trim();
                    int Col = int.Parse(tempText.Substring(0, tempText.IndexOf(",")));
                    int Row = int.Parse(tempText.Substring(tempText.IndexOf(",") + 1));
                    try
                    {
                        DataGridViewCell dataGridViewCell = dataGridViewCSSP[Col, Row];
                        dataGridViewCell.Style.BackColor = Color.LightBlue;
                    }
                    catch (Exception)
                    {
                        // Nothing for now
                    }
                }
            }
        }
        private void DoLogClearBackgroundColor()
        {
            textBoxTides.BackColor = TextBoxBackColor;
            textBoxSampleCrewInitials.BackColor = TextBoxBackColor;
            checkBoxIncubationStartSameDay.BackColor = TextBoxBackColor;
            checkBox2Coolers.BackColor = TextBoxBackColor;
            textBoxIncubationBath1StartTime.BackColor = TextBoxBackColor;
            textBoxIncubationBath2StartTime.BackColor = TextBoxBackColor;
            textBoxIncubationBath3StartTime.BackColor = TextBoxBackColor;
            textBoxIncubationBath1EndTime.BackColor = TextBoxBackColor;
            textBoxIncubationBath2EndTime.BackColor = TextBoxBackColor;
            textBoxIncubationBath3EndTime.BackColor = TextBoxBackColor;
            textBoxWaterBath1Number.BackColor = TextBoxBackColor;
            textBoxWaterBath2Number.BackColor = TextBoxBackColor;
            textBoxWaterBath3Number.BackColor = TextBoxBackColor;
            textBoxTCField1.BackColor = TextBoxBackColor;
            textBoxTCLab1.BackColor = TextBoxBackColor;
            textBoxTCField2.BackColor = TextBoxBackColor;
            textBoxTCLab2.BackColor = TextBoxBackColor;
            textBoxControlLot.BackColor = TextBoxBackColor;
            textBoxControlPositive35.BackColor = TextBoxBackColor;
            textBoxControlNonTarget35.BackColor = TextBoxBackColor;
            textBoxControlNegative35.BackColor = TextBoxBackColor;
            textBoxControlBath1Positive44_5.BackColor = TextBoxBackColor;
            textBoxControlBath2Positive44_5.BackColor = TextBoxBackColor;
            textBoxControlBath3Positive44_5.BackColor = TextBoxBackColor;
            textBoxControlBath1NonTarget44_5.BackColor = TextBoxBackColor;
            textBoxControlBath2NonTarget44_5.BackColor = TextBoxBackColor;
            textBoxControlBath3NonTarget44_5.BackColor = TextBoxBackColor;
            textBoxControlBath1Negative44_5.BackColor = TextBoxBackColor;
            textBoxControlBath2Negative44_5.BackColor = TextBoxBackColor;
            textBoxControlBath3Negative44_5.BackColor = TextBoxBackColor;
            textBoxControlBlank35.BackColor = TextBoxBackColor;
            textBoxControlBath1Blank44_5.BackColor = TextBoxBackColor;
            textBoxControlBath2Blank44_5.BackColor = TextBoxBackColor;
            textBoxControlBath3Blank44_5.BackColor = TextBoxBackColor;
            textBoxLot35.BackColor = TextBoxBackColor;
            textBoxLot44_5.BackColor = TextBoxBackColor;
            textBoxSampleBottleLotNumber.BackColor = TextBoxBackColor;
            textBoxSalinitiesReadBy.BackColor = TextBoxBackColor;
            textBoxResultsReadBy.BackColor = TextBoxBackColor;
            textBoxResultsRecordedBy.BackColor = TextBoxBackColor;
            richTextBoxRunWeatherComment.BackColor = TextBoxBackColor;
            richTextBoxRunComment.BackColor = TextBoxBackColor;

            for (int Col = 1, countCol = dataGridViewCSSP.Columns.Count; Col < countCol; Col++)
            {
                for (int Row = 1, countRow = dataGridViewCSSP.Rows.Count; Row < countRow; Row++)
                {
                    DataGridViewCell dataGridViewCell = dataGridViewCSSP[Col, Row];
                    dataGridViewCell.Style.BackColor = TextBoxBackColor;
                }
            }
        }
        private void DoLogCheckForReediting()
        {
            DoLogClearBackgroundColor();
            LogHistory = new List<string>();
            string TheLine = "";
            StringReader strReader = new StringReader(sbLog.ToString());
            while (true)
            {
                TheLine = strReader.ReadLine();
                if (TheLine != null)
                {
                    if (TheLine.IndexOf("|||") != -1)
                    {
                        int Pos1 = TheLine.IndexOf("|||") + 3;
                        int Pos2 = TheLine.LastIndexOf("|||");
                        string WithinBars = TheLine.Substring(Pos1, Pos2 - Pos1);
                        if (LogHistory.Where(c => c == WithinBars).Any())
                        {
                            DoLogCheckColorField(WithinBars);
                        }
                        else
                        {
                            LogHistory.Add(WithinBars);
                        }
                    }
                }
                else
                {
                    break;
                }
            }
        }
        private void DoSave()
        {
            lblStatus.Text = "Saving...";
            if (lblFilePath.Text.Length > 0)
            {
                FileInfo fi = new FileInfo(lblFilePath.Text);
                if (fi.FullName.Substring(fi.FullName.Length - 4) == ".txt")
                {
                    if (fi.FullName.Substring(fi.FullName.Length - 6) != "_C.txt")
                    {
                        File.Copy(fi.FullName, fi.FullName.Replace(fi.FullName.Substring(fi.FullName.Length - 6), "_C.txt"));
                        lblFilePath.Text = lblFilePath.Text.Replace(fi.FullName.Substring(fi.FullName.Length - 6), "_C.txt");
                        fi.Delete();
                    }
                }
            }
            SaveInfoOnLocalMachine(false);
            if (InternetConnection)
            {
                if (panelAppInputIsVisible)
                {
                    butSendToServer.Text = "Send to Server";
                    butSendToServer.Enabled = true;
                    butGetLabSheetsStatus.Enabled = true;
                    //if (comboBoxSubsectorNames.SelectedIndex != 0)
                    //{
                    //    butGetLabSheetsStatus.Enabled = true;
                    //}
                }
            }
            else
            {
                butSendToServer.Text = "No Internet Connection";
                butSendToServer.Enabled = false;
                butGetLabSheetsStatus.Enabled = false;
            }
            butViewFCForm.Enabled = true;

            DoLogCheckForReediting();
        }
        private bool EverythingEntered()
        {
            StringBuilder sb = new StringBuilder();
            if (csspWQInputSheetType == CSSPWQInputSheetTypeEnum.A1)
            {
                // Tides
                if (textBoxTides.Text.Contains("--"))
                {
                    butSendToServer.Text = "Tide info required";
                    textBoxTides.BackColor = Color.Red;
                    return false;
                }

                if (csspWQInputApp.IncludeLaboratoryQAQC)
                {
                    // Initial Field Crew
                    if (string.IsNullOrWhiteSpace(textBoxSampleCrewInitials.Text))
                    {
                        butSendToServer.Text = "Sample crew initial required";
                        textBoxSampleCrewInitials.BackColor = Color.Red;
                        textBoxSampleCrewInitials.Focus();
                        return false;
                    }

                    // Incubation Start Time
                    if (string.IsNullOrWhiteSpace(textBoxIncubationBath1StartTime.Text))
                    {
                        butSendToServer.Text = "Incubation start time required";
                        textBoxIncubationBath1StartTime.BackColor = Color.Red;
                        textBoxIncubationBath1StartTime.Focus();
                        return false;
                    }

                    // Incubation End Time
                    if (string.IsNullOrWhiteSpace(textBoxIncubationBath1EndTime.Text))
                    {
                        butSendToServer.Text = "Incubation end time required";
                        textBoxIncubationBath1EndTime.BackColor = Color.Red;
                        textBoxIncubationBath1EndTime.Focus();
                        return false;
                    }

                    // Water Bath Number
                    if (string.IsNullOrWhiteSpace(textBoxWaterBath1Number.Text))
                    {
                        butSendToServer.Text = "Water bath number required";
                        textBoxWaterBath1Number.BackColor = Color.Red;
                        textBoxWaterBath1Number.Focus();
                        return false;
                    }

                    if (radioButton2Baths.Checked || radioButton3Baths.Checked)
                    {
                        // Incubation Start Time
                        if (string.IsNullOrWhiteSpace(textBoxIncubationBath2StartTime.Text))
                        {
                            butSendToServer.Text = "Incubation start time required";
                            textBoxIncubationBath2StartTime.BackColor = Color.Red;
                            textBoxIncubationBath2StartTime.Focus();
                            return false;
                        }

                        // Incubation End Time
                        if (string.IsNullOrWhiteSpace(textBoxIncubationBath2EndTime.Text))
                        {
                            butSendToServer.Text = "Incubation end time required";
                            textBoxIncubationBath2EndTime.BackColor = Color.Red;
                            textBoxIncubationBath2EndTime.Focus();
                            return false;
                        }

                        // Water Bath Number
                        if (string.IsNullOrWhiteSpace(textBoxWaterBath2Number.Text))
                        {
                            butSendToServer.Text = "Water bath number required";
                            textBoxWaterBath2Number.BackColor = Color.Red;
                            textBoxWaterBath2Number.Focus();
                            return false;
                        }
                    }

                    if (radioButton3Baths.Checked)
                    {
                        // Incubation Start Time
                        if (string.IsNullOrWhiteSpace(textBoxIncubationBath3StartTime.Text))
                        {
                            butSendToServer.Text = "Incubation start time required";
                            textBoxIncubationBath3StartTime.BackColor = Color.Red;
                            textBoxIncubationBath3StartTime.Focus();
                            return false;
                        }

                        // Incubation End Time
                        if (string.IsNullOrWhiteSpace(textBoxIncubationBath3EndTime.Text))
                        {
                            butSendToServer.Text = "Incubation end time required";
                            textBoxIncubationBath3EndTime.BackColor = Color.Red;
                            textBoxIncubationBath3EndTime.Focus();
                            return false;
                        }

                        // Water Bath Number
                        if (string.IsNullOrWhiteSpace(textBoxWaterBath3Number.Text))
                        {
                            butSendToServer.Text = "Water bath number required";
                            textBoxWaterBath3Number.BackColor = Color.Red;
                            textBoxWaterBath3Number.Focus();
                            return false;
                        }
                    }

                    // Temperature Control #1 Field
                    if (!string.IsNullOrWhiteSpace(textBoxTCField1.Text))
                    {
                        float temp = 0.0f;
                        if (!float.TryParse(textBoxTCField1.Text, out temp))
                        {
                            butSendToServer.Text = "TC Field #1 (number or empty)";
                            lblStatus.Text = "Temperature control field #1 should be a number or empty";
                            textBoxTCField1.BackColor = Color.Red;
                            textBoxTCField1.ForeColor = Color.Black;
                            return false;
                        }
                    }
                    else
                    {
                        textBoxTCField1.BackColor = Color.Red;
                    }

                    // Temperature Control #1 Lab
                    if (!string.IsNullOrWhiteSpace(textBoxTCLab1.Text))
                    {
                        float temp = 0.0f;
                        if (!float.TryParse(textBoxTCLab1.Text, out temp))
                        {
                            butSendToServer.Text = "TC Lab #1 (number or empty)";
                            lblStatus.Text = "Temperature control Lab #1 should be a number or empty";
                            textBoxTCLab1.BackColor = Color.Red;
                            textBoxTCLab1.ForeColor = Color.Black;
                            return false;
                        }
                    }
                    else
                    {
                        textBoxTCLab1.BackColor = Color.Red;
                    }

                    if (checkBox2Coolers.Checked)
                    {
                        // Temperature Control #2 Field
                        if (!string.IsNullOrWhiteSpace(textBoxTCField2.Text))
                        {
                            float temp = 0.0f;
                            if (!float.TryParse(textBoxTCField2.Text, out temp))
                            {
                                butSendToServer.Text = "TC Field #2 (number or empty)";
                                lblStatus.Text = "Temperature control field #2 should be a number or empty";
                                textBoxTCField2.BackColor = Color.Red;
                                textBoxTCField2.ForeColor = Color.Black;
                                return false;
                            }
                        }
                        else
                        {
                            textBoxTCField2.BackColor = Color.Red;
                        }

                        // Temperature Control #2 Lab
                        if (!string.IsNullOrWhiteSpace(textBoxTCLab2.Text))
                        {
                            float temp = 0.0f;
                            if (!float.TryParse(textBoxTCLab2.Text, out temp))
                            {
                                butSendToServer.Text = "TC Lab #2 (number or empty)";
                                lblStatus.Text = "Temperature control Lab #2 should be a number or empty";
                                textBoxTCLab2.BackColor = Color.Red;
                                textBoxTCLab2.ForeColor = Color.Black;
                                return false;
                            }
                        }
                        else
                        {
                            textBoxTCLab2.BackColor = Color.Red;
                        }
                    }

                    // Control Lot
                    if (string.IsNullOrWhiteSpace(textBoxControlLot.Text))
                    {
                        butSendToServer.Text = "Control lot required";
                        textBoxControlLot.BackColor = Color.Red;
                        textBoxControlLot.Focus();
                        return false;
                    }

                    // Positive35
                    if (string.IsNullOrWhiteSpace(textBoxControlPositive35.Text))
                    {
                        butSendToServer.Text = "Control Positive 35ºC required";
                        textBoxControlPositive35.BackColor = Color.Red;
                        textBoxControlPositive35.Focus();
                        return false;
                    }

                    // NonTarget 35
                    if (string.IsNullOrWhiteSpace(textBoxControlNonTarget35.Text))
                    {
                        butSendToServer.Text = "Control NonTarget 35ºC required";
                        textBoxControlNonTarget35.BackColor = Color.Red;
                        textBoxControlNonTarget35.Focus();
                        return false;
                    }

                    // Negative 35
                    if (string.IsNullOrWhiteSpace(textBoxControlNegative35.Text))
                    {
                        butSendToServer.Text = "Control Negative 35ºC required";
                        textBoxControlNegative35.BackColor = Color.Red;
                        textBoxControlNegative35.Focus();
                        return false;
                    }

                    // Positive 44.5
                    if (string.IsNullOrWhiteSpace(textBoxControlBath1Positive44_5.Text))
                    {
                        butSendToServer.Text = "Control Positive 44.5ºC required";
                        textBoxControlBath1Positive44_5.BackColor = Color.Red;
                        textBoxControlBath1Positive44_5.Focus();
                        return false;
                    }

                    // NonTarget 44.5
                    if (string.IsNullOrWhiteSpace(textBoxControlBath1NonTarget44_5.Text))
                    {
                        butSendToServer.Text = "Control NonTarget 44.5ºC required";
                        textBoxControlBath1NonTarget44_5.BackColor = Color.Red;
                        textBoxControlBath1NonTarget44_5.Focus();
                        return false;
                    }

                    // Negative 44.5
                    if (string.IsNullOrWhiteSpace(textBoxControlBath1Negative44_5.Text))
                    {
                        butSendToServer.Text = "Control Negative 44.5ºC required";
                        textBoxControlBath1Negative44_5.BackColor = Color.Red;
                        textBoxControlBath1Negative44_5.Focus();
                        return false;
                    }

                    // Blank 
                    if (string.IsNullOrWhiteSpace(textBoxControlBlank35.Text))
                    {
                        butSendToServer.Text = "Control blank 35ºC required";
                        textBoxControlBlank35.BackColor = Color.Red;
                        textBoxControlBlank35.Focus();
                        return false;
                    }

                    // Lot 35
                    if (string.IsNullOrWhiteSpace(textBoxLot35.Text))
                    {
                        butSendToServer.Text = "Media lot 1X required";
                        textBoxLot35.BackColor = Color.Red;
                        textBoxLot35.Focus();
                        return false;
                    }

                    // Lot 44.5
                    if (string.IsNullOrWhiteSpace(textBoxLot44_5.Text))
                    {
                        butSendToServer.Text = "Media lot 2X required";
                        textBoxLot44_5.BackColor = Color.Red;
                        textBoxLot44_5.Focus();
                        return false;
                    }

                    if (textBoxLot35.Text.Trim().ToUpper() == textBoxLot44_5.Text.Trim().ToUpper())
                    {
                        butSendToServer.Text = "Media lots should not be equal";
                        textBoxLot35.ForeColor = Color.Black;
                        textBoxLot44_5.ForeColor = Color.Black;
                        textBoxLot35.BackColor = Color.Red;
                        textBoxLot44_5.BackColor = Color.Red;
                        textBoxLot35.Focus();
                        return false;
                    }

                    //// Run Weather Comment
                    //if (string.IsNullOrWhiteSpace(richTextBoxRunWeatherComment.Text))
                    //{
                    //    butSendToServer.Text = "Run Weather Comment info required";
                    //    richTextBoxRunWeatherComment.BackColor = Color.Red;
                    //    richTextBoxRunWeatherComment.Focus();
                    //    return false;
                    //}

                    //// Run Comment
                    //if (string.IsNullOrWhiteSpace(richTextBoxRunComment.Text))
                    //{
                    //    butSendToServer.Text = "Run comment required";
                    //    richTextBoxRunComment.BackColor = Color.Red;
                    //    richTextBoxRunComment.Focus();
                    //    return false;
                    //}

                    // Daily Duplicate Precision Criteria
                    if (string.IsNullOrWhiteSpace(textBoxDailyDuplicatePrecisionCriteria.Text))
                    {
                        butSendToServer.Text = "Daily Duplicate precision criteria required";
                        textBoxDailyDuplicatePrecisionCriteria.BackColor = Color.Red;
                        textBoxDailyDuplicatePrecisionCriteria.Focus();
                        return false;
                    }

                    // Intertech Duplicate Precision Criteria
                    if (string.IsNullOrWhiteSpace(textBoxIntertechDuplicatePrecisionCriteria.Text))
                    {
                        butSendToServer.Text = "Intertech Duplicate precision criteria required";
                        textBoxIntertechDuplicatePrecisionCriteria.BackColor = Color.Red;
                        textBoxIntertechDuplicatePrecisionCriteria.Focus();
                        return false;
                    }

                    // Sample Bottle Lot Number
                    if (string.IsNullOrWhiteSpace(textBoxSampleBottleLotNumber.Text))
                    {
                        butSendToServer.Text = "Sample bottle lot number required";
                        textBoxSampleBottleLotNumber.BackColor = Color.Red;
                        textBoxSampleBottleLotNumber.Focus();
                        return false;
                    }

                    // Salinities read by
                    if (string.IsNullOrWhiteSpace(textBoxSalinitiesReadBy.Text))
                    {
                        butSendToServer.Text = "Salinities read by required";
                        textBoxSalinitiesReadBy.BackColor = Color.Red;
                        textBoxSalinitiesReadBy.Focus();
                        return false;
                    }

                    // Results read by
                    if (string.IsNullOrWhiteSpace(textBoxResultsReadBy.Text))
                    {
                        butSendToServer.Text = "Results read by required";
                        textBoxResultsReadBy.BackColor = Color.Red;
                        textBoxResultsReadBy.Focus();
                        return false;
                    }

                    // Results recorded by
                    if (string.IsNullOrWhiteSpace(textBoxResultsRecordedBy.Text))
                    {
                        butSendToServer.Text = "Results recorded by required";
                        textBoxResultsRecordedBy.BackColor = Color.Red;
                        textBoxResultsRecordedBy.Focus();
                        return false;
                    }
                }

                // Data Grid View 
                string ErrorMessage = "";
                List<int> EmptyRow = new List<int>();
                List<int> ColumnNumber = new List<int>() { 2, 4, 5, 6, 7, 8, 9, 12 };

                for (int row = 0, countRow = dataGridViewCSSP.Rows.Count; row < countRow; row++)
                {
                    bool AllEmpty = true;
                    for (int col = 0, countCol = dataGridViewCSSP.Columns.Count - 1; col < countCol; col++)
                    {
                        if (ColumnNumber.Contains(col))
                        {
                            if (dataGridViewCSSP[col, row].Value != null && !string.IsNullOrWhiteSpace(dataGridViewCSSP[col, row].Value.ToString()))
                            {
                                AllEmpty = false;
                            }
                        }
                    }
                    if (AllEmpty)
                    {
                        EmptyRow.Add(row);
                        continue;
                    }
                }


                if (labSheetA1Sheet.IncludeLaboratoryQAQC)
                {
                    ColumnNumber = new List<int>() { 2, 4, 5, 6, 9 };
                }
                else
                {
                    ColumnNumber = new List<int>() { 2, 4, 5, 6 };
                }

                for (int row = 0, countRow = dataGridViewCSSP.Rows.Count; row < countRow; row++)
                {
                    if (EmptyRow.Contains(row))
                    {
                        continue;
                    }

                    bool CellEmpty = false;
                    for (int col = 0, countCol = dataGridViewCSSP.Columns.Count - 1; col < countCol; col++)
                    {
                        if (ColumnNumber.Contains(col))
                        {
                            if (dataGridViewCSSP[col, row].Value == null || string.IsNullOrWhiteSpace(dataGridViewCSSP[col, row].Value.ToString()))
                            {
                                CellEmpty = true;
                                string firstCol = dataGridViewCSSP[0, row].Value.ToString();
                                ErrorMessage = firstCol.Substring(0, firstCol.IndexOf("  ")) + " --- " + dataGridViewCSSP.Columns[col].HeaderText.ToString() + "\r\n";
                            }
                        }
                    }
                    if (CellEmpty)
                    {
                        butSendToServer.Text = "Data missing in grid";
                        DialogResult dialogResult = MessageBox.Show("Please correct before sending lab sheet to server.\r\n\r\n" + ErrorMessage, "Data missing in grid.", MessageBoxButtons.OK);
                        return false;
                    }
                }

                ColumnNumber = new List<int>() { 4, 5, 6 };
                for (int row = 0, countRow = dataGridViewCSSP.Rows.Count; row < countRow; row++)
                {
                    if (EmptyRow.Contains(row))
                    {
                        continue;
                    }

                    bool InvalidTubeCombination = false;
                    for (int col = 0, countCol = dataGridViewCSSP.Columns.Count - 1; col < countCol; col++)
                    {
                        if (ColumnNumber.Contains(col))
                        {
                            if (dataGridViewCSSP[col, row].Value != null)
                            {
                                int tubeNumber = -1;
                                if (!int.TryParse(dataGridViewCSSP[col, row].Value.ToString(), out tubeNumber))
                                {
                                    string firstCol = dataGridViewCSSP[0, row].Value.ToString();
                                    ErrorMessage = firstCol.Substring(0, firstCol.IndexOf("  ")) + " --- " + dataGridViewCSSP.Columns[col].HeaderText.ToString() + "\r\n";
                                    InvalidTubeCombination = true;
                                    break;
                                }
                                else if (tubeNumber > 5 || tubeNumber < 0)
                                {
                                    string firstCol = dataGridViewCSSP[0, row].Value.ToString();
                                    ErrorMessage = firstCol.Substring(0, firstCol.IndexOf("  ")) + " --- " + dataGridViewCSSP.Columns[col].HeaderText.ToString() + "\r\n";
                                    InvalidTubeCombination = true;
                                    break;
                                }
                            }
                        }
                    }
                    if (InvalidTubeCombination)
                    {
                        butSendToServer.Text = "Invalid tube combination";
                        DialogResult dialogResult = MessageBox.Show("Please correct before sending lab sheet to server.\r\n\r\n" + ErrorMessage, "Invalid tube combination.", MessageBoxButtons.OK);
                        return false;
                    }

                }

                ColumnNumber = new List<int>() { 7, 8 };
                for (int row = 0, countRow = dataGridViewCSSP.Rows.Count; row < countRow; row++)
                {
                    if (EmptyRow.Contains(row))
                    {
                        continue;
                    }

                    bool SalOrTempIsMissing = false;
                    for (int col = 0, countCol = dataGridViewCSSP.Columns.Count - 1; col < countCol; col++)
                    {
                        if (ColumnNumber.Contains(col))
                        {
                            if (dataGridViewCSSP[col, row].Value == null || string.IsNullOrWhiteSpace(dataGridViewCSSP[col, row].Value.ToString()))
                            {
                                SalOrTempIsMissing = true;
                                string firstCol = dataGridViewCSSP[0, row].Value.ToString();
                                ErrorMessage = firstCol.Substring(0, firstCol.IndexOf("  ")) + " --- " + dataGridViewCSSP.Columns[col].HeaderText.ToString() + "\r\n";
                                break;
                            }
                        }
                    }
                    if (SalOrTempIsMissing)
                    {
                        butSendToServer.Text = "Salinity and/or temperature missing";
                        DialogResult dialogResult = MessageBox.Show(ErrorMessage + "\r\nDo you still want to send the lab sheet to server?", "Salinity and/or temperature missing.", MessageBoxButtons.YesNo);
                        if (dialogResult == DialogResult.No)
                        {
                            return false;
                        }
                    }

                }

                ColumnNumber = new List<int>() { 3 };
                for (int row = 0, countRow = dataGridViewCSSP.Rows.Count; row < countRow; row++)
                {
                    if (EmptyRow.Contains(row))
                    {
                        continue;
                    }

                    bool UnusualTubeCombination = false;
                    for (int col = 0, countCol = dataGridViewCSSP.Columns.Count - 1; col < countCol; col++)
                    {
                        if (ColumnNumber.Contains(col))
                        {
                            if (dataGridViewCSSP[col, row].Value.ToString() == "Error")
                            {
                                UnusualTubeCombination = true;
                                string firstCol = dataGridViewCSSP[0, row].Value.ToString();
                                ErrorMessage = firstCol.Substring(0, firstCol.IndexOf("  ")) + "\r\n";
                                break;
                            }
                        }
                    }
                    if (UnusualTubeCombination)
                    {
                        butSendToServer.Text = "Unusual tube combination";
                        DialogResult dialogResult = MessageBox.Show(ErrorMessage + "\r\nDo you still want to send the lab sheet to server?", "Unusual tube combination.", MessageBoxButtons.YesNo);
                        if (dialogResult == DialogResult.No)
                        {
                            return false;
                        }
                    }
                }
            }

            if (sb.ToString().Length > 0)
            {
                sb.AppendLine("");
                sb.AppendLine("Do you want to continue sending the lab sheet to the server?");
                if (MessageBox.Show(sb.ToString(), "Warning", MessageBoxButtons.YesNo) == DialogResult.No)
                {
                    return false;
                }
            }
            return true;
        }
        private void FillFormWithParsedLabSheetA1(LabSheetA1Sheet labSheetA1Sheet)
        {
            if (labSheetA1Sheet.WaterBathCount == 3)
            {
                radioButton3Baths.Checked = true;
            }
            else if (labSheetA1Sheet.WaterBathCount == 2)
            {
                radioButton2Baths.Checked = true;
            }
            else
            {
                radioButton1Baths.Checked = true;
            }
            textBoxTides.Text = labSheetA1Sheet.Tides;
            textBoxSampleCrewInitials.Text = labSheetA1Sheet.SampleCrewInitials;
            textBoxIncubationBath1StartTime.Text = labSheetA1Sheet.IncubationBath1StartTime;
            textBoxIncubationBath2StartTime.Text = labSheetA1Sheet.IncubationBath2StartTime;
            textBoxIncubationBath3StartTime.Text = labSheetA1Sheet.IncubationBath3StartTime;
            textBoxIncubationBath1EndTime.Text = labSheetA1Sheet.IncubationBath1EndTime;
            textBoxIncubationBath2EndTime.Text = labSheetA1Sheet.IncubationBath2EndTime;
            textBoxIncubationBath3EndTime.Text = labSheetA1Sheet.IncubationBath3EndTime;
            lblIncubationBath1TimeCalculated.Text = labSheetA1Sheet.IncubationBath1TimeCalculated;
            lblIncubationBath2TimeCalculated.Text = labSheetA1Sheet.IncubationBath2TimeCalculated;
            lblIncubationBath3TimeCalculated.Text = labSheetA1Sheet.IncubationBath3TimeCalculated;
            textBoxWaterBath1Number.Text = labSheetA1Sheet.WaterBath1;
            textBoxWaterBath2Number.Text = labSheetA1Sheet.WaterBath2;
            textBoxWaterBath3Number.Text = labSheetA1Sheet.WaterBath3;
            checkBox2Coolers.Checked = (labSheetA1Sheet.TCHas2Coolers == "false" ? false : true);
            textBoxTCField1.Text = labSheetA1Sheet.TCField1;
            textBoxTCLab1.Text = labSheetA1Sheet.TCLab1;
            textBoxTCField2.Text = labSheetA1Sheet.TCField2;
            textBoxTCLab2.Text = labSheetA1Sheet.TCLab2;
            lblTCFirst.Text = labSheetA1Sheet.TCFirst;
            lblTCAverage.Text = labSheetA1Sheet.TCAverage;
            textBoxControlLot.Text = labSheetA1Sheet.ControlLot;
            textBoxControlPositive35.Text = labSheetA1Sheet.Positive35;
            textBoxControlNonTarget35.Text = labSheetA1Sheet.NonTarget35;
            textBoxControlNegative35.Text = labSheetA1Sheet.Negative35;
            textBoxControlBath1Positive44_5.Text = labSheetA1Sheet.Bath1Positive44_5;
            textBoxControlBath2Positive44_5.Text = labSheetA1Sheet.Bath2Positive44_5;
            textBoxControlBath3Positive44_5.Text = labSheetA1Sheet.Bath3Positive44_5;
            textBoxControlBath1NonTarget44_5.Text = labSheetA1Sheet.Bath1NonTarget44_5;
            textBoxControlBath2NonTarget44_5.Text = labSheetA1Sheet.Bath2NonTarget44_5;
            textBoxControlBath3NonTarget44_5.Text = labSheetA1Sheet.Bath3NonTarget44_5;
            textBoxControlBath1Negative44_5.Text = labSheetA1Sheet.Bath1Negative44_5;
            textBoxControlBath2Negative44_5.Text = labSheetA1Sheet.Bath2Negative44_5;
            textBoxControlBath3Negative44_5.Text = labSheetA1Sheet.Bath3Negative44_5;
            textBoxControlBlank35.Text = labSheetA1Sheet.Blank35;
            textBoxControlBath1Blank44_5.Text = labSheetA1Sheet.Bath1Blank44_5;
            textBoxControlBath2Blank44_5.Text = labSheetA1Sheet.Bath2Blank44_5;
            textBoxControlBath3Blank44_5.Text = labSheetA1Sheet.Bath3Blank44_5;
            textBoxLot35.Text = labSheetA1Sheet.Lot35;
            textBoxLot44_5.Text = labSheetA1Sheet.Lot44_5;
            richTextBoxRunWeatherComment.Text = labSheetA1Sheet.RunWeatherComment;
            richTextBoxRunComment.Text = labSheetA1Sheet.RunComment;
            textBoxSampleBottleLotNumber.Text = labSheetA1Sheet.SampleBottleLotNumber;
            textBoxSalinitiesReadBy.Text = labSheetA1Sheet.SalinitiesReadBy;
            if (labSheetA1Sheet.SalinitiesReadYear == null)
            {
                dateTimePickerSalinitiesReadDate.Value = DateTime.Now;
            }
            else
            {
                dateTimePickerSalinitiesReadDate.Value = new DateTime(int.Parse(labSheetA1Sheet.SalinitiesReadYear), int.Parse(labSheetA1Sheet.SalinitiesReadMonth), int.Parse(labSheetA1Sheet.SalinitiesReadDay));
            }
            textBoxResultsReadBy.Text = labSheetA1Sheet.ResultsReadBy;
            if (labSheetA1Sheet.ResultsReadYear == null)
            {
                dateTimePickerResultsReadDate.Value = DateTime.Now.AddDays(1);
            }
            else
            {
                dateTimePickerResultsReadDate.Value = new DateTime(int.Parse(labSheetA1Sheet.ResultsReadYear), int.Parse(labSheetA1Sheet.ResultsReadMonth), int.Parse(labSheetA1Sheet.ResultsReadDay));
            }
            textBoxResultsRecordedBy.Text = labSheetA1Sheet.ResultsRecordedBy;
            if (labSheetA1Sheet.ResultsRecordedYear == null)
            {
                dateTimePickerResultsRecordedDate.Value = DateTime.Now.AddDays(1);
            }
            else
            {
                dateTimePickerResultsRecordedDate.Value = new DateTime(int.Parse(labSheetA1Sheet.ResultsRecordedYear), int.Parse(labSheetA1Sheet.ResultsRecordedMonth), int.Parse(labSheetA1Sheet.ResultsRecordedDay));
            }
            textBoxDailyDuplicatePrecisionCriteria.Text = labSheetA1Sheet.DailyDuplicatePrecisionCriteria;
            lblDailyDuplicateAcceptableOrUnacceptable.Text = labSheetA1Sheet.DailyDuplicateAcceptableOrUnacceptable;

            if (dateTimePickerSalinitiesReadDate.Value == dateTimePickerRun.Value)
            {
                butSalinitySameDay.Text = "Next Day";
            }
            else
            {
                butSalinitySameDay.Text = "Same Day";
            }

            textBoxIntertechDuplicatePrecisionCriteria.Text = labSheetA1Sheet.IntertechDuplicatePrecisionCriteria;
        }
        private string FillInternetConnectionVariable()
        {
            try
            {
                using (WebClient webClient = new WebClient())
                {
                    WebProxy webProxy = new WebProxy();
                    webClient.Proxy = webProxy;
                    string s = webClient.DownloadString(new Uri("http://www.google.com"));
                    this.Text = FormTitle + " (Internet connection)";
                    InternetConnection = true;
                    if (panelAppInputIsVisible)
                    {
                        butSendToServer.Enabled = true;
                        if (comboBoxSubsectorNames.SelectedIndex != 0)
                        {
                            butGetLabSheetsStatus.Enabled = true;
                        }
                    }
                }
            }
            catch (Exception)
            {
                this.Text = FormTitle + " (No internet connection)";
                InternetConnection = false;
                butSendToServer.Enabled = false;
                butGetLabSheetsStatus.Enabled = false;
                butGetTides.Enabled = false;
                return "Error";
            }

            return "";
        }
        private void FillComboboxes()
        {
            comboBoxSubsectorNames.Items.Clear();
            comboBoxSubsectorNames.Items.Add(new FileItem("Subsector", 0));
            foreach (CSSPWQInputParam csspWQInputParam in csspWQInputParamList.Where(c => c.CSSPWQInputType == csspWQInputTypeCurrent).OrderBy(c => c.Name))
            {
                comboBoxSubsectorNames.Items.Add(new FileItem(csspWQInputParam.Name, csspWQInputParam.TVItemID));
            }
            comboBoxSubsectorNames.DisplayMember = "Name";
            comboBoxSubsectorNames.ValueMember = "TVItemID";

            if (comboBoxSubsectorNames.Items.Count > 0)
            {
                comboBoxSubsectorNames.SelectedIndex = 0;
            }

            comboBoxRunNumber.Items.Clear();
            for (int i = 1; i < 21; i++)
            {
                comboBoxRunNumber.Items.Add((i < 10 ? "0" : "") + i.ToString());
            }

            if (comboBoxRunNumber.Items.Count > 0)
            {
                comboBoxRunNumber.SelectedIndex = 0;
            }
            SetupAppInputFiles();

        }
        private void FillCSSPMPNTable()
        {
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 0, Tube1_0 = 0, Tube0_1 = 0, MPN = 1 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 0, Tube1_0 = 0, Tube0_1 = 1, MPN = 2 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 0, Tube1_0 = 1, Tube0_1 = 0, MPN = 2 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 0, Tube1_0 = 1, Tube0_1 = 1, MPN = 4 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 0, Tube1_0 = 2, Tube0_1 = 0, MPN = 4 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 0, Tube1_0 = 2, Tube0_1 = 1, MPN = 6 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 0, Tube1_0 = 3, Tube0_1 = 0, MPN = 6 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 1, Tube1_0 = 0, Tube0_1 = 0, MPN = 2 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 1, Tube1_0 = 0, Tube0_1 = 1, MPN = 4 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 1, Tube1_0 = 0, Tube0_1 = 2, MPN = 6 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 1, Tube1_0 = 1, Tube0_1 = 0, MPN = 4 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 1, Tube1_0 = 1, Tube0_1 = 1, MPN = 6 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 1, Tube1_0 = 1, Tube0_1 = 2, MPN = 8 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 1, Tube1_0 = 2, Tube0_1 = 0, MPN = 6 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 1, Tube1_0 = 2, Tube0_1 = 1, MPN = 8 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 1, Tube1_0 = 3, Tube0_1 = 0, MPN = 8 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 1, Tube1_0 = 3, Tube0_1 = 1, MPN = 10 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 1, Tube1_0 = 4, Tube0_1 = 0, MPN = 10 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 2, Tube1_0 = 0, Tube0_1 = 0, MPN = 4 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 2, Tube1_0 = 0, Tube0_1 = 1, MPN = 7 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 2, Tube1_0 = 0, Tube0_1 = 2, MPN = 9 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 2, Tube1_0 = 1, Tube0_1 = 0, MPN = 7 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 2, Tube1_0 = 1, Tube0_1 = 1, MPN = 9 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 2, Tube1_0 = 1, Tube0_1 = 2, MPN = 12 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 2, Tube1_0 = 2, Tube0_1 = 0, MPN = 9 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 2, Tube1_0 = 2, Tube0_1 = 1, MPN = 12 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 2, Tube1_0 = 2, Tube0_1 = 2, MPN = 14 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 2, Tube1_0 = 3, Tube0_1 = 0, MPN = 12 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 2, Tube1_0 = 3, Tube0_1 = 1, MPN = 14 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 2, Tube1_0 = 4, Tube0_1 = 0, MPN = 15 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 3, Tube1_0 = 0, Tube0_1 = 0, MPN = 8 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 3, Tube1_0 = 0, Tube0_1 = 1, MPN = 11 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 3, Tube1_0 = 0, Tube0_1 = 2, MPN = 13 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 3, Tube1_0 = 1, Tube0_1 = 0, MPN = 11 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 3, Tube1_0 = 1, Tube0_1 = 1, MPN = 14 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 3, Tube1_0 = 1, Tube0_1 = 2, MPN = 17 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 3, Tube1_0 = 2, Tube0_1 = 0, MPN = 14 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 3, Tube1_0 = 2, Tube0_1 = 1, MPN = 17 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 3, Tube1_0 = 2, Tube0_1 = 2, MPN = 20 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 3, Tube1_0 = 3, Tube0_1 = 0, MPN = 17 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 3, Tube1_0 = 3, Tube0_1 = 1, MPN = 21 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 3, Tube1_0 = 3, Tube0_1 = 2, MPN = 24 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 3, Tube1_0 = 4, Tube0_1 = 0, MPN = 21 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 3, Tube1_0 = 4, Tube0_1 = 1, MPN = 24 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 3, Tube1_0 = 5, Tube0_1 = 0, MPN = 25 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 4, Tube1_0 = 0, Tube0_1 = 0, MPN = 13 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 4, Tube1_0 = 0, Tube0_1 = 1, MPN = 17 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 4, Tube1_0 = 0, Tube0_1 = 2, MPN = 21 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 4, Tube1_0 = 0, Tube0_1 = 3, MPN = 25 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 4, Tube1_0 = 1, Tube0_1 = 0, MPN = 17 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 4, Tube1_0 = 1, Tube0_1 = 1, MPN = 21 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 4, Tube1_0 = 1, Tube0_1 = 2, MPN = 26 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 4, Tube1_0 = 1, Tube0_1 = 3, MPN = 31 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 4, Tube1_0 = 2, Tube0_1 = 0, MPN = 22 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 4, Tube1_0 = 2, Tube0_1 = 1, MPN = 26 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 4, Tube1_0 = 2, Tube0_1 = 2, MPN = 32 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 4, Tube1_0 = 2, Tube0_1 = 3, MPN = 38 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 4, Tube1_0 = 3, Tube0_1 = 0, MPN = 27 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 4, Tube1_0 = 3, Tube0_1 = 1, MPN = 33 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 4, Tube1_0 = 3, Tube0_1 = 2, MPN = 39 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 4, Tube1_0 = 4, Tube0_1 = 0, MPN = 34 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 4, Tube1_0 = 4, Tube0_1 = 1, MPN = 40 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 4, Tube1_0 = 4, Tube0_1 = 2, MPN = 47 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 4, Tube1_0 = 5, Tube0_1 = 0, MPN = 41 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 4, Tube1_0 = 5, Tube0_1 = 1, MPN = 48 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 5, Tube1_0 = 0, Tube0_1 = 0, MPN = 23 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 5, Tube1_0 = 0, Tube0_1 = 1, MPN = 31 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 5, Tube1_0 = 0, Tube0_1 = 2, MPN = 43 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 5, Tube1_0 = 0, Tube0_1 = 3, MPN = 58 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 5, Tube1_0 = 1, Tube0_1 = 0, MPN = 33 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 5, Tube1_0 = 1, Tube0_1 = 1, MPN = 46 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 5, Tube1_0 = 1, Tube0_1 = 2, MPN = 63 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 5, Tube1_0 = 1, Tube0_1 = 3, MPN = 84 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 5, Tube1_0 = 2, Tube0_1 = 0, MPN = 49 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 5, Tube1_0 = 2, Tube0_1 = 1, MPN = 70 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 5, Tube1_0 = 2, Tube0_1 = 2, MPN = 94 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 5, Tube1_0 = 2, Tube0_1 = 3, MPN = 120 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 5, Tube1_0 = 2, Tube0_1 = 4, MPN = 150 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 5, Tube1_0 = 3, Tube0_1 = 0, MPN = 79 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 5, Tube1_0 = 3, Tube0_1 = 1, MPN = 110 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 5, Tube1_0 = 3, Tube0_1 = 2, MPN = 140 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 5, Tube1_0 = 3, Tube0_1 = 3, MPN = 170 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 5, Tube1_0 = 3, Tube0_1 = 4, MPN = 210 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 5, Tube1_0 = 4, Tube0_1 = 0, MPN = 130 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 5, Tube1_0 = 4, Tube0_1 = 1, MPN = 170 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 5, Tube1_0 = 4, Tube0_1 = 2, MPN = 220 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 5, Tube1_0 = 4, Tube0_1 = 3, MPN = 280 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 5, Tube1_0 = 4, Tube0_1 = 4, MPN = 350 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 5, Tube1_0 = 4, Tube0_1 = 5, MPN = 430 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 5, Tube1_0 = 5, Tube0_1 = 0, MPN = 240 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 5, Tube1_0 = 5, Tube0_1 = 1, MPN = 350 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 5, Tube1_0 = 5, Tube0_1 = 2, MPN = 540 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 5, Tube1_0 = 5, Tube0_1 = 3, MPN = 920 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 5, Tube1_0 = 5, Tube0_1 = 4, MPN = 1600 });
            csspMPNTableList.Add(new CSSPMPNTable() { Tube10 = 5, Tube1_0 = 5, Tube0_1 = 5, MPN = 1700 });
        }
        private string GetAfterSampleTypeSpace(string SampleTypeText)
        {
            string AfterSampleTypeSpace = "";
            switch (SampleTypeText)
            {
                case "Routine":
                    AfterSampleTypeSpace = "                             ";
                    break;
                case "Infrastructure":
                    AfterSampleTypeSpace = "                    ";
                    break;
                case "RainCMPRoutine":
                    AfterSampleTypeSpace = "                  ";
                    break;
                case "RainRun":
                    AfterSampleTypeSpace = "                             ";
                    break;
                case "ReopeningEmergencyRain":
                    AfterSampleTypeSpace = " ";
                    break;
                case "ReopeningSpill":
                    AfterSampleTypeSpace = "                     ";
                    break;
                case "Sanitary":
                    AfterSampleTypeSpace = "                           ";
                    break;
                case "Study":
                    AfterSampleTypeSpace = "                                 ";
                    break;
                case "DailyDuplicate":
                    AfterSampleTypeSpace = "                     ";
                    break;
                case "IntertechDuplicate":
                    AfterSampleTypeSpace = "              ";
                    break;
                case "IntertechRead":
                    AfterSampleTypeSpace = "                     ";
                    break;
                default:
                    break;
            }

            return AfterSampleTypeSpace;
        }
        private AcceptedOrRejected GetAcceptedOrRejected(FileInfo fi)
        {
            AcceptedOrRejected acceptedOrRejected = new AcceptedOrRejected() { AcceptedOrRejectedBy = "", AcceptedOrRejectedDate = new DateTime(), RejectReason = "" };

            if (fi.FullName.EndsWith("_A.txt") && InternetConnection)
            {
                try
                {

                    using (WebClient webClient = new WebClient())
                    {
                        WebProxy webProxy = new WebProxy();
                        webClient.Proxy = webProxy;

                        NameValueCollection paramList = new NameValueCollection();

                        if (!fi.Exists)
                        {
                            lblStatus.Text = "File [" + fi.FullName + "] does not exist";
                            return acceptedOrRejected;
                        }

                        paramList.Add("SamplingPlanName", lblSamplingPlanFileName.Text);
                        string Rest = fi.FullName.Replace(CurrentPath, "");
                        int Pos = Rest.IndexOf("_");
                        string Subsector = Rest.Substring(0, Pos);
                        int Year = 0;
                        int.TryParse(Rest.Substring(Pos + 1, 4), out Year);
                        if (Year == 0)
                        {
                            lblStatus.Text = "Year not found";
                            return acceptedOrRejected;
                        }
                        int Month = 0;
                        int.TryParse(Rest.Substring(Pos + 6, 2), out Month);
                        if (Month == 0)
                        {
                            lblStatus.Text = "Month not found";
                            return acceptedOrRejected;
                        }
                        int Day = 0;
                        int.TryParse(Rest.Substring(Pos + 9, 2), out Day);
                        if (Day == 0)
                        {
                            lblStatus.Text = "Day not found";
                            return acceptedOrRejected;
                        }

                        DateTime dateTimeRun = new DateTime(Year, Month, Day);

                        paramList.Add("Year", Year.ToString());
                        paramList.Add("Month", Month.ToString());
                        paramList.Add("Day", Day.ToString());
                        if (comboBoxSubsectorNames.SelectedItem == null)
                        {
                            lblStatus.Text = "Subsector not selected";
                            return acceptedOrRejected;
                        }
                        FileItem item = (FileItem)comboBoxSubsectorNames.SelectedItem;
                        if (item.TVItemID == 0)
                        {
                            lblStatus.Text = "Subsector not selected";
                            return acceptedOrRejected;
                        }
                        paramList.Add("SubsectorTVItemID", item.TVItemID.ToString());
                        paramList.Add("SamplingPlanType", ((int)labSheetA1Sheet.SamplingPlanType).ToString());
                        paramList.Add("SampleType", ((int)labSheetA1Sheet.SampleType).ToString());
                        paramList.Add("LabSheetType", ((int)labSheetA1Sheet.LabSheetType).ToString());

                        webClient.Headers.Add("Content-Type", "application/x-www-form-urlencoded");
                        byte[] ret = webClient.UploadValues(new Uri("http://cssplabsheet.azurewebsites.net/GetLabSheetAcceptedOrRejectedBy.aspx"), "POST", paramList);
                        //byte[] ret = webClient.UploadValues(new Uri("http://localhost:7668/GetLabSheetAcceptedOrRejectedBy.aspx"), "POST", paramList);

                        if (ret.Length > 0)
                        {
                            string FullText = System.Text.Encoding.Default.GetString(ret);
                            List<string> ParamTextList = FullText.Split("\r\n".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();
                            if (ParamTextList.Count != 7)
                            {
                                lblStatus.Text = "Information returned from server not well formed";
                                return acceptedOrRejected;
                            }

                            List<string> ShouldContain = new List<string> { "OtherServerLabSheetID", "ApprovedOrRejectedBy", "Year", "Month", "Day", "Hour", "Minute" };

                            foreach (string s in ShouldContain)
                            {
                                if (!FullText.Contains(s))
                                {
                                    lblStatus.Text = "Information returned from the server is missing " + s;
                                    return acceptedOrRejected;
                                }
                            }

                            int OtherServerLabSheetID = 0;
                            string ApprovedOrRejectedBy = "";
                            int ServerYear = 0;
                            int ServerMonth = 0;
                            int ServerDay = 0;
                            int ServerHour = 0;
                            int ServerMinute = 0;
                            string RejectReason = "";

                            foreach (string s in ParamTextList)
                            {
                                List<string> ParamVal = s.Split("|||||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();
                                if (ParamVal.Count != 2)
                                {
                                    lblStatus.Text = ParamVal + " was not parsed properly";
                                    return acceptedOrRejected;
                                }

                                switch (ParamVal[0])
                                {
                                    case "OtherServerLabSheetID":
                                        {
                                            string Temp = ParamVal[1].Replace("[", "");
                                            Temp = Temp.Replace("]", "");
                                            if (!int.TryParse(Temp, out OtherServerLabSheetID))
                                            {
                                                lblStatus.Text = ParamVal[1] + " OtherServerLabSheetID was not parsed properly";
                                                return acceptedOrRejected;
                                            }

                                            if (OtherServerLabSheetID == 0)
                                            {
                                                lblStatus.Text = "OtherServerLabSheetID should not equal 0";
                                                return acceptedOrRejected;
                                            }
                                        }
                                        break;
                                    case "AcceptedOrRejectedBy":
                                        {
                                            string Temp = ParamVal[1].Replace("[", "");
                                            Temp = Temp.Replace("]", "");
                                            if (string.IsNullOrWhiteSpace(Temp))
                                            {
                                                lblStatus.Text = "AcceptedOrRejectedBy should not equal empty";
                                                return acceptedOrRejected;
                                            }
                                            ApprovedOrRejectedBy = Temp.Trim();
                                        }
                                        break;
                                    case "Year":
                                        {
                                            string Temp = ParamVal[1].Replace("[", "");
                                            Temp = Temp.Replace("]", "");
                                            if (!int.TryParse(Temp, out ServerYear))
                                            {
                                                lblStatus.Text = ParamVal[1] + " Year was not parsed properly";
                                                return acceptedOrRejected;
                                            }

                                            if (ServerYear == 0)
                                            {
                                                lblStatus.Text = "Year should not equal 0";
                                                return acceptedOrRejected;
                                            }
                                        }
                                        break;
                                    case "Month":
                                        {
                                            string Temp = ParamVal[1].Replace("[", "");
                                            Temp = Temp.Replace("]", "");
                                            if (!int.TryParse(Temp, out ServerMonth))
                                            {
                                                lblStatus.Text = ParamVal[1] + " Month was not parsed properly";
                                                return acceptedOrRejected;
                                            }

                                            if (ServerMonth == 0)
                                            {
                                                lblStatus.Text = "Month should not equal 0";
                                                return acceptedOrRejected;
                                            }
                                        }
                                        break;
                                    case "Day":
                                        {
                                            string Temp = ParamVal[1].Replace("[", "");
                                            Temp = Temp.Replace("]", "");
                                            if (!int.TryParse(Temp, out ServerDay))
                                            {
                                                lblStatus.Text = ParamVal[1] + " Day was not parsed properly";
                                                return acceptedOrRejected;
                                            }

                                            if (ServerDay == 0)
                                            {
                                                lblStatus.Text = "Day should not equal 0";
                                                return acceptedOrRejected;
                                            }
                                        }
                                        break;
                                    case "Hour":
                                        {
                                            string Temp = ParamVal[1].Replace("[", "");
                                            Temp = Temp.Replace("]", "");
                                            if (!int.TryParse(Temp, out ServerHour))
                                            {
                                                lblStatus.Text = ParamVal[1] + " Hour was not parsed properly";
                                                return acceptedOrRejected;
                                            }
                                        }
                                        break;
                                    case "Minute":
                                        {
                                            string Temp = ParamVal[1].Replace("[", "");
                                            Temp = Temp.Replace("]", "");
                                            if (!int.TryParse(Temp, out ServerMinute))
                                            {
                                                lblStatus.Text = ParamVal[1] + " ServerMinute was not parsed properly";
                                                return acceptedOrRejected;
                                            }
                                        }
                                        break;
                                    case "RejectReason":
                                        {
                                            string Temp = ParamVal[1].Replace("[", "");
                                            Temp = Temp.Replace("]", "");
                                            if (string.IsNullOrWhiteSpace(Temp))
                                            {
                                                lblStatus.Text = "RejectReason should not equal empty";
                                                return acceptedOrRejected;
                                            }
                                            RejectReason = Temp.Trim();
                                        }
                                        break;
                                    default:
                                        break;
                                }
                            }

                            acceptedOrRejected.AcceptedOrRejectedBy = ApprovedOrRejectedBy;
                            acceptedOrRejected.AcceptedOrRejectedDate = new DateTime(ServerYear, ServerMonth, ServerDay, ServerHour, ServerMinute, 0);
                            acceptedOrRejected.RejectReason = RejectReason;

                            return acceptedOrRejected;
                        }
                        else
                        {
                            return acceptedOrRejected;
                        }
                    }

                }
                catch (Exception ex)
                {
                    lblStatus.Text = ex.Message + (ex.InnerException == null ? "" : ex.InnerException.Message);
                    return acceptedOrRejected;
                }
            }

            return acceptedOrRejected;

        }
        private string GetCodeString(string code)
        {
            string retStr = "";
            List<int> intList = new List<int>();
            List<string> strList = code.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();

            for (int i = 0, count = strList.Count(); i < count; i = i + 2)
            {
                retStr = retStr + r.Substring((int.Parse(strList[i]) % int.Parse(strList[i + 1])), 1);
            }

            return retStr;
        }
        private string GetLabSheet(FileInfo fi)
        {
            string retStr = FillInternetConnectionVariable();
            if (!string.IsNullOrWhiteSpace(retStr))
            {
                return retStr;
            }

            try
            {

                using (WebClient webClient = new WebClient())
                {
                    WebProxy webProxy = new WebProxy();
                    webClient.Proxy = webProxy;

                    NameValueCollection paramList = new NameValueCollection();

                    if (!fi.Exists)
                    {
                        return "File [" + fi.FullName + "] does not exist";
                    }

                    paramList.Add("SamplingPlanName", lblSamplingPlanFileName.Text);
                    string Rest = fi.FullName.Replace(CurrentPath, "");
                    int Pos = Rest.IndexOf("_");
                    string Subsector = Rest.Substring(0, Pos);
                    int Year = 0;
                    int.TryParse(Rest.Substring(Pos + 1, 4), out Year);
                    if (Year == 0)
                    {
                        return "Year not found";
                    }
                    int Month = 0;
                    int.TryParse(Rest.Substring(Pos + 6, 2), out Month);
                    if (Month == 0)
                    {
                        return "Month not found";
                    }
                    int Day = 0;
                    int.TryParse(Rest.Substring(Pos + 9, 2), out Day);
                    if (Day == 0)
                    {
                        return "Day not found";
                    }

                    DateTime dateTimeRun = new DateTime(Year, Month, Day);

                    paramList.Add("Year", Year.ToString());
                    paramList.Add("Month", Month.ToString());
                    paramList.Add("Day", Day.ToString());
                    paramList.Add("RunNumber", labSheetA1Sheet.RunNumber.ToString());
                    if (comboBoxSubsectorNames.SelectedItem == null)
                    {
                        return "Subsector not selected";
                    }
                    FileItem item = (FileItem)comboBoxSubsectorNames.SelectedItem;
                    if (item.TVItemID == 0)
                    {
                        return "Subsector not selected";
                    }
                    paramList.Add("SubsectorTVItemID", item.TVItemID.ToString());
                    paramList.Add("SamplingPlanType", ((int)labSheetA1Sheet.SamplingPlanType).ToString());
                    paramList.Add("SampleType", ((int)labSheetA1Sheet.SampleType).ToString());
                    paramList.Add("LabSheetType", ((int)labSheetA1Sheet.LabSheetType).ToString());

                    webClient.Headers.Add("Content-Type", "application/x-www-form-urlencoded");
                    byte[] ret = webClient.UploadValues(new Uri("http://cssplabsheet.azurewebsites.net/GetLabSheet.aspx"), "POST", paramList);
                    //byte[] ret = webClient.UploadValues(new Uri("http://localhost:7668/GetLabSheet.aspx"), "POST", paramList);

                    if (ret.Length > 0)
                    {
                        return System.Text.Encoding.Default.GetString(ret);
                    }
                }

            }
            catch (Exception ex)
            {
                return ex.Message + (ex.InnerException == null ? "" : ex.InnerException.Message);
            }

            return "";
        }
        private string GetLabSheetExist(FileInfo fi)
        {
            string retStr = FillInternetConnectionVariable();
            if (!string.IsNullOrWhiteSpace(retStr))
            {
                return retStr;
            }

            try
            {

                using (WebClient webClient = new WebClient())
                {
                    WebProxy webProxy = new WebProxy();
                    webClient.Proxy = webProxy;

                    NameValueCollection paramList = new NameValueCollection();

                    if (!fi.Exists)
                    {
                        return "File [" + fi.FullName + "] does not exist";
                    }

                    paramList.Add("SamplingPlanName", lblSamplingPlanFileName.Text);
                    // "C:\\CSSPLabSheets\\SamplingPlan_Subsector_Routine_A1_2017_aaa.txt"
                    string Rest = lblSamplingPlanFileName.Text;
                    //Rest = Rest.Substring(Rest.IndexOf("SamplingPlan"));
                    Rest = Rest.Substring(Rest.IndexOf("_") + 1);
                    string SamplingPlanTypeText = Rest.Substring(0, Rest.IndexOf("_"));
                    Rest = Rest.Substring(SamplingPlanTypeText.Length + 1);
                    SamplingPlanTypeEnum SamplingPlanType = SamplingPlanTypeEnum.Error;
                    for (int i = 0, count = Enum.GetNames(typeof(SamplingPlanTypeEnum)).Count(); i < count; i++)
                    {
                        if (((SamplingPlanTypeEnum)i).ToString() == SamplingPlanTypeText)
                        {
                            SamplingPlanType = (SamplingPlanTypeEnum)i;
                        }
                    }
                    string SampleTypeText = Rest.Substring(0, Rest.IndexOf("_"));
                    Rest = Rest.Substring(SampleTypeText.Length + 1);
                    SampleTypeEnum SampleType = SampleTypeEnum.Error;
                    for (int i = 101, count = Enum.GetNames(typeof(SampleTypeEnum)).Count() + 101; i < count; i++)
                    {
                        if (((SampleTypeEnum)i).ToString() == SampleTypeText)
                        {
                            SampleType = (SampleTypeEnum)i;
                        }
                    }
                    string LabSheetTypeText = Rest.Substring(0, Rest.IndexOf("_"));
                    Rest = Rest.Substring(LabSheetTypeText.Length + 1);
                    LabSheetTypeEnum LabSheetType = LabSheetTypeEnum.Error;
                    for (int i = 0, count = Enum.GetNames(typeof(LabSheetTypeEnum)).Count(); i < count; i++)
                    {
                        if (((LabSheetTypeEnum)i).ToString() == LabSheetTypeText)
                        {
                            LabSheetType = (LabSheetTypeEnum)i;
                        }
                    }

                    Rest = fi.FullName.Replace(CurrentPath, "");
                    int Pos = Rest.IndexOf(@"\");
                    if (Pos > 0)
                    {
                        Rest = Rest.Substring(Pos + 1);
                    }
                    string Subsector = Rest.Substring(0, Rest.IndexOf("_"));
                    Rest = Rest.Substring(Subsector.Length + 1);
                    int Year = 0;
                    int.TryParse(Rest.Substring(0, 4), out Year);
                    if (Year == 0)
                    {
                        return "Year not found";
                    }
                    Rest = Rest.Substring(5);
                    int Month = 0;
                    int.TryParse(Rest.Substring(0, 2), out Month);
                    if (Month == 0)
                    {
                        return "Month not found";
                    }
                    Rest = Rest.Substring(3);
                    int Day = 0;
                    int.TryParse(Rest.Substring(0, 2), out Day);
                    if (Day == 0)
                    {
                        return "Day not found";
                    }
                    Rest = Rest.Substring(3);
                    string LabSheetTypeNotUsed = Rest.Substring(0, 2);
                    Rest = Rest.Substring(3);
                    int RunNumber = 0;
                    int.TryParse(Rest.Substring(1, 2), out RunNumber);
                    if (RunNumber == 0)
                    {
                        return "Day not found";
                    }
                    Rest = Rest.Substring(4);


                    DateTime dateTimeRun = new DateTime(Year, Month, Day);

                    paramList.Add("Year", Year.ToString());
                    paramList.Add("Month", Month.ToString());
                    paramList.Add("Day", Day.ToString());
                    paramList.Add("RunNumber", RunNumber.ToString());
                    FileItem item = null;
                    foreach (FileItem fileItem in comboBoxSubsectorNames.Items)
                    {
                        if (fileItem.Name.StartsWith(Subsector))
                        {
                            item = fileItem;
                            break;
                        }
                    }
                    paramList.Add("SubsectorTVItemID", item.TVItemID.ToString());
                    paramList.Add("SamplingPlanType", ((int)SamplingPlanType).ToString());
                    paramList.Add("SampleType", ((int)SampleType).ToString());
                    paramList.Add("LabSheetType", ((int)LabSheetType).ToString());

                    webClient.Headers.Add("Content-Type", "application/x-www-form-urlencoded");
                    byte[] ret = webClient.UploadValues(new Uri("http://cssplabsheet.azurewebsites.net/LabSheetExistAndStatus.aspx"), "POST", paramList);
                    //byte[] ret = webClient.UploadValues(new Uri("http://localhost:7668/LabSheetExistAndStatus.aspx"), "POST", paramList);

                    if (ret.Length > 0)
                    {
                        return System.Text.Encoding.Default.GetString(ret);
                    }
                }

            }
            catch (Exception ex)
            {
                return ex.Message + (ex.InnerException == null ? "" : ex.InnerException.Message);
            }

            return "";
        }
        private void GetTides()
        {
            if (CSSPWQInputParamCurrent.sidList.Count == 0)
            {
                lblStatus.Text = "Tide site missing";
                return;
            }

            timerGetTides.Enabled = false;
            webBrowserCSSP.Navigate(new Uri("http://www.tides.gc.ca/eng/station?type=0&date=" + dateTimePickerRun.Value.Year + "%2F" +
                         dateTimePickerRun.Value.Month + "%2F" + dateTimePickerRun.Value.Day + "&sid=" + CSSPWQInputParamCurrent.sidList[TideToTryIndex]));
        }
        private string GetTideText()
        {
            string StartTideTxt = "--";
            string EndTideTxt = "--";

            if (!ReadFileFromLocalMachine())
                return "-- / --";

            List<float> TideByHourList = new List<float>();
            foreach (HtmlElement heAll in webBrowserCSSP.Document.All)
            {
                if (heAll.TagName.ToLower() == "table")
                {
                    HtmlElement heTable = heAll;
                    if (heTable.GetAttribute("title") == "Predicted Hourly Heights (m)")
                    {
                        HtmlElement heTBody = heTable.Children[1];
                        HtmlElement heRow = heTBody.Children[0];
                        if (heRow.Children.Count != 25)
                        {
                            textBoxTides.Text = "-- / --";
                            return textBoxTides.Text;
                        }
                        if (heRow.Children[0].TagName.ToLower() == "th")
                        {
                            string dateTxt = heRow.Children[0].InnerText.Trim();
                            if (dateTxt.Substring(0, 4) != dateTimePickerRun.Value.Year.ToString())
                            {
                                textBoxTides.Text = "-- / --";
                                return textBoxTides.Text;
                            }

                            if (dateTxt.Substring(5, 2) != (dateTimePickerRun.Value.Month > 9 ? dateTimePickerRun.Value.Month.ToString() : "0" + dateTimePickerRun.Value.Month.ToString()))
                            {
                                textBoxTides.Text = "-- / --";
                                return textBoxTides.Text;
                            }

                            if (dateTxt.Substring(8, 2) != (dateTimePickerRun.Value.Day > 9 ? dateTimePickerRun.Value.Day.ToString() : "0" + dateTimePickerRun.Value.Day.ToString()))
                            {
                                textBoxTides.Text = "-- / --";
                                return textBoxTides.Text;
                            }
                        }
                        for (int i = 1; i < 25; i++)
                        {
                            TideByHourList.Add(float.Parse(heRow.Children[i].InnerText));
                        }

                        string TimeStartTxt = ((DateTime)labSheetA1Sheet.LabSheetA1MeasurementList.Where(c => c.Time != null).OrderBy(c => c.Time).First().Time).ToString("HH:mm");
                        string TimeEndTxt = ((DateTime)labSheetA1Sheet.LabSheetA1MeasurementList.Where(c => c.Time != null).OrderByDescending(c => c.Time).First().Time).ToString("HH:mm");

                        if (string.IsNullOrWhiteSpace(TimeStartTxt) || string.IsNullOrWhiteSpace(TimeEndTxt))
                        {
                            textBoxTides.Text = "-- / --";
                            return textBoxTides.Text;
                        }

                        int HourStart = int.Parse(TimeStartTxt.Substring(0, 2));
                        int MinuteStart = int.Parse(TimeStartTxt.Substring(3, 2));
                        if (MinuteStart > 30)
                        {
                            HourStart += 1;
                        }
                        int HourEnd = int.Parse(TimeEndTxt.Substring(0, 2));
                        int MinuteEnd = int.Parse(TimeStartTxt.Substring(3, 2));
                        if (MinuteEnd > 30)
                        {
                            HourEnd += 1;
                        }

                        float TideMax = TideByHourList.Max();
                        float TideMin = TideByHourList.Min();

                        float MinMidTide = TideMin + ((TideMax - TideMin) * (1.0f / 3.0f));
                        float MaxMidTide = TideMin + ((TideMax - TideMin) * (2.0f / 3.0f));

                        if (TideByHourList[HourStart] < MinMidTide)
                        {
                            StartTideTxt = "L";
                        }
                        else if (TideByHourList[HourStart] > MaxMidTide)
                        {
                            StartTideTxt = "H";
                        }
                        else
                        {
                            StartTideTxt = "M";
                        }

                        if (HourStart > 22)
                        {
                            for (int i = 22; i < HourStart; i++)
                            {
                                TideByHourList.Add(TideByHourList.Last());
                            }
                        }

                        if (TideByHourList[HourStart - 1] < TideByHourList[HourStart] || TideByHourList[HourStart] < TideByHourList[HourStart + 1])
                        {
                            StartTideTxt += "R";
                        }
                        else if (TideByHourList[HourStart - 1] > TideByHourList[HourStart] || TideByHourList[HourStart] > TideByHourList[HourStart + 1])
                        {
                            StartTideTxt += "F";
                        }
                        else
                        {
                            StartTideTxt += "T";
                        }


                        if (TideByHourList[HourEnd] < MinMidTide)
                        {
                            EndTideTxt = "L";
                        }
                        else if (TideByHourList[HourEnd] > MaxMidTide)
                        {
                            EndTideTxt = "H";
                        }
                        else
                        {
                            EndTideTxt = "M";
                        }

                        if (TideByHourList[HourEnd - 1] < TideByHourList[HourEnd] || TideByHourList[HourEnd] < TideByHourList[HourEnd + 1])
                        {
                            EndTideTxt += "R";
                        }
                        else if (TideByHourList[HourEnd - 1] > TideByHourList[HourEnd] || TideByHourList[HourEnd] > TideByHourList[HourEnd + 1])
                        {
                            EndTideTxt += "F";
                        }
                        else
                        {
                            EndTideTxt += "T";
                        }

                    }
                }
            }
            return StartTideTxt + " / " + EndTideTxt;
        }
        private string GetVariableText(int SpaceLength, string VariableText)
        {
            string blankStr = "                                                         ";
            int length = (SpaceLength - VariableText.Length);

            if (length < 0)
                length = 0;

            return VariableText + blankStr.Substring(0, length) + "\t";
        }
        private void LoadFileList()
        {
            richTextBoxFile.Text = "";
            listBoxFiles.Items.Clear();
            listBoxFiles.DisplayMember = "Text";
            listBoxFiles.ValueMember = "FileName";
            if (dateTimePickerArchiveFilterFrom.Value == null || dateTimePickerArchiveFilterTo.Value == null || comboBoxFileSubsector.SelectedItem == null)
                return;


            DirectoryInfo di = new DirectoryInfo(CurrentPath);

            if (!di.Exists)
            {
                lblStatus.Text = "Could not find directory [" + CurrentPath + "]";
                return;
            }

            int FromYear = dateTimePickerArchiveFilterFrom.Value.Year;
            int FromMonth = dateTimePickerArchiveFilterFrom.Value.Month;
            int FromDay = dateTimePickerArchiveFilterFrom.Value.Day;
            int ToYear = dateTimePickerArchiveFilterTo.Value.Year;
            int ToMonth = dateTimePickerArchiveFilterTo.Value.Month;
            int ToDay = dateTimePickerArchiveFilterTo.Value.Day;
            string Subsector = comboBoxFileSubsector.SelectedItem.ToString();

            List<FileInfo> fileList = new List<FileInfo>();
            for (int year = FromYear; year < ToYear + 1; year++)
            {
                DirectoryInfo diYear = new DirectoryInfo(CurrentPath + year.ToString() + @"\");
                if (!diYear.Exists)
                    continue;

                string FromText = FromYear.ToString() + "_"
                    + (FromMonth > 9 ? FromMonth.ToString() : "0" + FromMonth.ToString()) + "_"
                    + (FromDay > 9 ? FromDay.ToString() : "0" + FromDay.ToString()) + "_";
                string ToText = ToYear.ToString() + "_"
                    + (ToMonth > 9 ? ToMonth.ToString() : "0" + ToMonth.ToString()) + "_"
                    + (ToDay > 9 ? ToDay.ToString() : "0" + ToDay.ToString()) + "_";

                if (FileListViewTotalColiformLabSheets)
                {
                    diYear = new DirectoryInfo(diYear + (FileListViewTotalColiformLabSheets ? @"Total Coliform\" : ""));
                }
                if (!diYear.Exists)
                {
                    try
                    {
                        diYear.Create();
                    }
                    catch (Exception ex)
                    {
                        lblStatus.Text = "Could not create directory [" + diYear.FullName + "] " + ex.Message;
                        return;
                    }
                    diYear = new DirectoryInfo(diYear.FullName);
                }
                List<FileInfo> tempFileInfoList = new List<FileInfo>();

                foreach (FileInfo fileInfo in diYear.GetFiles().Where(c => c.Name.EndsWith(".txt")))
                {
                    string DateText = fileInfo.Name.Substring(fileInfo.Name.IndexOf("_") + 1, FromText.Length);

                    if (DateText.CompareTo(FromText) >= 0)
                    {
                        if (DateText.CompareTo(ToText) <= 0)
                        {
                            tempFileInfoList.Add(fileInfo);
                        }
                    }
                }

                foreach (FileInfo fiA in tempFileInfoList)
                {
                    if (FileListViewTotalColiformLabSheets)
                    {
                        fileList.Add(fiA);
                    }
                    else
                    {
                        if (FileListOnlyChangedAndRejected)
                        {
                            if (fiA.FullName.EndsWith("_C.txt") || fiA.FullName.EndsWith("_R.txt"))
                            {
                                fileList.Add(fiA);
                            }
                        }
                        else
                        {
                            fileList.Add(fiA);
                        }
                    }
                }

            }

            if (Subsector != "All")
            {
                string PartialSubsector = Subsector.Substring(0, Subsector.IndexOf(" ") - 1);
                fileList = fileList.Where(c => c.FullName.Contains(PartialSubsector)).ToList();
            }

            string OldSubsector = "";
            foreach (FileInfo fi in fileList)
            {
                FileStatusEnum fileStatus = FileStatusEnum.Error;

                string FileName = fi.FullName.Replace(CurrentPath, "");
                FileName = FileName.Replace(FileName.Substring(0, 4) + @"\" + (FileListViewTotalColiformLabSheets ? @"Total Coliform\" : ""), "");
                string CurrentSubsector = FileName.Substring(0, FileName.IndexOf("_"));
                string LabSheetDateTxt = FileName.Substring(FileName.IndexOf("_") + 1);
                DateTime LabSheetDate = new DateTime(int.Parse(LabSheetDateTxt.Substring(0, 4)), int.Parse(LabSheetDateTxt.Substring(5, 2)), int.Parse(LabSheetDateTxt.Substring(8, 2)));

                switch (FileName.Substring(FileName.Length - 6))
                {
                    case "_C.txt":
                        {
                            fileStatus = FileStatusEnum.Changed;
                        }
                        break;
                    case "_S.txt":
                        {
                            fileStatus = FileStatusEnum.Sent;
                        }
                        break;
                    case "_A.txt":
                        {
                            fileStatus = FileStatusEnum.Accepted;
                        }
                        break;
                    case "_R.txt":
                        {
                            fileStatus = FileStatusEnum.Rejected;
                        }
                        break;
                    case "_E.txt":
                        {
                            fileStatus = FileStatusEnum.Error;
                        }
                        break;
                    case "_F.txt":
                        {
                            fileStatus = FileStatusEnum.Fail;
                        }
                        break;
                    default:
                        break;
                }

                string SubsectorName = "";
                for (int i = 1, count = comboBoxFileSubsector.Items.Count; i < count; i++)
                {
                    if (((FileItem)comboBoxSubsectorNames.Items[i]).Name.Contains(CurrentSubsector))
                    {
                        SubsectorName = ((FileItem)comboBoxSubsectorNames.Items[i]).Name;
                        break;
                    }
                }

                if (SubsectorName == "")
                    continue;

                if (Subsector == "All")
                {
                    if (OldSubsector != SubsectorName)
                    {
                        listBoxFiles.Items.Add(new FileItemList(SubsectorName, ""));
                        OldSubsector = SubsectorName;
                    }
                }
                else
                {
                    if (fi.FullName.Contains(".txt") && fi.FullName.Contains(Subsector))
                    {
                        if (OldSubsector != SubsectorName)
                        {
                            listBoxFiles.Items.Add(new FileItemList(SubsectorName, ""));
                            OldSubsector = SubsectorName;
                        }
                    }
                }
                string RunText = fi.Name.Substring(fi.Name.IndexOf("_R") + 1);
                RunText = RunText.Substring(0, RunText.IndexOf("_"));
                listBoxFiles.Items.Add(new FileItemList(LabSheetDate.ToString("\tyyyy MMMMM dd") + "\t" + RunText + "\t" + fileStatus.ToString(), fi.FullName));
            }

            if (listBoxFiles.Items.Count > 0)
            {
                listBoxFiles.SelectedIndex = 0;
            }
        }
        private string MakeSureLabSheetFilesIsUniqueDocx(FileInfo fi)
        {
            List<FileInfo> fiList = (from c in fi.Directory.GetFiles()
                                     from p in PossibleLabSheetFileNamesDocx
                                     where fi.FullName.StartsWith(c.FullName.Substring(0, c.FullName.Length - 7))
                                     && c.FullName.Substring(c.FullName.Length - 7) == p
                                     && c.FullName.EndsWith(".docx")
                                     select c).ToList();

            if (fiList.Count > 1)
            {
                string retStr = "";
                foreach (FileInfo fi2 in fiList)
                {
                    retStr += fi2.FullName + "\r\n";
                }

                return retStr;
            }

            return "";
        }
        private string MakeSureLabSheetFilesIsUniqueTxt(FileInfo fi)
        {
            List<FileInfo> fiList = (from c in fi.Directory.GetFiles()
                                     from p in PossibleLabSheetFileNamesTxt
                                     where fi.FullName.StartsWith(c.FullName.Substring(0, c.FullName.Length - 6))
                                     && c.FullName.Substring(c.FullName.Length - 6) == p
                                     && c.FullName.EndsWith(".txt")
                                     select c).ToList();

            if (fiList.Count > 1)
            {
                string retStr = "";
                foreach (FileInfo fi2 in fiList)
                {
                    retStr += fi2.FullName + "\r\n";
                }

                return retStr;
            }

            return "";
        }
        private void Modifying()
        {
            if (InLoadingFile)
                return;

            if (lblFilePath.Text.Contains("F.txt") || lblFilePath.Text.Contains("A.txt") || lblFilePath.Text.Contains("S.txt"))
            {
                string EndFileText = lblFilePath.Text.Substring(lblFilePath.Text.Length - 5);
                string StatusText = "";
                switch (EndFileText)
                {
                    case "F.txt":
                        StatusText = "Fail";
                        break;
                    case "A.txt":
                        StatusText = "Accepted";
                        break;
                    case "S.txt":
                        StatusText = "Sent To Server";
                        break;
                    default:
                        break;
                }

                if (DialogResult.Yes != MessageBox.Show("Are you sure you want to change this file.\r\n\r\n" +
                    "It will change the status of the file from [" + StatusText + "] to [Changed]", "Warning", MessageBoxButtons.YesNo))
                    return;
            }

            lblStatus.Text = "Modified";
            butSendToServer.Text = "Saving ...";
            butSendToServer.Enabled = false;
            butGetLabSheetsStatus.Enabled = false;
            butGetTides.Enabled = false;
            butViewFCForm.Enabled = false;
            if (!timerSave.Enabled)
            {
                timerSave.Enabled = true;
                timerSave.Start();
            }
        }
        private void OnlyChangedAndRejected()
        {
            if (checkBoxOnlyChangedAndRejected.Checked == true)
            {
                FileListOnlyChangedAndRejected = true;
            }
            else
            {
                FileListOnlyChangedAndRejected = false;
            }
            LoadFileList();
        }
        private void OpenSamplingPlanFile(bool IsTest)
        {
            DirectoryInfo di = new DirectoryInfo(RootCurrentPath);

            if (!di.Exists)
            {
                try
                {
                    di.Create();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message + (ex.InnerException != null ? ex.InnerException.Message : ""), "Error while trying to create a directory [" + RootCurrentPath + @"] under C:\");
                    return;
                }
            }

            if (IsTest)
            {
                openFileDialogCSSP.FileName = TestFile;
            }
            else
            {
                openFileDialogCSSP.FileName = "SamplingPlan*.txt";
                openFileDialogCSSP.InitialDirectory = RootCurrentPath;
                openFileDialogCSSP.Filter = "SamplingPlan*.txt|Sampling Plan file";
                if (openFileDialogCSSP.ShowDialog() != DialogResult.OK)
                {
                    lblStatus.Text = "Please select a sampling plan file.";
                    return;
                }
            }
            if (openFileDialogCSSP.FileName.Length > 0)
            {
                lblSamplingPlanFileName.Text = openFileDialogCSSP.FileName;
                SamplingPlanName = openFileDialogCSSP.FileName;

                if (ReadSamplingPlan())
                {
                    panelAccessCode.Visible = true;
                }
                textBoxInitials.Focus();

                if (!string.IsNullOrWhiteSpace(csspWQInputApp.ApprovalCode))
                {
                    butApprove.Enabled = true;
                }
                else
                {
                    butApprove.Enabled = false;
                }

                //textBoxInitials.Text = "AA";
                //Initials = textBoxInitials.Text;
                //textBoxAccessCode.Focus();
                //textBoxAccessCode.Text = "Microlab12";

                AdjustVisualForIncludeLaboratoryQAQC();
            }
        }
        private void OpenFileName()
        {
            if (listBoxFiles.SelectedItem == null)
            {
                lblFilePath.Text = "";
                return;
            }

            string FileName = ((FileItemList)listBoxFiles.SelectedItem).FileName;
            if (string.IsNullOrWhiteSpace(FileName))
            {
                richTextBoxFile.Text = "Select a date";
                butViewFCForm.Visible = false;
                butOpen.Enabled = false;
                return;
            }

            string Rest = FileName.Replace(CurrentPath, "");
            Rest = Rest.Replace(Rest.Substring(0, 5) + (FileListViewTotalColiformLabSheets ? @"Total Coliform\" : ""), "");
            int Pos = Rest.IndexOf("_");
            string Subsector = Rest.Substring(0, Pos);
            string Year = Rest.Substring(Pos + 1, 4);
            string Month = Rest.Substring(Pos + 6, 2);
            string Day = Rest.Substring(Pos + 9, 2);

            NoUpdate = true;

            foreach (FileItem item in comboBoxSubsectorNames.Items)
            {
                if (item.Name.Contains(Subsector))
                {
                    comboBoxSubsectorNames.SelectedItem = item;
                    break;
                }
            }

            string RunText = Rest.Substring(Rest.IndexOf("_R") + 1);
            RunText = RunText.Substring(0, RunText.IndexOf("_"));

            foreach (string r in comboBoxRunNumber.Items)
            {
                if ("R" + r == RunText)
                {
                    comboBoxRunNumber.SelectedItem = r;
                    break;
                }
            }

            DateTime dateTimeRun = new DateTime(int.Parse(Year), int.Parse(Month), int.Parse(Day));

            dateTimePickerRun.Value = dateTimeRun;

            NoUpdate = false;

            UpdatePanelApp();

        }
        private void PostErrorWhileReadingFile(StreamReader sr, int LineNumber, string ErrorTxt)
        {
            InLoadingFile = false;
            lblStatus.Text = "Error reading file at line [" + LineNumber + "]. " + ErrorTxt;
        }
        private string PostLabSheet()
        {
            string retStr = FillInternetConnectionVariable();
            if (!string.IsNullOrWhiteSpace(retStr))
            {
                return retStr;
            }

            try
            {
                using (WebClient webClient = new WebClient())
                {
                    WebProxy webProxy = new WebProxy();
                    webClient.Proxy = webProxy;

                    NameValueCollection paramList = new NameValueCollection();

                    if (lblFilePath.Text == "")
                    {
                        return "File not identified";
                    }
                    paramList.Add("SamplingPlanName", lblSamplingPlanFileName.Text);
                    string Rest = lblFilePath.Text.Replace(CurrentPath, "");
                    int Pos = Rest.IndexOf("_");
                    string Subsector = Rest.Substring(0, Pos);
                    int Year = 0;
                    int.TryParse(Rest.Substring(Pos + 1, 4), out Year);
                    if (Year == 0)
                    {
                        return "Year not found";
                    }
                    int Month = 0;
                    int.TryParse(Rest.Substring(Pos + 6, 2), out Month);
                    if (Month == 0)
                    {
                        return "Month not found";
                    }
                    int Day = 0;
                    int.TryParse(Rest.Substring(Pos + 9, 2), out Day);
                    if (Day == 0)
                    {
                        return "Day not found";
                    }

                    DateTime dateTimeRun = new DateTime(Year, Month, Day);

                    paramList.Add("Year", Year.ToString());
                    paramList.Add("Month", Month.ToString());
                    paramList.Add("Day", Day.ToString());
                    paramList.Add("RunNumber", labSheetA1Sheet.RunNumber.ToString());
                    if (comboBoxSubsectorNames.SelectedItem == null)
                    {
                        return "Subsector not selected";
                    }
                    FileItem item = (FileItem)comboBoxSubsectorNames.SelectedItem;
                    if (item.TVItemID == 0)
                    {
                        return "Subsector not selected";
                    }
                    paramList.Add("SubsectorTVItemID", item.TVItemID.ToString());
                    paramList.Add("SamplingPlanType", ((int)labSheetA1Sheet.SamplingPlanType).ToString());
                    paramList.Add("SampleType", ((int)labSheetA1Sheet.SampleType).ToString());
                    paramList.Add("LabSheetType", ((int)labSheetA1Sheet.LabSheetType).ToString());
                    FileInfo fi = new FileInfo(lblFilePath.Text);
                    if (!fi.Exists)
                    {
                        return "File [" + fi.FullName + "] does not exist";
                    }
                    paramList.Add("FileName", lblFilePath.Text);
                    paramList.Add("FileLastModifiedDate_Local", fi.LastWriteTime.Year + "," + fi.LastWriteTime.Month +
                        "," + fi.LastWriteTime.Day + "," + fi.LastWriteTime.Hour + "," + fi.LastWriteTime.Minute + "," + fi.LastWriteTime.Second);

                    TextReader tr = fi.OpenText();
                    string FileContent = tr.ReadToEnd();
                    tr.Close();

                    paramList.Add("FileContent", FileContent);

                    webClient.Headers.Add("Content-Type", "application/x-www-form-urlencoded");
                    byte[] ret = webClient.UploadValues(new Uri("http://cssplabsheet.azurewebsites.net/PostLabSheet.aspx"), "POST", paramList);
                    //byte[] ret = webClient.UploadValues(new Uri("http://localhost:7668/PostLabSheet.aspx"), "POST", paramList);

                    if (ret.Length > 0)
                    {
                        return System.Text.Encoding.Default.GetString(ret);
                    }
                }

            }
            catch (Exception ex)
            {
                return ex.Message + (ex.InnerException == null ? "" : ex.InnerException.Message);
            }

            return "";
        }
        private void RadioButtonBathNumberChanged()
        {
            if (!csspWQInputApp.IncludeLaboratoryQAQC)
                return;

            if (radioButton3Baths.Checked)
            {
                if (labSheetA1Sheet != null && labSheetA1Sheet.WaterBathCount != 3)
                {
                    labSheetA1Sheet.WaterBathCount = 3;
                    AddLog("Water Bath Count", "3");
                }

                panelAppInputTopIncubation.Height = 200;
                panelAppInputTopTideCrew.Height = 200;
                panelTC.Height = 200;
                panelControl.Height = 200;
                panelAppInputTop.Height = 209;

                lblBath2IncubationTimeTxt.Visible = true;
                lblIncubationBath2TimeCalculated.Visible = true;
                lblBath2.Visible = true;
                textBoxIncubationBath2StartTime.Visible = true;
                textBoxIncubationBath2EndTime.Visible = true;
                textBoxWaterBath2Number.Visible = true;
                lbl2.Visible = true;
                lblBath2_44_5C.Visible = true;
                textBoxControlBath2Positive44_5.Visible = true;
                textBoxControlBath2NonTarget44_5.Visible = true;
                textBoxControlBath2Negative44_5.Visible = true;
                textBoxControlBath2Blank44_5.Visible = true;

                lblBath3IncubationTimeTxt.Visible = true;
                lblIncubationBath3TimeCalculated.Visible = true;
                lblBath3.Visible = true;
                textBoxIncubationBath3StartTime.Visible = true;
                textBoxIncubationBath3EndTime.Visible = true;
                textBoxWaterBath3Number.Visible = true;
                lbl3.Visible = true;
                lblBath3_44_5C.Visible = true;
                textBoxControlBath3Positive44_5.Visible = true;
                textBoxControlBath3NonTarget44_5.Visible = true;
                textBoxControlBath3Negative44_5.Visible = true;
                textBoxControlBath3Blank44_5.Visible = true;
            }
            else if (radioButton2Baths.Checked)
            {
                if (labSheetA1Sheet != null && labSheetA1Sheet.WaterBathCount != 2)
                {
                    labSheetA1Sheet.WaterBathCount = 2;
                    AddLog("Water Bath Count", "2");
                }

                panelAppInputTopIncubation.Height = 164;
                panelAppInputTopTideCrew.Height = 164;
                panelTC.Height = 164;
                panelControl.Height = 164;
                panelAppInputTop.Height = 173;

                lblBath2IncubationTimeTxt.Visible = true;
                lblIncubationBath2TimeCalculated.Visible = true;
                lblBath2.Visible = true;
                textBoxIncubationBath2StartTime.Visible = true;
                textBoxIncubationBath2EndTime.Visible = true;
                textBoxWaterBath2Number.Visible = true;
                lbl2.Visible = true;
                lblBath2_44_5C.Visible = true;
                textBoxControlBath2Positive44_5.Visible = true;
                textBoxControlBath2NonTarget44_5.Visible = true;
                textBoxControlBath2Negative44_5.Visible = true;
                textBoxControlBath2Blank44_5.Visible = true;

                lblBath3IncubationTimeTxt.Visible = false;
                lblIncubationBath3TimeCalculated.Visible = false;
                lblBath3.Visible = false;
                textBoxIncubationBath3StartTime.Visible = false;
                textBoxIncubationBath3EndTime.Visible = false;
                textBoxWaterBath3Number.Visible = false;
                lbl3.Visible = false;
                lblBath3_44_5C.Visible = false;
                textBoxControlBath3Positive44_5.Visible = false;
                textBoxControlBath3NonTarget44_5.Visible = false;
                textBoxControlBath3Negative44_5.Visible = false;
                textBoxControlBath3Blank44_5.Visible = false;

            }
            else
            {
                if (labSheetA1Sheet != null && labSheetA1Sheet.WaterBathCount != 1)
                {
                    labSheetA1Sheet.WaterBathCount = 1;
                    AddLog("Water Bath Count", "1");
                }

                panelAppInputTopIncubation.Height = 133;
                panelAppInputTopTideCrew.Height = 133;
                panelTC.Height = 133;
                panelControl.Height = 133;
                panelAppInputTop.Height = 142;

                lblBath2IncubationTimeTxt.Visible = false;
                lblIncubationBath2TimeCalculated.Visible = false;
                lblBath2.Visible = false;
                textBoxIncubationBath2StartTime.Visible = false;
                textBoxIncubationBath2EndTime.Visible = false;
                textBoxWaterBath2Number.Visible = false;
                lbl2.Visible = false;
                lblBath2_44_5C.Visible = false;
                textBoxControlBath2Positive44_5.Visible = false;
                textBoxControlBath2NonTarget44_5.Visible = false;
                textBoxControlBath2Negative44_5.Visible = false;
                textBoxControlBath2Blank44_5.Visible = false;

                lblBath3IncubationTimeTxt.Visible = false;
                lblIncubationBath3TimeCalculated.Visible = false;
                lblBath3.Visible = false;
                textBoxIncubationBath3StartTime.Visible = false;
                textBoxIncubationBath3EndTime.Visible = false;
                textBoxWaterBath3Number.Visible = false;
                lbl3.Visible = false;
                lblBath3_44_5C.Visible = false;
                textBoxControlBath3Positive44_5.Visible = false;
                textBoxControlBath3NonTarget44_5.Visible = false;
                textBoxControlBath3Negative44_5.Visible = false;
                textBoxControlBath3Blank44_5.Visible = false;
            }
            if (!InLoadingFile)
            {
                Modifying();
            }
        }
        private bool ReadSamplingPlan()
        {
            bool HasSubsector = false;
            bool HasInfrastructure = false;
            labSheetA1Sheet = new LabSheetA1Sheet();
            csspWQInputParamList = new List<CSSPWQInputParam>();
            FileInfo fi = new FileInfo(SamplingPlanName);

            if (!fi.Exists)
            {
                lblStatus.Text = fi.FullName + " does not exist.";

                return false;
            }

            csspWQInputApp.IncludeLaboratoryQAQC = true;
            string OldLineObj = "";
            StreamReader sr = fi.OpenText();
            int LineNumb = 0;
            while (!sr.EndOfStream)
            {
                LineNumb += 1;
                string LineTxt = sr.ReadLine();
                if (string.IsNullOrWhiteSpace(LineTxt))
                {
                    lblStatus.Text = "Sampling Plan File was not read properly. We found an empty line. The Sampling Plan file should not have empty lines.";
                    return false;
                }
                int pos = LineTxt.IndexOf("\t");
                int pos2 = LineTxt.IndexOf("\t", pos + 1);
                int pos3 = LineTxt.IndexOf("\t", pos2 + 1);
                switch (LineTxt.Substring(0, pos))
                {
                    case "Version":
                        {
                            VersionOfSamplingPlanFile = int.Parse(LineTxt.Substring("Version\t".Length));
                        }
                        break;
                    case "Sampling Plan Type":
                        {
                            SamplingPlanType = LineTxt.Substring("Samping Plan Type\t".Length).Trim();
                            lblSamplingPlanType.Text = SamplingPlanType;
                        }
                        break;
                    case "Sample Type":
                        {
                            SampleType = LineTxt.Substring("Sample Type\t".Length).Trim();
                            lblSampleType.Text = SampleType;
                        }
                        break;
                    case "Lab Sheet Type":
                        {
                            LabSheetType = LineTxt.Substring("Lab Sheet Type\t".Length).Trim();
                            lblLabSheetType.Text = LabSheetType;
                        }
                        break;
                    case "Subsector":
                        {
                            HasSubsector = true;
                            if (OldLineObj != "Lab Sheet Type" && OldLineObj != "Daily Duplicate" && OldLineObj != "MWQM Sites")
                            {
                                lblStatus.Text = "Could not read Sampling Plan file. Error at line " + LineNumb + ". Subsector line need to follow either [LabSheet Type, Daily Duplicate, MWQM Sites] line.";
                                return false;
                            }
                            CSSPWQInputParam csspWQInputParam = new CSSPWQInputParam();
                            csspWQInputParam.Name = LineTxt.Substring(pos + 1, pos2 - pos - 1);
                            csspWQInputParam.TVItemID = LineTxt.Substring(pos2 + 1, pos3 - pos2 - 1).Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).Select(Int32.Parse).ToList().First();
                            csspWQInputParam.sidList = LineTxt.Substring(pos3 + 1).Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();
                            csspWQInputParam.CSSPWQInputType = CSSPWQInputTypeEnum.Subsector;
                            csspWQInputParamList.Add(csspWQInputParam);
                        }
                        break;
                    case "MWQM Sites":
                        {
                            if (OldLineObj != "Subsector")
                            {
                                lblStatus.Text = "Could not read Sampling Plan File. Error at line " + LineNumb + ". MWQMSite line need to follow Subsector line.";
                                return false;
                            }
                            csspWQInputParamList.Last().MWQMSiteList = LineTxt.Substring(pos + 1, pos2 - pos - 1).Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();
                            csspWQInputParamList.Last().MWQMSiteTVItemIDList = LineTxt.Substring(pos2 + 1).Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).Select(Int32.Parse).ToList();

                            if (csspWQInputParamList.Last().MWQMSiteList.Count != csspWQInputParamList.Last().MWQMSiteList.Count)
                            {
                                lblStatus.Text = "Could not read Sampling Plan File. Error at line " + LineNumb + "";
                                return false;
                            }
                        }
                        break;
                    case "Daily Duplicate":
                        {
                            if (OldLineObj != "MWQM Sites")
                            {
                                lblStatus.Text = "Could not read Sampling Plan File. Error at line " + LineNumb + ". Duplicate line need to follow MWQM Sites line.";
                                return false;
                            }
                            csspWQInputParamList.Last().DailyDuplicateMWQMSiteList = LineTxt.Substring(pos + 1, pos2 - pos - 1).Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();
                            csspWQInputParamList.Last().DailyDuplicateMWQMSiteTVItemIDList = LineTxt.Substring(pos2 + 1).Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).Select(Int32.Parse).ToList();

                            if (csspWQInputParamList.Last().DailyDuplicateMWQMSiteList.Count != csspWQInputParamList.Last().DailyDuplicateMWQMSiteList.Count)
                            {
                                lblStatus.Text = "Could not read Sampling Plan File. Error at line " + LineNumb + "";
                                return false;
                            }
                        }
                        break;
                    case "Municipality":
                        {
                            HasInfrastructure = true;
                            CSSPWQInputParam csspWQInputParam = new CSSPWQInputParam();
                            csspWQInputParam.Name = LineTxt.Substring(pos + 1, pos2 - pos - 1);
                            csspWQInputParam.TVItemID = LineTxt.Substring(pos2 + 1).Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).Select(Int32.Parse).ToList().First();
                            csspWQInputParam.CSSPWQInputType = CSSPWQInputTypeEnum.Municipality;
                            csspWQInputParamList.Add(csspWQInputParam);
                        }
                        break;
                    case "Infrastructures":
                        {
                            if (OldLineObj != "Municipality")
                            {
                                lblStatus.Text = "Could not read Sampling Plan File. Error at line " + LineNumb + ". Infrastructure line need to follow Municipality line.";
                                return false;
                            }
                            csspWQInputParamList.Last().InfrastructureList = LineTxt.Substring(pos + 1, pos2 - pos - 1).Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();
                            csspWQInputParamList.Last().InfrastructureTVItemIDList = LineTxt.Substring(pos2 + 1).Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).Select(Int32.Parse).ToList();

                            if (csspWQInputParamList.Last().InfrastructureList.Count != csspWQInputParamList.Last().InfrastructureTVItemIDList.Count)
                            {
                                lblStatus.Text = "Could not read Sampling Plan File. Error at line " + LineNumb + "";
                                return false;
                            }
                        }
                        break;
                    case "App":
                        {
                            csspWQInputApp.AccessCode = GetCodeString(LineTxt.Substring(pos + 1, pos2 - pos - 1));
                            csspWQInputApp.ActiveYear = GetCodeString(LineTxt.Substring(pos2 + 1)).Trim();
                        }
                        break;
                    case "Precision Criteria":
                        {
                            csspWQInputApp.DailyDuplicatePrecisionCriteria = float.Parse(GetCodeString(LineTxt.Substring(pos + 1, pos2 - pos - 1)).Trim());
                            csspWQInputApp.IntertechDuplicatePrecisionCriteria = float.Parse(GetCodeString(LineTxt.Substring(pos2 + 1)).Trim());
                            textBoxDailyDuplicatePrecisionCriteria.Text = csspWQInputApp.DailyDuplicatePrecisionCriteria.ToString();
                            textBoxIntertechDuplicatePrecisionCriteria.Text = csspWQInputApp.IntertechDuplicatePrecisionCriteria.ToString();
                            timerSave.Enabled = false;
                        }
                        break;
                    case "Include Laboratory QA/QC":
                        {
                            bool IncludeLaboratoryQAQC;
                            if (!bool.TryParse(LineTxt.Substring(pos + 1, pos2 - pos - 1), out IncludeLaboratoryQAQC))
                            {
                                lblStatus.Text = "Could not read Sampling Plan File. Error at line " + LineNumb + "";
                                return false;
                            }
                            csspWQInputApp.IncludeLaboratoryQAQC = IncludeLaboratoryQAQC;
                            csspWQInputApp.ApprovalCode = GetCodeString(LineTxt.Substring(pos2 + 1)).Trim();
                        }
                        break;
                    default:
                        {
                            lblStatus.Text = "First item in line " + LineNumb + " not recognized [" + LineTxt.Substring(0, LineTxt.IndexOf("\t")) + "]";
                            return false;
                        }
                }

                OldLineObj = LineTxt.Substring(0, pos);
            }
            sr.Close();

            if (VersionOfSamplingPlanFile == 0)
            {
                lblStatus.Text = "Version was not read properly";
                return false;
            }

            if (!new List<string>() { "Subsector", "Infrastructure" }.Contains(SamplingPlanType))
            {
                lblStatus.Text = "Sampling Plan Type was not read properly. It need to be either Subsector or Infrastructure";
                return false;
            }

            if (SamplingPlanType == "Subsector" && HasInfrastructure)
            {
                lblStatus.Text = "Sampling Plan Type was not read properly. Municipality and Infrastructure cannot be contain in this file if you specify to by subsector type";
                return false;
            }

            if (SamplingPlanType == "Infrastructure" && HasSubsector)
            {
                lblStatus.Text = "Sampling Plan Type was not read properly. Subsector and MWQM Sites cannot be contain in this file if you specify to by infrastructure type";
                return false;
            }

            List<string> SampleTypeList = new List<string>()
            {
                SampleTypeEnum.Routine.ToString(),
                SampleTypeEnum.RainCMPRoutine.ToString(),
                SampleTypeEnum.RainRun.ToString(),
                SampleTypeEnum.ReopeningEmergencyRain.ToString(),
                SampleTypeEnum.ReopeningSpill.ToString(),
                SampleTypeEnum.Sanitary.ToString(),
                SampleTypeEnum.Study.ToString(),
            };
            if (!SampleTypeList.Contains(SampleType))
            {
                lblStatus.Text = "Sample Type was not read properly";
                return false;
            }

            if (LabSheetType == "" || !(LabSheetType == "A1" || LabSheetType == "EC" || LabSheetType == "LTB"))
            {
                lblStatus.Text = "Lab Sheet Type was not read properly";
                return false;
            }

            if (string.IsNullOrWhiteSpace(csspWQInputApp.AccessCode))
            {
                lblStatus.Text = "Access Code was not read properly";
                return false;
            }

            if (string.IsNullOrWhiteSpace(csspWQInputApp.ActiveYear))
            {
                lblStatus.Text = "Active Year was not read properly";
                return false;
            }

            if (string.IsNullOrWhiteSpace(csspWQInputApp.DailyDuplicatePrecisionCriteria.ToString()))
            {
                lblStatus.Text = "Daily Duplicate Precision Criteria was not read properly";
                return false;
            }

            if (string.IsNullOrWhiteSpace(csspWQInputApp.IntertechDuplicatePrecisionCriteria.ToString()))
            {
                lblStatus.Text = "Intertech Duplicate Precision Criteria was not read properly";
                return false;
            }

            lblStatus.Text = "Sampling Plan File OK.";
            return true;
        }
        private bool ReadFileFromLocalMachine()
        {
            labSheetA1Sheet = new LabSheetA1Sheet();
            InLoadingFile = true;
            if (lblFilePath.Text == "")
            {
                lblStatus.Text = "lblFilePath.Text is empty";
                return false;
            }

            FileInfo fi = new FileInfo(lblFilePath.Text);

            if (!fi.Exists)
            {
                lblStatus.Text = "Could not find file [" + fi.FullName + "]";
                return false;
            }

            StreamReader sr = fi.OpenText();
            string FullFileText = sr.ReadToEnd();
            sr.Close();

            labSheetA1Sheet = csspLabSheetParser.ParseLabSheetA1(FullFileText);
            if (!string.IsNullOrWhiteSpace(labSheetA1Sheet.Error))
            {
                lblStatus.Text = labSheetA1Sheet.Error;
                return false;
            }

            sbLog = new StringBuilder();
            if (labSheetA1Sheet.Log != null)
                sbLog.AppendLine(labSheetA1Sheet.Log.Trim());

            FillFormWithParsedLabSheetA1(labSheetA1Sheet);
            SetupDataGridViewCSSPA1(labSheetA1Sheet);
            SetButGetTidesEnabledOrNot();
            CalculateDuplicate();

            for (int row = 0, count = labSheetA1Sheet.LabSheetA1MeasurementList.Count; row < count; row++)
            {
                TryToCalculateMPNA1(row);
            }

            RadioButtonBathNumberChanged();

            InLoadingFile = false;
            textBoxSampleCrewInitials.Focus();

            if (fi.FullName.EndsWith("_S.txt"))
            {
                butSendToServer.Text = "Already saved on server";
                butSendToServer.Enabled = false;
            }
            else if (fi.FullName.EndsWith("_R.txt"))
            {
                butSendToServer.Text = "Rejected on server";
                butSendToServer.Enabled = true;
            }
            else if (fi.FullName.EndsWith("_A.txt"))
            {
                butSendToServer.Text = "Accepted on server";
                butSendToServer.Enabled = false;
            }
            else
            {
                butSendToServer.Text = "Send to Server";
                butSendToServer.Enabled = true;
            }

            butViewFCForm.Enabled = true;

            DoLogCheckForReediting();

            lblSupervisorInitials.Text = "";
            lblApprovalDate.Text = "";
            if (!string.IsNullOrWhiteSpace(labSheetA1Sheet.ApprovedBySupervisorInitials))
            {
                lblSupervisorInitials.Text = labSheetA1Sheet.ApprovedBySupervisorInitials;
            }
            if (!string.IsNullOrWhiteSpace(labSheetA1Sheet.ApprovalYear))
            {
                int Year;
                int Month;
                int Day;

                if (int.TryParse(labSheetA1Sheet.ApprovalYear, out Year))
                {
                    if (Year > 1981)
                    {
                        if (int.TryParse(labSheetA1Sheet.ApprovalMonth, out Month))
                        {
                            if (int.TryParse(labSheetA1Sheet.ApprovalDay, out Day))
                            {
                                csspWQInputApp.ApprovalDate = new DateTime(Year, Month, Day);
                                lblApprovalDate.Text = csspWQInputApp.ApprovalDate.ToString("yyyy MMMM dd");
                            }
                        }
                    }
                }
            }
            return true;
        }
        private void ReplaceFileFromTo(FileInfo fiFrom, FileInfo fiTo, bool FromLocal)
        {
            butFileArchiveCancel.Enabled = true;
            butFileArchiveSkip.Enabled = true;
            butFileArchiveCopy.Enabled = true;

            string LocalFileText = "";
            string ServerFileText = "";

            lblSendingFileName.Text = fiFrom.Name;

            lblStatus.Text = "Replacing - " + fiFrom.FullName + " to " + fiTo.FullName;
            richTextBoxLabSheetSender.AppendText("Replacing - " + fiFrom.FullName + "\r\n to " + fiTo.FullName + "\r\n");

            lblLocalFileDateTime.Text = fiFrom.LastWriteTime.ToString("yyyy MMMM dd HH:mm:ss");
            lblServerFileDateTime.Text = fiTo.LastWriteTime.ToString("yyyy MMMM dd HH:mm:ss");

            richTextBoxLabSheetReceiver.SelectionColor = Color.Black;
            richTextBoxLabSheetSender.SelectionColor = Color.Black;

            if (fiFrom.Extension == ".txt")
            {
                StreamReader sr = fiFrom.OpenText();
                LocalFileText = sr.ReadToEnd();
                sr.Close();

                richTextBoxLabSheetSender.Text = LocalFileText;

                StreamReader sr2 = fiTo.OpenText();
                ServerFileText = sr2.ReadToEnd();
                sr2.Close();

                richTextBoxLabSheetReceiver.Text = ServerFileText;

                CompareTheDocs();
            }
            else
            {
                richTextBoxLabSheetSender.Text = fiFrom.Extension + " document (no comparison)";
                richTextBoxLabSheetReceiver.Text = fiFrom.Extension + " document (no comparison)";
            }
            WaitingForUserAction = true;
            UserActionFileArchiveCopy = false;
            UserActionFileArchiveSkip = false;
            UserActionFileArchiveCancel = false;
            panelSenderTop.Focus();
            while (WaitingForUserAction)
            {
                Application.DoEvents();
                if (UserActionFileArchiveCopy)
                {
                    try
                    {
                        File.Delete(fiTo.FullName);
                        File.Copy(fiFrom.FullName, fiTo.FullName);
                    }
                    catch (Exception ex)
                    {
                        lblStatus.Text = "Error:" + ex.Message + (ex.InnerException != null ? ex.InnerException.Message : "");
                        richTextBoxLabSheetSender.AppendText("Error:" + ex.Message + (ex.InnerException != null ? ex.InnerException.Message : "") + "\r\n");
                        MessageBox.Show("Error: " + ex.Message + (ex.InnerException != null ? ex.InnerException.Message : ""), "Error");
                        WaitingForUserAction = false;
                        return;
                    }
                    richTextBoxLabSheetSender.AppendText("Replaced - " + fiFrom.FullName + "\r\n to " + fiTo.FullName + "\r\n");
                    WaitingForUserAction = false;
                }
                if (UserActionFileArchiveSkip)
                {
                    WaitingForUserAction = false;
                }
                if (UserActionFileArchiveCancel)
                {
                    WaitingForUserAction = false;
                    panelSendToServerCompare.SendToBack();
                }
            }

            return;
        }
        private void ResetBottomPanels()
        {
            if (panelAddInputBottomLeft.Width + panelAddInputBottomLeftDuplicate.Width + panelAddInputBottomRight.Width + 20 < Width)
            {
                AppIsWide = true;
            }
            else
            {
                AppIsWide = false;
            }

            if (csspWQInputApp.IncludeLaboratoryQAQC)
            {
                if (AppIsWide)
                {
                    panelAddInputBottomLeftDuplicate.Left = panelAddInputBottomLeft.Width + 20;
                    panelAddInputBottomLeftDuplicate.Visible = true;
                    panelAddInputBottomLeft.Visible = true;
                }
                else
                {
                    panelAddInputBottomLeftDuplicate.Left = 10;
                    if (IsOnDailyDuplicate)
                    {
                        panelAddInputBottomLeftDuplicate.Visible = true;
                        panelAddInputBottomLeft.Visible = false;
                    }
                    else
                    {
                        panelAddInputBottomLeftDuplicate.Visible = false;
                        panelAddInputBottomLeft.Visible = true;
                    }
                }
            }
        }
        private void SaveInfoOnLocalMachine(bool ForChangeDate)
        {
            string SampleCrewInitials = "";
            string IncubationStartSameDay = "";
            string IncubationBath1StartTime = "";
            string WaterBathCount = "";
            string IncubationBath2StartTime = "";
            string IncubationBath3StartTime = "";
            string IncubationBath1EndTime = "";
            string IncubationBath2EndTime = "";
            string IncubationBath3EndTime = "";
            string IncubationBath1TimeCalculated = "";
            string IncubationBath2TimeCalculated = "";
            string IncubationBath3TimeCalculated = "";
            string WaterBath1 = "";
            string WaterBath2 = "";
            string WaterBath3 = "";
            string TCField1 = "";
            string TCLab1 = "";
            string TCHas2Coolers = "";
            string TCField2 = "";
            string TCLab2 = "";
            string TCFirst = "";
            string TCAverage = "";
            string ControlLot = "";
            string Positive35 = "";
            string NonTarget35 = "";
            string Negative35 = "";
            string Bath1Positive44_5 = "";
            string Bath2Positive44_5 = "";
            string Bath3Positive44_5 = "";
            string Bath1NonTarget44_5 = "";
            string Bath2NonTarget44_5 = "";
            string Bath3NonTarget44_5 = "";
            string Bath1Negative44_5 = "";
            string Bath2Negative44_5 = "";
            string Bath3Negative44_5 = "";
            string Blank35 = "";
            string Bath1Blank44_5 = "";
            string Bath2Blank44_5 = "";
            string Bath3Blank44_5 = "";
            string Lot35 = "";
            string Lot44_5 = "";
            string DailyDuplicateRLogValue = "";
            string DailyDuplicatePrecisionCriteria = "";
            string DailyDuplicateAcceptableOrUnacceptable = "";
            string IntertechDuplicateRLogValue = "";
            string IntertechDuplicatePrecisionCriteria = "";
            string IntertechDuplicateAcceptableOrUnacceptable = "";
            string IntertechReadAcceptableOrUnacceptable = "";
            string SampleBottleLotNumber = "";
            string SalinitiesReadBy = "";
            DateTime? dateTimeCSSPSalinitiesRead = null;
            string ResultsReadBy = "";
            DateTime? dateTimeCSSPResultsRead = null;
            string ResultsRecordedBy = "";
            DateTime? dateTimeCSSPResultsRecorded = null;

            IsSaving = true;
            if (lblFilePath.Text.Length == 0)
            {
                lblStatus.Text = "No file";
                IsSaving = false;
                return;
            }

            StringBuilder sb = new StringBuilder();
            FileInfo fi = new FileInfo(lblFilePath.Text);

            if (!fi.Exists)
            {
                lblStatus.Text = "File does not exist [" + fi.FullName + "]";
                lblStatus.Text = "Not Saved";
                IsSaving = false;
                return;
            }

            FileItem item = (FileItem)comboBoxSubsectorNames.SelectedItem;
            if (item == null)
            {
                lblStatus.Text = "Subsector unknown";
                lblStatus.Text = "Not Saved";
                IsSaving = false;
                return;
            }

            DateTime dateTimeCSSP = dateTimePickerRun.Value;
            if (dateTimeCSSP == null)
            {
                lblStatus.Text = "Date unknown";
                lblStatus.Text = "Not Saved";
                IsSaving = false;
                return;
            }

            if (ForChangeDate)
            {
                dateTimeCSSP = dateTimePickerChangeDate.Value;
                if (dateTimeCSSP == null)
                {
                    lblStatus.Text = "Date unknown";
                    lblStatus.Text = "Not Saved";
                    IsSaving = false;
                    return;
                }
            }

            string TidesText = textBoxTides.Text;
            if (csspWQInputApp.IncludeLaboratoryQAQC)
            {
                SampleCrewInitials = textBoxSampleCrewInitials.Text;
                IncubationStartSameDay = (checkBoxIncubationStartSameDay.Checked ? "true" : "false");
                IncubationBath1StartTime = "";
                WaterBathCount = "1";
                if (radioButton2Baths.Checked)
                {
                    WaterBathCount = "2";
                }
                if (radioButton3Baths.Checked)
                {
                    WaterBathCount = "3";
                }
                if (textBoxIncubationBath1StartTime.Text.Length == 5)
                {
                    IncubationBath1StartTime = textBoxIncubationBath1StartTime.Text;
                }
                else
                {
                    IncubationBath1StartTime = "";
                }
                IncubationBath2StartTime = "";
                if (textBoxIncubationBath2StartTime.Text.Length == 5)
                {
                    IncubationBath2StartTime = (WaterBathCount != "1" ? textBoxIncubationBath2StartTime.Text : "");
                }
                else
                {
                    IncubationBath2StartTime = "";
                }
                IncubationBath3StartTime = "";
                if (textBoxIncubationBath3StartTime.Text.Length == 5)
                {
                    IncubationBath3StartTime = (WaterBathCount == "3" ? textBoxIncubationBath3StartTime.Text : "");
                }
                else
                {
                    IncubationBath3StartTime = "";
                }
                IncubationBath1EndTime = "";
                if (textBoxIncubationBath1EndTime.Text.Length == 5)
                {
                    IncubationBath1EndTime = textBoxIncubationBath1EndTime.Text;
                }
                else
                {
                    IncubationBath1EndTime = "";
                }
                IncubationBath2EndTime = "";
                if (textBoxIncubationBath2EndTime.Text.Length == 5)
                {
                    IncubationBath2EndTime = (WaterBathCount != "1" ? textBoxIncubationBath2EndTime.Text : "");
                }
                else
                {
                    IncubationBath2EndTime = "";
                }
                IncubationBath3EndTime = "";
                if (textBoxIncubationBath3EndTime.Text.Length == 5)
                {
                    IncubationBath3EndTime = (WaterBathCount == "3" ? textBoxIncubationBath3EndTime.Text : "");
                }
                else
                {
                    IncubationBath3EndTime = "";
                }
                IncubationBath1TimeCalculated = lblIncubationBath1TimeCalculated.Text;
                IncubationBath2TimeCalculated = (WaterBathCount != "1" ? lblIncubationBath2TimeCalculated.Text : "");
                IncubationBath3TimeCalculated = (WaterBathCount == "3" ? lblIncubationBath3TimeCalculated.Text : "");
                WaterBath1 = textBoxWaterBath1Number.Text;
                WaterBath2 = (WaterBathCount != "1" ? textBoxWaterBath2Number.Text : "");
                WaterBath3 = (WaterBathCount == "3" ? textBoxWaterBath3Number.Text : "");
                TCField1 = textBoxTCField1.Text;
                TCLab1 = textBoxTCLab1.Text;
                TCHas2Coolers = (checkBox2Coolers.Checked ? "true" : "false");
                TCField2 = textBoxTCField2.Text;
                TCLab2 = textBoxTCLab2.Text;
                TCFirst = lblTCFirst.Text;
                TCAverage = lblTCAverage.Text;
                ControlLot = textBoxControlLot.Text;
                Positive35 = textBoxControlPositive35.Text;
                NonTarget35 = textBoxControlNonTarget35.Text;
                Negative35 = textBoxControlNegative35.Text;
                Bath1Positive44_5 = textBoxControlBath1Positive44_5.Text;
                Bath2Positive44_5 = (WaterBathCount != "1" ? textBoxControlBath2Positive44_5.Text : "");
                Bath3Positive44_5 = (WaterBathCount == "3" ? textBoxControlBath3Positive44_5.Text : "");
                Bath1NonTarget44_5 = textBoxControlBath1NonTarget44_5.Text;
                Bath2NonTarget44_5 = (WaterBathCount != "1" ? textBoxControlBath2NonTarget44_5.Text : "");
                Bath3NonTarget44_5 = (WaterBathCount == "3" ? textBoxControlBath3NonTarget44_5.Text : "");
                Bath1Negative44_5 = textBoxControlBath1Negative44_5.Text;
                Bath2Negative44_5 = (WaterBathCount != "1" ? textBoxControlBath2Negative44_5.Text : "");
                Bath3Negative44_5 = (WaterBathCount == "3" ? textBoxControlBath3Negative44_5.Text : "");
                Blank35 = textBoxControlBlank35.Text;
                Bath1Blank44_5 = textBoxControlBath1Blank44_5.Text;
                Bath2Blank44_5 = (WaterBathCount != "1" ? textBoxControlBath2Blank44_5.Text : "");
                Bath3Blank44_5 = (WaterBathCount == "3" ? textBoxControlBath3Blank44_5.Text : "");
                Lot35 = textBoxLot35.Text;
                Lot44_5 = textBoxLot44_5.Text;
                DailyDuplicateRLogValue = lblDailyDuplicateRLogValue.Text;
                DailyDuplicatePrecisionCriteria = textBoxDailyDuplicatePrecisionCriteria.Text;
                DailyDuplicateAcceptableOrUnacceptable = lblDailyDuplicateAcceptableOrUnacceptable.Text;
                IntertechDuplicateRLogValue = lblIntertechDuplicateRLogValue.Text;
                IntertechDuplicatePrecisionCriteria = textBoxIntertechDuplicatePrecisionCriteria.Text;
                IntertechDuplicateAcceptableOrUnacceptable = lblIntertechDuplicateAcceptableOrUnacceptable.Text;
                IntertechReadAcceptableOrUnacceptable = lblIntertechReadAcceptableOrUnacceptable.Text;
                SampleBottleLotNumber = textBoxSampleBottleLotNumber.Text;

                SalinitiesReadBy = textBoxSalinitiesReadBy.Text;
                dateTimeCSSPSalinitiesRead = dateTimePickerSalinitiesReadDate.Value;
                if (dateTimeCSSPSalinitiesRead == null)
                {
                    lblStatus.Text = "Date of Salinities Read unknown";
                    lblStatus.Text = "Not Saved";
                    IsSaving = false;
                    return;
                }

                ResultsReadBy = textBoxResultsReadBy.Text;
                dateTimeCSSPResultsRead = dateTimePickerResultsReadDate.Value;
                if (dateTimeCSSPResultsRead == null)
                {
                    lblStatus.Text = "Date of Results Read unknown";
                    lblStatus.Text = "Not Saved";
                    IsSaving = false;
                    return;
                }

                ResultsRecordedBy = textBoxResultsRecordedBy.Text;
                dateTimeCSSPResultsRecorded = dateTimePickerResultsRecordedDate.Value;
                if (dateTimeCSSPResultsRecorded == null)
                {
                    lblStatus.Text = "Date of Results Recorded unknown";
                    lblStatus.Text = "Not Saved";
                    IsSaving = false;
                    return;
                }
            }
            string RunWeatherComment = richTextBoxRunWeatherComment.Text;
            string FieldComment = richTextBoxRunComment.Text;


            List<List<string>> SiteValuesList = new List<List<string>>();

            if (csspWQInputSheetType == CSSPWQInputSheetTypeEnum.A1)
            {

                List<string> RowValues = new List<string>()
                {
                    "Site", "Time", "MPN", "Tube10", "Tube1_0", "Tube0_1", "Sal", "Temp", "Proc. By", "Sample Type", "ID", "Comment"
                };

                SiteValuesList.Add(RowValues);

                for (int row = 0, countRow = dataGridViewCSSP.Rows.Count; row < countRow; row++)
                {
                    RowValues = new List<string>();

                    for (int col = 1, countCol = dataGridViewCSSP.Columns.Count; col < countCol; col++)
                    {
                        if (dataGridViewCSSP[col, row].Value == null)
                        {
                            RowValues.Add("");
                        }
                        else
                        {
                            RowValues.Add(dataGridViewCSSP[col, row].Value.ToString());
                        }
                    }

                    SiteValuesList.Add(RowValues);
                }
            }

            int LongestLength = "Intertech Duplicate Acceptable Or Unacceptable".Length;

            sb.Append("[Version|" + VersionOfResultFile + "]");
            sb.Append("[Sampling Plan Type|" + SamplingPlanType + "]");
            sb.Append("[Sample Type|" + SampleType + "]");
            sb.AppendLine("[Lab Sheet Type|" + LabSheetType + "]");
            sb.AppendLine("[Subsector|" + item.Name + "|" + item.TVItemID + "]");
            sb.Append("[Date|" + dateTimeCSSP.Year + "|" + dateTimeCSSP.Month + "|" + dateTimeCSSP.Day + "]");
            sb.Append("[Run|" + RunNumberCurrent + "]");
            sb.AppendLine("[Tides|" + TidesText + "]");
            if (csspWQInputApp.IncludeLaboratoryQAQC)
            {
                sb.AppendLine("[Sample Crew Initials|" + SampleCrewInitials + "]");
                sb.Append("[Incubation Start Same Day|" + IncubationStartSameDay + "]");
                sb.Append("[Water Bath Count|" + WaterBathCount + "]");
                sb.AppendLine("[Incubation Start Time|" + IncubationBath1StartTime + "|" + IncubationBath2StartTime + "|" + IncubationBath3StartTime + "]");
                sb.Append("[Incubation End Time|" + IncubationBath1EndTime + "|" + IncubationBath2EndTime + "|" + IncubationBath3EndTime + "]");
                sb.AppendLine("[Incubation Time Calculated|" + IncubationBath1TimeCalculated + "|" + IncubationBath2TimeCalculated + "|" + IncubationBath3TimeCalculated + "]");
                sb.Append("[Water Bath|" + WaterBath1 + "|" + WaterBath2 + "|" + WaterBath3 + "]");
                sb.Append("[TC Has 2 Coolers|" + TCHas2Coolers + "]");
                sb.AppendLine("[TC Field|" + TCField1 + "|" + TCField2 + "]");
                sb.Append("[TC Lab|" + TCLab1 + "|" + TCLab2 + "]");
                sb.Append("[TC First|" + TCFirst + "]");
                sb.AppendLine("[TC Average|" + TCAverage + "]");
                sb.Append("[Control Lot|" + ControlLot + "]");
                sb.Append("[Positive 35|" + Positive35 + "]");
                sb.Append("[Non Target 35|" + NonTarget35 + "]");
                sb.AppendLine("[Negative 35|" + Negative35 + "]");
                sb.Append("[Positive 44.5|" + Bath1Positive44_5 + "|" + Bath2Positive44_5 + "|" + Bath3Positive44_5 + "]");
                sb.Append("[Non Target 44.5|" + Bath1NonTarget44_5 + "|" + Bath2NonTarget44_5 + "|" + Bath3NonTarget44_5 + "]");
                sb.AppendLine("[Negative 44.5|" + Bath1Negative44_5 + "|" + Bath2Negative44_5 + "|" + Bath3Negative44_5 + "]");
                sb.Append("[Blank 35|" + Blank35 + "]");
                sb.Append("[Blank 44.5|" + Bath1Blank44_5 + "|" + Bath2Blank44_5 + "|" + Bath3Blank44_5 + "]");
                sb.Append("[Lot 35|" + Lot35 + "]");
                sb.AppendLine("[Lot 44.5|" + Lot44_5 + "]");
                sb.AppendLine("[Daily Duplicate|" + DailyDuplicateRLogValue + "|" + DailyDuplicatePrecisionCriteria + "|" + DailyDuplicateAcceptableOrUnacceptable + "]");
                sb.Append("[Intertech Duplicate|" + IntertechDuplicateRLogValue + "|" + IntertechDuplicatePrecisionCriteria + "|" + IntertechDuplicateAcceptableOrUnacceptable + "]");
                sb.AppendLine("[Intertech Read|" + IntertechReadAcceptableOrUnacceptable + "]");
            }
            sb.AppendLine("[Run Weather Comment|" + RunWeatherComment + "]");
            sb.AppendLine("[Run Comment|" + FieldComment + "]");
            if (csspWQInputApp.IncludeLaboratoryQAQC)
            {
                sb.AppendLine("[Sample Bottle Lot Number|" + SampleBottleLotNumber + "]");
                sb.Append("[Salinities|" + SalinitiesReadBy + "|" + ((DateTime)dateTimeCSSPSalinitiesRead).Year + "|" + ((DateTime)dateTimeCSSPSalinitiesRead).Month + "|" + ((DateTime)dateTimeCSSPSalinitiesRead).Day + "]");
                sb.Append("[Results|" + ResultsReadBy + "|" + ((DateTime)dateTimeCSSPResultsRead).Year + "|" + ((DateTime)dateTimeCSSPResultsRead).Month + "|" + ((DateTime)dateTimeCSSPResultsRead).Day + "]");
                sb.AppendLine("[Recorded|" + ResultsRecordedBy + "|" + ((DateTime)dateTimeCSSPResultsRecorded).Year + "|" + ((DateTime)dateTimeCSSPResultsRecorded).Month + "|" + ((DateTime)dateTimeCSSPResultsRecorded).Day + "]");
            }

            sb.AppendLine("[IncludeLaboratoryQAQC|" + csspWQInputApp.IncludeLaboratoryQAQC + "]");
            if (csspWQInputApp.IncludeLaboratoryQAQC)
            {
                sb.AppendLine("[Approved By Supervisor Initials|" + lblSupervisorInitials.Text + "]");
                sb.AppendLine("[Approval Date|" + csspWQInputApp.ApprovalDate.Year + "|" + csspWQInputApp.ApprovalDate.Month + "|" + csspWQInputApp.ApprovalDate.Day + "]");
            }

            foreach (List<string> RowValues in SiteValuesList)
            {
                int col = 0;
                foreach (string s in RowValues)
                {
                    if (col < 1)
                    {
                        sb.Append(GetVariableText(7, s.Trim()));
                    }
                    else if (col < 2)
                    {
                        sb.Append(GetVariableText(6, s.Trim()));
                    }
                    else if (col < 3)
                    {
                        string ss = s;
                        if (ss == "1")
                        {
                            ss = "< 2";
                        }
                        if (ss == "1700")
                        {
                            ss = "> 1600";
                        }
                        sb.Append(GetVariableText(10, ss.Trim()));
                    }
                    else if (col < 6)
                    {
                        sb.Append(GetVariableText(7, s.Trim()));
                    }
                    else if (col < 7)
                    {
                        sb.Append(GetVariableText(5, s.Trim()));
                    }
                    else if (col < 8)
                    {
                        sb.Append(GetVariableText(5, s.Trim()));
                    }
                    else if (col < 9)
                    {
                        sb.Append(GetVariableText(8, s.Trim()));
                    }
                    else if (col < 10)
                    {
                        sb.Append(GetVariableText(14, s.Trim()));
                    }
                    else if (col < 11)
                    {
                        sb.Append(GetVariableText(8, s.Trim()));
                    }
                    else
                    {
                        sb.Append(GetVariableText(12, s.Trim()));
                    }
                    col += 1;
                }
                sb.AppendLine("");
            }

            //sb.AppendLine("________________________________");
            //sb.AppendLine("Log");
            sb.AppendLine(sbLog.ToString());

            try
            {
                fi.Delete();
                fi = new FileInfo(fi.FullName);
                StreamWriter sw = fi.CreateText();
                sw.Write(sb.ToString());
                sw.Close();
            }
            catch (Exception ex)
            {
                lblStatus.Text = ex.Message + (ex.InnerException != null ? ex.InnerException.Message : "");
            }

            lblStatus.Text = "Saved";
            timerSave.Enabled = false;
            timerSave.Stop();

            SetButGetTidesEnabledOrNot();

            IsSaving = false;
        }
        private void SendToServer()
        {
            if (lblFilePath.Text.EndsWith("_S.txt"))
            {
                MessageBox.Show("Can't post lab sheet that has already been sent or has the status of sent [ends with _S.txt].", "Error");
                return;
            }

            if (csspWQInputApp.IncludeLaboratoryQAQC)
            {
                if (dateTimePickerSalinitiesReadDate.Value == null)
                {
                    MessageBox.Show("Salinities read date is invalid", "Unable to send information to server.         ", MessageBoxButtons.OK);
                    return;
                }
                if (dateTimePickerResultsReadDate.Value == null)
                {
                    MessageBox.Show("Results read date is invalid", "Unable to send information to server.         ", MessageBoxButtons.OK);
                    return;
                }
                if (dateTimePickerResultsRecordedDate.Value == null)
                {
                    MessageBox.Show("Results recorded date is invalid", "Unable to send information to server.         ", MessageBoxButtons.OK);
                    return;
                }
                if (dateTimePickerRun.Value == null)
                {
                    MessageBox.Show("Run date is invalid", "Unable to send information to server.", MessageBoxButtons.OK);
                    return;
                }

                if (dateTimePickerResultsReadDate.Value != null && dateTimePickerRun.Value != null)
                {
                    if (dateTimePickerResultsReadDate.Value != dateTimePickerRun.Value.AddDays(1 + (checkBoxIncubationStartSameDay.Checked == true ? 0 : 1)))
                    {
                        MessageBox.Show("Results read date is incorrect. \r\n\r\n Please correct date before sending to server", "Unable to send information to server.         ", MessageBoxButtons.OK);
                        return;
                    }
                }

                if (dateTimePickerSalinitiesReadDate.Value != null)
                {
                    if (dateTimePickerSalinitiesReadDate.Value > new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 23, 59, 59))
                    {
                        MessageBox.Show("Salinities read date is after today.\r\n\r\nCan't send future results to the server.", "Unable to send information to server.         ", MessageBoxButtons.OK);
                        return;
                    }
                }

                if (dateTimePickerResultsReadDate.Value != null)
                {
                    if (dateTimePickerResultsReadDate.Value > new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 23, 59, 59))
                    {
                        MessageBox.Show("Results read date is after today.\r\n\r\nCan't send future results to the server.", "Unable to send information to server.         ", MessageBoxButtons.OK);
                        return;
                    }
                }

                if (dateTimePickerResultsRecordedDate.Value != null)
                {
                    if (dateTimePickerResultsRecordedDate.Value > new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 23, 59, 59))
                    {
                        MessageBox.Show("Results recorded date is after today.\r\n\r\nCan't send future results to the server.", "Unable to send information to server.         ", MessageBoxButtons.OK);
                        return;
                    }
                }
            }

            FillInternetConnectionVariable();
            if (!InternetConnection)
            {
                butSendToServer.Text = "No Internet Connection";
                butSendToServer.Enabled = false;
                butGetLabSheetsStatus.Enabled = false;
                MessageBox.Show("No internet connection", "Internet connection");
                return;
            }

            //if (!EverythingEntered())
            //{
            //    return;
            //}
            butSendToServer.Text = "Working ...";
            butSendToServer.Enabled = false;
            butGetLabSheetsStatus.Enabled = false;
            lblStatus.Text = "Sending lab sheet to server ... Working ...";
            lblStatus.Refresh();
            Application.DoEvents();
            string retStr = PostLabSheet();
            if (string.IsNullOrWhiteSpace(retStr))
            {
                butSendToServer.Text = "Lab sheet sent ok";
                butSendToServer.Enabled = false;
                butGetLabSheetsStatus.Enabled = true;
                //if (comboBoxSubsectorNames.SelectedIndex != 0)
                //{
                //    butGetLabSheetsStatus.Enabled = true;
                //}
                lblStatus.Text = "Lab sheet sent ok";

                File.Copy(lblFilePath.Text, lblFilePath.Text.Replace("_C.txt", "_S.txt"));
                File.Delete(lblFilePath.Text);
                lblFilePath.Text = lblFilePath.Text.Replace("_C.txt", "_S.txt");
            }
            else
            {
                butSendToServer.Text = "Error sending lab sheet";
                if (InternetConnection)
                {
                    butSendToServer.Enabled = true;
                    butGetLabSheetsStatus.Enabled = true;
                    //if (comboBoxSubsectorNames.SelectedIndex != 0)
                    //{
                    //    butGetLabSheetsStatus.Enabled = true;
                    //}
                }
                lblStatus.Text = retStr;
            }
        }
        private void SetButGetTidesEnabledOrNot()
        {
            int TimeColumn = 2;
            bool ShowTideButton = false;
            for (int row = 0, countRow = dataGridViewCSSP.Rows.Count; row < countRow; row++)
            {
                if (!string.IsNullOrWhiteSpace(dataGridViewCSSP[TimeColumn, row].Value.ToString()))
                {
                    if (dataGridViewCSSP[TimeColumn, row].Value.ToString().Length == 5)
                    {
                        if (dataGridViewCSSP[TimeColumn, row].Value.ToString().Substring(2, 1) == ":")
                        {
                            ShowTideButton = true;
                        }
                    }
                }
            }

            butGetTides.Enabled = ShowTideButton;
        }
        private void SetupAppInputFiles()
        {
            listBoxFiles.Focus();
            panelAppInputFiles.BringToFront();
            panelSendToServerCompare.SendToBack();
            comboBoxFileSubsector.Items.Clear();

            DirectoryInfo di = new DirectoryInfo(CurrentPath);

            if (!di.Exists)
            {
                lblStatus.Text = "Could not find directory [" + CurrentPath + "]";
                return;
            }

            comboBoxFileSubsector.Items.Add("All");
            for (int i = 1, count = comboBoxSubsectorNames.Items.Count; i < count; i++)
            {
                string Name = ((FileItem)comboBoxSubsectorNames.Items[i]).Name;
                comboBoxFileSubsector.Items.Add(Name);
            }

            comboBoxFileSubsector.SelectedIndex = 0;

            CurrentPanel = panelAppInputFiles;
            panelAppInputIsVisible = false;
            butSendToServer.Enabled = false;
            if (InternetConnection)
            {
                butGetLabSheetsStatus.Enabled = true;
            }
        }
        private void SetupCSSPWQInputTool()
        {
            CreateCSSPSamplingPlanFilePath();
            textBoxAccessCode.Text = "";
            panelApp.BringToFront();
            CurrentPanel = panelApp;
            panelButtonBar.Visible = true;
            FillInternetConnectionVariable();
            lblFilePath.Text = "";
            butCreateFile.Visible = false;
            FillComboboxes();
            dataGridViewCellStyleEdit.BackColor = Color.White;
            dataGridViewCellStyleEdit.ForeColor = Color.Green;
            dataGridViewCellStyleEditError.BackColor = Color.Red;
            dataGridViewCellStyleEditError.ForeColor = Color.Black;
            dataGridViewCellStyleEditRowCell.BackColor = Color.Orange;
            dataGridViewCellStyleEditRowCell.ForeColor = Color.Green;
        }
        private void SetupDataGridViewCSSP()
        {
            dataGridViewCSSP.Rows.Clear();

            if (csspWQInputSheetType == CSSPWQInputSheetTypeEnum.LTB)
            {
                SetupDataGridViewCSSPLTB();
            }
            else if (csspWQInputSheetType == CSSPWQInputSheetTypeEnum.EC)
            {
                SetupDataGridViewCSSPEC();
            }
            else
            {
                SetupDataGridViewCSSPA1(new LabSheetA1Sheet());
            }
            dataGridViewCellStyleDefault = dataGridViewCSSP.DefaultCellStyle;

        }
        private void SetupDataGridViewCSSPA1(LabSheetA1Sheet labSheetA1Sheet)
        {
            dataGridViewCSSP.ColumnCount = 0;
            dataGridViewCSSP.ColumnHeadersDefaultCellStyle.BackColor = Color.Navy;
            dataGridViewCSSP.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridViewCSSP.ColumnHeadersDefaultCellStyle.Font = new Font(dataGridViewCSSP.Font, FontStyle.Bold);

            dataGridViewCSSP.Name = "dataGridViewCSSP";
            dataGridViewCSSP.Location = new Point(8, 8);
            dataGridViewCSSP.Size = new Size(500, 250);
            dataGridViewCSSP.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCellsExceptHeaders;
            dataGridViewCSSP.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
            dataGridViewCSSP.CellBorderStyle = DataGridViewCellBorderStyle.Single;
            dataGridViewCSSP.GridColor = Color.Black;
            dataGridViewCSSP.RowHeadersVisible = false;

            // MWQM Site, Sample Type, MPN Column 0
            DataGridViewTextBoxColumn dgvc = new DataGridViewTextBoxColumn();
            dgvc.HeaderText = "Site - Type                                 MPN";
            dgvc.Width = 260;
            dgvc.ReadOnly = true;
            DataGridViewCellStyle dataGridViewCellStyle = new DataGridViewCellStyle();
            dataGridViewCellStyle.BackColor = Color.LightGray;
            dgvc.DefaultCellStyle = dataGridViewCellStyle;
            dataGridViewCSSP.Columns.Add(dgvc);

            // MWQM Site Column 1
            dgvc = new DataGridViewTextBoxColumn();
            dgvc.HeaderText = "MWQM Site";
            dgvc.Width = 1;
            dgvc.ReadOnly = true;
            dgvc.Visible = false;
            dataGridViewCSSP.Columns.Add(dgvc);

            // Time Column 2
            dgvc = new DataGridViewTextBoxColumn();
            dgvc.HeaderText = "Time";
            dgvc.Width = 60;
            dgvc.MaxInputLength = 5;
            dataGridViewCellStyle = new DataGridViewCellStyle();
            dataGridViewCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvc.DefaultCellStyle = dataGridViewCellStyle;
            dataGridViewCSSP.Columns.Add(dgvc);

            // MPN / 100 ml Fecal Coliform Column 3
            dgvc = new DataGridViewTextBoxColumn();
            dgvc.HeaderText = "MPN/100ml Fecal Coliform";
            dgvc.Width = 1;
            dgvc.ReadOnly = true;
            dgvc.Visible = false;
            dataGridViewCSSP.Columns.Add(dgvc);

            // Positive Tubes 10 Column 4
            dgvc = new DataGridViewTextBoxColumn();
            dgvc.HeaderText = "Pos. Tubes 10";
            dgvc.Width = 50;
            dgvc.MaxInputLength = 1;
            dataGridViewCellStyle = new DataGridViewCellStyle();
            dataGridViewCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCSSP.Columns.Add(dgvc);

            // Positive Tubes 1.0 Column 5
            dgvc = new DataGridViewTextBoxColumn();
            dgvc.HeaderText = "Pos. Tubes 1.0";
            dgvc.Width = 50;
            dgvc.MaxInputLength = 1;
            dataGridViewCellStyle = new DataGridViewCellStyle();
            dataGridViewCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCSSP.Columns.Add(dgvc);

            // Positive Tubes 0.1 Column 6
            dgvc = new DataGridViewTextBoxColumn();
            dgvc.HeaderText = "Pos. Tubes 0.1";
            dgvc.Width = 50;
            dgvc.MaxInputLength = 1;
            dataGridViewCellStyle = new DataGridViewCellStyle();
            dataGridViewCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCSSP.Columns.Add(dgvc);

            // Salinity (ppt) Column 7
            dgvc = new DataGridViewTextBoxColumn();
            dgvc.HeaderText = "Salinity (ppt)";
            dgvc.Name = "salinity";
            dgvc.Width = 60;
            dgvc.MaxInputLength = 4;
            dataGridViewCellStyle = new DataGridViewCellStyle();
            dataGridViewCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCSSP.Columns.Add(dgvc);

            // Temp (ºC) Column 8
            dgvc = new DataGridViewTextBoxColumn();
            dgvc.HeaderText = "Temp (ºC)";
            dgvc.Width = 60;
            dgvc.MaxInputLength = 4;
            dataGridViewCellStyle = new DataGridViewCellStyle();
            dataGridViewCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCSSP.Columns.Add(dgvc);

            // Processed by Column 9
            dgvc = new DataGridViewTextBoxColumn();
            dgvc.HeaderText = "Processed by";
            dgvc.Width = 70;
            dgvc.MaxInputLength = 20;
            dgvc.Visible = csspWQInputApp.IncludeLaboratoryQAQC;
            dataGridViewCSSP.Columns.Add(dgvc);

            // Sampling Type Column 10
            dgvc = new DataGridViewTextBoxColumn();
            dgvc.HeaderText = "Sample Type";
            dgvc.Width = 1;
            dgvc.ReadOnly = true;
            dgvc.Visible = false;
            dataGridViewCSSP.Columns.Add(dgvc);

            // TVItemID Column 11
            dgvc = new DataGridViewTextBoxColumn();
            dgvc.HeaderText = "ID";
            dgvc.Width = 1;
            dgvc.ReadOnly = true;
            dgvc.Visible = false;
            dataGridViewCSSP.Columns.Add(dgvc);

            int TotWidth = 0;
            for (int i = 0; i < 12; i++)
            {
                TotWidth += dataGridViewCSSP.Columns[i].Width;
            }

            // Comment Column 12
            dgvc = new DataGridViewTextBoxColumn();
            dgvc.HeaderText = "Comment";
            dgvc.Width = dataGridViewCSSP.Width - TotWidth - 5;
            //dgvc.MaxInputLength = 200;
            dataGridViewCSSP.Columns.Add(dgvc);

            if (labSheetA1Sheet.LabSheetA1MeasurementList.Count > 0)
            {
                foreach (LabSheetA1Measurement labSheetA1Measurment in labSheetA1Sheet.LabSheetA1MeasurementList)
                {
                    string MPNText = "";
                    if (labSheetA1Measurment.MPN != null)
                    {
                        switch (labSheetA1Measurment.MPN)
                        {
                            case 1:
                                {
                                    MPNText = "< 2";
                                }
                                break;
                            case 1700:
                                {
                                    MPNText = "> 1600";
                                }
                                break;
                            default:
                                {
                                    MPNText = labSheetA1Measurment.MPN.ToString();
                                }
                                break;
                        }
                    }

                    string AfterSampleTypeSpace = GetAfterSampleTypeSpace(labSheetA1Measurment.SampleType.ToString());

                    object[] row = { labSheetA1Measurment.Site + " - " + labSheetA1Measurment.SampleType + AfterSampleTypeSpace +
                            SpaceStr.Substring(0, SpaceStr.Length - MPNText.Length) + MPNText,
                            labSheetA1Measurment.Site,
                        (labSheetA1Measurment.Time == null ? "" : ((DateTime)labSheetA1Measurment.Time).ToString("HH:mm")),
                        labSheetA1Measurment.MPN.ToString(),
                        (labSheetA1Measurment.Tube10 == null ? "" : labSheetA1Measurment.Tube10.ToString()),
                        (labSheetA1Measurment.Tube1_0 == null ? "" : labSheetA1Measurment.Tube1_0.ToString()),
                        (labSheetA1Measurment.Tube0_1 == null ? "" : labSheetA1Measurment.Tube0_1.ToString()),
                        (labSheetA1Measurment.Salinity == null ? "" : ((float)labSheetA1Measurment.Salinity).ToString("F1")),
                        (labSheetA1Measurment.Temperature == null ? "" : ((float)labSheetA1Measurment.Temperature).ToString("F1")),
                        (labSheetA1Measurment.ProcessedBy == null ? "" : labSheetA1Measurment.ProcessedBy),
                        (labSheetA1Measurment.SampleType == null ? "" : labSheetA1Measurment.SampleType.ToString()),
                        labSheetA1Measurment.TVItemID.ToString(),
                        labSheetA1Measurment.SiteComment };
                    dataGridViewCSSP.Rows.Add(row);
                }
            }
            else
            {
                for (int i = 0, count = CSSPWQInputParamCurrent.MWQMSiteList.Count; i < count; i++)
                {
                    object[] row = { CSSPWQInputParamCurrent.MWQMSiteList[i] + " - " + lblSampleType.Text, CSSPWQInputParamCurrent.MWQMSiteList[i], "", "", "", "", "", "", "", "", lblSampleType.Text, CSSPWQInputParamCurrent.MWQMSiteTVItemIDList[i], "" };
                    dataGridViewCSSP.Rows.Add(row);
                }

                for (int i = 0, count = CSSPWQInputParamCurrent.DailyDuplicateMWQMSiteList.Count; i < count; i++)
                {
                    object[] row = { CSSPWQInputParamCurrent.DailyDuplicateMWQMSiteList[i] + " - " + SampleTypeEnum.DailyDuplicate.ToString(), CSSPWQInputParamCurrent.DailyDuplicateMWQMSiteList[i], "", "", "", "", "", "", "", "", SampleTypeEnum.DailyDuplicate.ToString(), CSSPWQInputParamCurrent.DailyDuplicateMWQMSiteTVItemIDList[i], "" };
                    dataGridViewCSSP.Rows.Add(row);
                }
            }

            for (int i = 0, ColCount = dataGridViewCSSP.Columns.Count; i < ColCount; i++)
            {
                dataGridViewCSSP.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            }

        }
        private void SetupDataGridViewCSSPEC()
        {
            dataGridViewCSSP.ColumnCount = 0;
            dataGridViewCSSP.ColumnHeadersDefaultCellStyle.BackColor = Color.Navy;
            dataGridViewCSSP.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridViewCSSP.ColumnHeadersDefaultCellStyle.Font = new Font(dataGridViewCSSP.Font, FontStyle.Bold);

            dataGridViewCSSP.Name = "dataGridViewCSSP";
            dataGridViewCSSP.Location = new Point(8, 8);
            dataGridViewCSSP.Size = new Size(500, 250);
            dataGridViewCSSP.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCellsExceptHeaders;
            dataGridViewCSSP.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
            dataGridViewCSSP.CellBorderStyle = DataGridViewCellBorderStyle.Single;
            dataGridViewCSSP.GridColor = Color.Black;
            dataGridViewCSSP.RowHeadersVisible = false;

            // Site Column 0
            DataGridViewTextBoxColumn dgvc = new DataGridViewTextBoxColumn();
            dgvc.HeaderText = "EC Lab Sheet Not Implemented";
            dgvc.Name = "NI";
            dgvc.Width = 500;
            dgvc.ReadOnly = true;
            dataGridViewCSSP.Columns.Add(dgvc);


        }
        private void SetupDataGridViewCSSPLTB()
        {
            dataGridViewCSSP.ColumnCount = 0;
            dataGridViewCSSP.ColumnHeadersDefaultCellStyle.BackColor = Color.Navy;
            dataGridViewCSSP.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridViewCSSP.ColumnHeadersDefaultCellStyle.Font = new Font(dataGridViewCSSP.Font, FontStyle.Bold);

            dataGridViewCSSP.Name = "dataGridViewCSSP";
            dataGridViewCSSP.Location = new Point(8, 8);
            dataGridViewCSSP.Size = new Size(500, 250);
            dataGridViewCSSP.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCellsExceptHeaders;
            dataGridViewCSSP.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
            dataGridViewCSSP.CellBorderStyle = DataGridViewCellBorderStyle.Single;
            dataGridViewCSSP.GridColor = Color.Black;
            dataGridViewCSSP.RowHeadersVisible = false;

            // Site Column 0
            DataGridViewTextBoxColumn dgvc = new DataGridViewTextBoxColumn();
            dgvc.HeaderText = "LTB Lab Sheet Not Implemented";
            dgvc.Name = "NI";
            dgvc.Width = 500;
            dgvc.ReadOnly = true;
            dataGridViewCSSP.Columns.Add(dgvc);


        }
        private void SyncArchives(DirectoryInfo di, DirectoryInfo diShared)
        {
            lblStatus.Text = "Doing local To Shared Archived Directory";
            richTextBoxLabSheetSender.AppendText("Doing local To Shared Archived Directory \r\n");

            lblReceiverLocalOrServer.Text = Server;
            lblSenderLocalOrServer.Text = Local;

            string retStr = SyncArchiveFromLocal(di, diShared);
            if (!string.IsNullOrWhiteSpace(retStr))
                return;

            lblStatus.Text = "Doing Shared Archived Directory To local";
            richTextBoxLabSheetSender.AppendText("Doing Shared Archived Directory To local \r\n");

            lblReceiverLocalOrServer.Text = Local;
            lblSenderLocalOrServer.Text = Server;

            retStr = SyncArchiveFromSharedArchives(di, diShared);
            if (!string.IsNullOrWhiteSpace(retStr))
                return;

            SetupAppInputFiles();
        }
        private string SyncArchiveFromLocal(DirectoryInfo di, DirectoryInfo diShared)
        {
            lblSenderLocalOrServer.Text = Local;
            lblReceiverLocalOrServer.Text = Server;

            if (!di.Exists)
            {
                lblStatus.Text = "Could not find directory [" + di.FullName + "]";
                richTextBoxLabSheetSender.AppendText("Could not find directory [" + di.FullName + "]\r\n");
                return lblStatus.Text;
            }

            if (!diShared.Exists)
            {
                lblStatus.Text = "Could not find directory [" + diShared.FullName + "]";
                richTextBoxLabSheetSender.AppendText("Could not find directory [" + diShared.FullName + "]\r\n");
                return lblStatus.Text;
            }

            List<FileInfo> fileList = (from c in di.GetFiles() select c).ToList();
            foreach (FileInfo fi in fileList)
            {
                butFileArchiveCancel.Enabled = false;
                butFileArchiveCopy.Enabled = false;
                butFileArchiveSkip.Enabled = false;

                lblSendingFileName.Text = fi.Name;
                lblSendingFileName.Refresh();
                Application.DoEvents();

                lblStatus.Text = "Checking file [" + fi.FullName + "]";
                lblStatus.Refresh();
                Application.DoEvents();

                richTextBoxLabSheetSender.AppendText("Checking - " + fi.FullName + "\r\n");

                string retStr = "";
                string ArchiveFileName = textBoxSharedArchivedDirectory.Text + fi.FullName.Replace(RootCurrentPath, "");
                FileInfo fiArchive = new FileInfo(ArchiveFileName);
                if (fi.FullName.EndsWith(".txt"))
                {
                    retStr = MakeSureLabSheetFilesIsUniqueTxt(fi);
                    if (!string.IsNullOrWhiteSpace(retStr))
                    {
                        lblStatus.Text = "Error: Lab Sheet File for specific date, subsector and sampling plan file should be unique. More than one was found " + retStr;
                        richTextBoxLabSheetSender.AppendText(lblStatus.Text + "\r\n");
                        return lblStatus.Text;
                    }

                    retStr = MakeSureLabSheetFilesIsUniqueTxt(fiArchive);
                    if (!string.IsNullOrWhiteSpace(retStr))
                    {
                        lblStatus.Text = "Error: Lab Sheet File for specific date, subsector and sampling plan file should be unique. More than one was found " + retStr;
                        richTextBoxLabSheetSender.AppendText(lblStatus.Text + "\r\n");
                        return lblStatus.Text;
                    }

                    retStr = CheckDestinationFilesTxt(fi, fiArchive);
                }
                else if (fi.FullName.EndsWith(".docx"))
                {
                    retStr = MakeSureLabSheetFilesIsUniqueDocx(fi);
                    if (!string.IsNullOrWhiteSpace(retStr))
                    {
                        lblStatus.Text = "Error: FC Form for Lab Sheet File for specific date, subsector and sampling plan file should be unique. More than one was found " + retStr;
                        richTextBoxLabSheetSender.AppendText(lblStatus.Text + "\r\n");
                        return lblStatus.Text;
                    }

                    retStr = MakeSureLabSheetFilesIsUniqueDocx(fiArchive);
                    if (!string.IsNullOrWhiteSpace(retStr))
                    {
                        lblStatus.Text = "Error: FC Form for Lab Sheet File for specific date, subsector and sampling plan file should be unique. More than one was found " + retStr;
                        richTextBoxLabSheetSender.AppendText(lblStatus.Text + "\r\n");
                        return lblStatus.Text;
                    }

                    retStr = CheckDestinationFilesDocx(fi, fiArchive);
                }

                if (string.IsNullOrWhiteSpace(retStr))
                {
                    if (!fiArchive.Exists)
                    {
                        lblStatus.Text = "Copying - " + fi.FullName + " to " + fiArchive.FullName;
                        richTextBoxLabSheetSender.AppendText("Copying - " + fi.FullName + "\r\n to " + fiArchive.FullName + "\r\n");
                        try
                        {
                            File.Copy(fi.FullName, fiArchive.FullName);
                        }
                        catch (Exception ex)
                        {
                            lblStatus.Text = "Error:" + ex.Message + (ex.InnerException != null ? ex.InnerException.Message : "");
                            richTextBoxLabSheetSender.AppendText("Error:" + ex.Message + (ex.InnerException != null ? ex.InnerException.Message : "") + "\r\n");
                            return "Error:" + ex.Message + (ex.InnerException != null ? ex.InnerException.Message : "");
                        }
                        lblStatus.Text = "Copied - " + fi.FullName + "\r\n to " + fiArchive.FullName;
                        richTextBoxLabSheetSender.AppendText("Copied - " + fi.FullName + "\r\n to " + fiArchive.FullName + "\r\n");
                    }
                    else
                    {
                        // should check the last modified date
                        if (fi.LastWriteTimeUtc > fiArchive.LastWriteTimeUtc)
                        {
                            bool FromLocal = true;
                            ReplaceFileFromTo(fi, fiArchive, FromLocal);
                            if (UserActionFileArchiveCancel)
                                return "Cancelling ... ";
                        }
                    }
                    fiArchive = new FileInfo(ArchiveFileName);
                    if (!fiArchive.Exists)
                    {
                        lblStatus.Text = "Did not copy file [" + diShared.FullName + "]";
                        richTextBoxLabSheetSender.AppendText("Did not copy file [" + diShared.FullName + "]");
                        return lblStatus.Text;
                    }
                }
            }

            List<DirectoryInfo> dirInfoList = (from c in di.GetDirectories() select c).ToList();
            foreach (DirectoryInfo diNext in dirInfoList)
            {
                string ArchiveDirectory = textBoxSharedArchivedDirectory.Text + diNext.FullName.Replace(RootCurrentPath, "");
                DirectoryInfo diArchive = new DirectoryInfo(ArchiveDirectory);
                if (!diArchive.Exists)
                {
                    try
                    {
                        diArchive.Create();
                        diArchive = new DirectoryInfo(ArchiveDirectory);
                    }
                    catch (Exception ex)
                    {
                        lblStatus.Text = "Error:" + ex.Message + (ex.InnerException != null ? ex.InnerException.Message : "");
                        richTextBoxLabSheetSender.AppendText("Error:" + ex.Message + (ex.InnerException != null ? ex.InnerException.Message : "") + "\r\n");
                        return "Error:" + ex.Message + (ex.InnerException != null ? ex.InnerException.Message : "");
                    }
                }

                string retStr = SyncArchiveFromLocal(diNext, diArchive);
                if (!string.IsNullOrWhiteSpace(retStr))
                {
                    return "Error";
                }
            }

            return "";
        }
        private string SyncArchiveFromSharedArchives(DirectoryInfo di, DirectoryInfo diShared)
        {
            lblSenderLocalOrServer.Text = Server;
            lblReceiverLocalOrServer.Text = Local;

            if (!di.Exists)
            {
                lblStatus.Text = "Could not find directory [" + di.FullName + "]";
                richTextBoxLabSheetSender.AppendText("Could not find directory [" + di.FullName + "]\r\n");
                return lblStatus.Text;
            }

            if (!diShared.Exists)
            {
                lblStatus.Text = "Could not find directory [" + diShared.FullName + "]";
                richTextBoxLabSheetSender.AppendText("Could not find directory [" + diShared.FullName + "]\r\n");
                return lblStatus.Text;
            }

            List<FileInfo> fileList = (from c in diShared.GetFiles() select c).ToList();
            foreach (FileInfo fi in fileList)
            {
                lblSendingFileName.Text = fi.Name;
                lblSendingFileName.Refresh();
                Application.DoEvents();

                lblStatus.Text = "Checking file [" + fi.FullName + "]";
                lblStatus.Refresh();
                Application.DoEvents();

                richTextBoxLabSheetSender.AppendText("Checking - " + fi.FullName + "\r\n");

                string retStr = "";
                string LocalFileName = RootCurrentPath + fi.FullName.Replace(textBoxSharedArchivedDirectory.Text, "");
                FileInfo fiLocal = new FileInfo(LocalFileName);
                if (fi.FullName.EndsWith(".txt"))
                {
                    retStr = MakeSureLabSheetFilesIsUniqueTxt(fi);
                    if (!string.IsNullOrWhiteSpace(retStr))
                    {
                        lblStatus.Text = "Error: Lab Sheet File for specific date, subsector and sampling plan file should be unique. More than one was found [" + retStr + "]";
                        richTextBoxLabSheetSender.AppendText(lblStatus.Text + "\r\n");
                        return lblStatus.Text;
                    }

                    retStr = MakeSureLabSheetFilesIsUniqueTxt(fiLocal);
                    if (!string.IsNullOrWhiteSpace(retStr))
                    {
                        lblStatus.Text = "Error: Lab Sheet File for specific date, subsector and sampling plan file should be unique. More than one was found [" + retStr + "]";
                        richTextBoxLabSheetSender.AppendText(lblStatus.Text + "\r\n");
                        return lblStatus.Text;
                    }

                    retStr = CheckDestinationFilesTxt(fi, fiLocal);
                }
                else if (fi.FullName.EndsWith(".docx"))
                {
                    retStr = MakeSureLabSheetFilesIsUniqueDocx(fi);
                    if (!string.IsNullOrWhiteSpace(retStr))
                    {
                        lblStatus.Text = "Error: Lab Sheet File for specific date, subsector and sampling plan file should be unique. More than one was found [" + retStr + "]";
                        richTextBoxLabSheetSender.AppendText(lblStatus.Text + "\r\n");
                        return lblStatus.Text;
                    }

                    retStr = MakeSureLabSheetFilesIsUniqueDocx(fiLocal);
                    if (!string.IsNullOrWhiteSpace(retStr))
                    {
                        lblStatus.Text = "Error: Lab Sheet File for specific date, subsector and sampling plan file should be unique. More than one was found [" + retStr + "]";
                        richTextBoxLabSheetSender.AppendText(lblStatus.Text + "\r\n");
                        return lblStatus.Text;
                    }

                    retStr = CheckDestinationFilesDocx(fi, fiLocal);
                }
                if (string.IsNullOrWhiteSpace(retStr))
                {

                    if (!fiLocal.Exists)
                    {
                        lblStatus.Text = "Copying - " + fi.FullName + " to " + fiLocal.FullName;
                        richTextBoxLabSheetSender.AppendText("Copying - " + fi.FullName + "\r\n to " + fiLocal.FullName + "\r\n");
                        try
                        {
                            File.Copy(fi.FullName, fiLocal.FullName);
                        }
                        catch (Exception ex)
                        {
                            lblStatus.Text = "Error:" + ex.Message + (ex.InnerException != null ? ex.InnerException.Message : "");
                            richTextBoxLabSheetSender.AppendText("Error:" + ex.Message + (ex.InnerException != null ? ex.InnerException.Message : "") + "\r\n");
                            return "Error:" + ex.Message + (ex.InnerException != null ? ex.InnerException.Message : "");
                        }
                        lblStatus.Text = "Copied - " + fi.FullName + " to " + fiLocal.FullName;
                        richTextBoxLabSheetSender.AppendText("Copied - " + fi.FullName + "\r\n to " + fiLocal.FullName + "\r\n");
                    }
                    else
                    {
                        // should check the last modified date
                        if (fi.LastWriteTimeUtc > fiLocal.LastWriteTimeUtc)
                        {
                            bool FromLocal = false;
                            ReplaceFileFromTo(fi, fiLocal, FromLocal);
                            if (UserActionFileArchiveCancel)
                                return "Cancelling ... ";
                        }
                    }
                    fiLocal = new FileInfo(LocalFileName);
                    if (!fiLocal.Exists)
                    {
                        lblStatus.Text = "Did not copy file [" + diShared.FullName + "]";
                        richTextBoxLabSheetSender.AppendText("Did not copy file [" + diShared.FullName + "]");
                        return lblStatus.Text;
                    }
                }
            }

            List<DirectoryInfo> dirInfoList = (from c in diShared.GetDirectories() select c).ToList();
            foreach (DirectoryInfo diSharedNext in dirInfoList)
            {
                string LocalDirectory = RootCurrentPath + diSharedNext.FullName.Replace(textBoxSharedArchivedDirectory.Text, "");
                DirectoryInfo diLocal = new DirectoryInfo(LocalDirectory);
                if (!diLocal.Exists)
                {
                    try
                    {
                        diLocal.Create();
                        diLocal = new DirectoryInfo(LocalDirectory);
                    }
                    catch (Exception ex)
                    {
                        lblStatus.Text = "Error:" + ex.Message + (ex.InnerException != null ? ex.InnerException.Message : "");
                        richTextBoxLabSheetSender.AppendText("Error:" + ex.Message + (ex.InnerException != null ? ex.InnerException.Message : "") + "\r\n");
                        return "Error:" + ex.Message + (ex.InnerException != null ? ex.InnerException.Message : "");
                    }
                }

                string retStr = SyncArchiveFromSharedArchives(di, diSharedNext);
                if (!string.IsNullOrWhiteSpace(retStr))
                {
                    return "Error";
                }
            }

            return "";
        }
        private void GetLabSheetsStatus()
        {
            List<FileInfo> fileListAll = new List<FileInfo>();
            DirectoryInfo di = new DirectoryInfo(CurrentPath);

            if (!di.Exists)
            {
                lblStatus.Text = "Could not find directory [" + CurrentPath + "]";
                return;
            }

            for (int i = 1980; i < DateTime.Now.Year + 1; i++)
            {
                di = new DirectoryInfo(CurrentPath + i + @"\");
                if (di.Exists)
                {
                    List<FileInfo> fileList = di.GetFiles().Where(c => c.FullName.Contains("_S.txt")).OrderBy(c => c.FullName).ToList();

                    foreach (FileInfo fi in fileList)
                    {
                        fileListAll.Add(fi);
                    }
                }
            }

            int count = 0;
            int countFile = fileListAll.Count;
            foreach (FileInfo fi in fileListAll)
            {
                count += 1;
                butGetLabSheetsStatus.Text = "Doing ... " + count + "/" + countFile;
                lblStatus.Text = "Checking status of LabSheet loaded [" + fi.Name + "] Doing ... " + count + "/" + countFile;
                lblStatus.Refresh();
                Application.DoEvents();
                string retStr = GetLabSheetExist(fi);
                if (retStr.Substring(0, 1) == "[")
                {
                    int TempInt = -10;
                    int.TryParse(retStr.Substring(1, retStr.Length - 2), out TempInt);
                    if (TempInt >= (int)LabSheetStatusEnum.Error && TempInt <= (int)LabSheetStatusEnum.Rejected)
                    {
                        switch ((LabSheetStatusEnum)TempInt)
                        {
                            case LabSheetStatusEnum.Error:
                            case LabSheetStatusEnum.Created:
                            case LabSheetStatusEnum.Transferred:
                                break;
                            case LabSheetStatusEnum.Accepted:
                                {
                                    File.Copy(fi.FullName, fi.FullName.Replace("_S.txt", "_A.txt"));
                                    FileInfo fiCopied = new FileInfo(fi.FullName.Replace("_S.txt", "_A.txt"));
                                    if (fiCopied.Exists)
                                    {
                                        File.Delete(fi.FullName);
                                    }
                                }
                                break;
                            case LabSheetStatusEnum.Rejected:
                                {
                                    File.Copy(fi.FullName, fi.FullName.Replace("_S.txt", "_R.txt"));
                                    FileInfo fiCopied = new FileInfo(fi.FullName.Replace("_S.txt", "_R.txt"));
                                    if (fiCopied.Exists)
                                    {
                                        File.Delete(fi.FullName);
                                    }
                                }
                                break;
                            default:
                                break;
                        }
                    }
                }
                else
                {
                    lblStatus.Text = retStr;
                    return;
                }
            }
        }
        private void ToggleFailFileName()
        {
            FileInfo fi = new FileInfo(lblFilePath.Text);
            if (!fi.FullName.EndsWith("_F.txt"))
            {
                DialogResult dialogResult = MessageBox.Show("Do you want to save a copy of the lab sheet under subdirectory Total Coliform?\r\n\r\n" +
                    "Only applicable if circulation did not start (i.e. results are total coliform)", "Create Total Coliform Lab Sheet", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    DirectoryInfo directoryInfo = fi.Directory;
                    directoryInfo = new DirectoryInfo(directoryInfo.FullName + @"\" + "Total Coliform");

                    if (!directoryInfo.Exists)
                    {
                        try
                        {
                            directoryInfo.Create();
                        }
                        catch (Exception ex)
                        {
                            lblStatus.Text = "Error creating directory [" + directoryInfo.FullName + @"\" + "]. " + ex.Message;
                        }
                    }

                    directoryInfo = new DirectoryInfo(directoryInfo.FullName);

                    FileInfo fiDest = new FileInfo(directoryInfo.FullName + @"\" + fi.Name.Substring(0, fi.Name.Length - 6) + "_F.txt");
                    if (fiDest.Exists)
                    {
                        try
                        {
                            fiDest.Delete();
                        }
                        catch (Exception ex)
                        {
                            lblStatus.Text = "Error deleting existing total coliform file [" + fiDest.FullName + "]. " + ex.Message;
                        }
                    }

                    try
                    {
                        File.Copy(fi.FullName, fiDest.FullName);
                    }
                    catch (Exception ex)
                    {
                        lblStatus.Text = "Error copying file [" + fi.Name + "] to Total Coliform directory. " + ex.Message;
                    }
                }
            }
            if (fi.FullName.Substring(fi.FullName.Length - 6) == "_F.txt")
            {
                File.Copy(fi.FullName, fi.FullName.Replace("_F.txt", "_C.txt"));
                fi.Delete();
            }
            else
            {
                File.Copy(fi.FullName, fi.FullName.Replace(fi.FullName.Substring(fi.FullName.Length - 6), "_F.txt"));
                fi.Delete();
            }
        }
        private void TrySendingToServer()
        {
            lblServerFileDateTime.Text = "";
            lblLocalFileDateTime.Text = "";
            butContinueSendToServer.Visible = true;
            butCancelSendToServer.Visible = true;
            butSendToServer.Enabled = false;
            butFileArchiveSkip.Visible = false;
            butFileArchiveCopy.Visible = false;
            butFileArchiveCancel.Visible = false;
            lblSenderLocalOrServer.Text = Local;
            lblReceiverLocalOrServer.Text = Server;
            panelSendToServerCompare.BringToFront();
            richTextBoxLabSheetSender.Text = "";
            richTextBoxLabSheetSender.ForeColor = Color.Black;
            richTextBoxLabSheetReceiver.Text = "";
            richTextBoxLabSheetReceiver.ForeColor = Color.Black;
            panelSendToServerCompare.Refresh();
            Application.DoEvents();
            richTextBoxLabSheetSender.Refresh();
            Application.DoEvents();
            richTextBoxLabSheetReceiver.Refresh();
            Application.DoEvents();

            lblSendingFileName.Text = labSheetA1Sheet.SubsectorName + "         (" + new DateTime(int.Parse(labSheetA1Sheet.RunYear), int.Parse(labSheetA1Sheet.RunMonth), int.Parse(labSheetA1Sheet.RunDay)).ToString("yyyy MMMM dd") + ")";
            lblSendingFileName.Refresh();
            Application.DoEvents();

            if (!lblFilePath.Text.EndsWith("_C.txt"))
            {
                MessageBox.Show("Only changed files i.e. ending with _C.txt can be sent to the server.");
                return;
            }

            //if (!EverythingEntered())
            //{
            //    return;
            //}

            string FileTextToSend = "";
            FileInfo fi = new FileInfo(lblFilePath.Text);
            if (!fi.Exists)
            {
                richTextBoxLabSheetSender.Text = "Error reading file [" + lblFilePath.Text + "]";
                return;
            }
            StreamReader sr = fi.OpenText();
            FileTextToSend = sr.ReadToEnd();
            sr.Close();

            richTextBoxLabSheetSender.Text = FileTextToSend;

            richTextBoxLabSheetReceiver.Text = "Checking if lab sheet has already been sent to the server ...";

            string ServerLabSheet = GetLabSheet(fi);
            if (string.IsNullOrWhiteSpace(ServerLabSheet))
            {
                richTextBoxLabSheetReceiver.Text = "Lab Sheet has not been sent to the server yet";
            }
            else
            {
                richTextBoxLabSheetReceiver.Text = ServerLabSheet;
            }
            CompareTheDocs();
        }
        private void TryToCalculateIncubationTimeSpan()
        {
            lblIncubationBath1TimeCalculated.ForeColor = Color.Black;
            if (textBoxIncubationBath1StartTime.ForeColor == Color.Red || textBoxIncubationBath1EndTime.ForeColor == Color.Red)
            {
                lblIncubationBath1TimeCalculated.ForeColor = Color.Red;
                lblIncubationBath1TimeCalculated.Text = "Error";
            }

            if (textBoxIncubationBath1StartTime.Text.Length == 5 && textBoxIncubationBath1StartTime.Text[2].ToString() == ":"
                && textBoxIncubationBath1EndTime.Text.Length == 5 && textBoxIncubationBath1EndTime.Text[2].ToString() == ":")
            {
                int StartHour = int.Parse(textBoxIncubationBath1StartTime.Text.Substring(0, 2));
                int StartMinute = int.Parse(textBoxIncubationBath1StartTime.Text.Substring(3, 2));
                int EndHour = int.Parse(textBoxIncubationBath1EndTime.Text.Substring(0, 2));
                int EndMinute = int.Parse(textBoxIncubationBath1EndTime.Text.Substring(3, 2));
                int SpanHour = 0;
                int SpanMinute = 0;
                SpanHour = 24 - StartHour + EndHour;
                SpanMinute = EndMinute - StartMinute;

                if (SpanMinute < 0)
                {
                    SpanHour = SpanHour - 1;
                    SpanMinute = 60 + SpanMinute;
                }

                lblIncubationBath1TimeCalculated.Text = (SpanHour < 10 ? "0" + SpanHour.ToString() : SpanHour.ToString()) + ":" +
                    (SpanMinute < 10 ? "0" + SpanMinute.ToString() : SpanMinute.ToString());
            }
            else
            {
                lblIncubationBath1TimeCalculated.ForeColor = Color.Red;
                lblIncubationBath1TimeCalculated.Text = "Error";
            }

            if (lblIncubationBath1TimeCalculated.Text.Contains("-"))
            {
                // nothing
            }
            else
            {
                if (lblIncubationBath1TimeCalculated.Text.Contains(":"))
                {
                    int Hour = 0;
                    int.TryParse(lblIncubationBath1TimeCalculated.Text.Substring(0, lblIncubationBath1TimeCalculated.Text.IndexOf(":")), out Hour);
                    if (Hour < 22 || Hour > 26)
                    {
                        lblIncubationBath1TimeCalculated.ForeColor = Color.Red;
                    }
                    if (Hour == 26)
                    {
                        int Min = 0;
                        int.TryParse(lblIncubationBath1TimeCalculated.Text.Substring(lblIncubationBath1TimeCalculated.Text.IndexOf(":") + 1), out Min);
                        if (Min > 0)
                        {
                            lblIncubationBath1TimeCalculated.ForeColor = Color.Red;
                        }
                    }
                }
            }

            if (!radioButton1Baths.Checked)
            {
                lblIncubationBath2TimeCalculated.ForeColor = Color.Black;
                if (textBoxIncubationBath2StartTime.ForeColor == Color.Red || textBoxIncubationBath2EndTime.ForeColor == Color.Red)
                {
                    lblIncubationBath2TimeCalculated.ForeColor = Color.Red;
                    lblIncubationBath2TimeCalculated.Text = "Error";
                }

                if (textBoxIncubationBath2StartTime.Text.Length == 5 && textBoxIncubationBath2StartTime.Text[2].ToString() == ":"
                    && textBoxIncubationBath2EndTime.Text.Length == 5 && textBoxIncubationBath2EndTime.Text[2].ToString() == ":")
                {
                    int StartHour = int.Parse(textBoxIncubationBath2StartTime.Text.Substring(0, 2));
                    int StartMinute = int.Parse(textBoxIncubationBath2StartTime.Text.Substring(3, 2));
                    int EndHour = int.Parse(textBoxIncubationBath2EndTime.Text.Substring(0, 2));
                    int EndMinute = int.Parse(textBoxIncubationBath2EndTime.Text.Substring(3, 2));
                    int SpanHour = 0;
                    int SpanMinute = 0;
                    SpanHour = 24 - StartHour + EndHour;
                    SpanMinute = EndMinute - StartMinute;

                    if (SpanMinute < 0)
                    {
                        SpanHour = SpanHour - 1;
                        SpanMinute = 60 + SpanMinute;
                    }

                    lblIncubationBath2TimeCalculated.Text = (SpanHour < 10 ? "0" + SpanHour.ToString() : SpanHour.ToString()) + ":" +
                        (SpanMinute < 10 ? "0" + SpanMinute.ToString() : SpanMinute.ToString());
                }
                else
                {
                    lblIncubationBath2TimeCalculated.ForeColor = Color.Red;
                    lblIncubationBath2TimeCalculated.Text = "Error";
                }

                if (lblIncubationBath2TimeCalculated.Text.Contains("-"))
                {
                    // nothing
                }
                else
                {
                    if (lblIncubationBath2TimeCalculated.Text.Contains(":"))
                    {
                        int Hour = 0;
                        int.TryParse(lblIncubationBath2TimeCalculated.Text.Substring(0, lblIncubationBath2TimeCalculated.Text.IndexOf(":")), out Hour);
                        if (Hour < 22 || Hour > 26)
                        {
                            lblIncubationBath2TimeCalculated.ForeColor = Color.Red;
                        }
                        if (Hour == 26)
                        {
                            int Min = 0;
                            int.TryParse(lblIncubationBath2TimeCalculated.Text.Substring(lblIncubationBath2TimeCalculated.Text.IndexOf(":") + 1), out Min);
                            if (Min > 0)
                            {
                                lblIncubationBath2TimeCalculated.ForeColor = Color.Red;
                            }
                        }
                    }
                }
            }

            if (!(radioButton1Baths.Checked || radioButton2Baths.Checked))
            {
                lblIncubationBath3TimeCalculated.ForeColor = Color.Black;
                if (textBoxIncubationBath3StartTime.ForeColor == Color.Red || textBoxIncubationBath3EndTime.ForeColor == Color.Red)
                {
                    lblIncubationBath3TimeCalculated.ForeColor = Color.Red;
                    lblIncubationBath3TimeCalculated.Text = "Error";
                }

                if (textBoxIncubationBath3StartTime.Text.Length == 5 && textBoxIncubationBath3StartTime.Text[2].ToString() == ":"
                    && textBoxIncubationBath3EndTime.Text.Length == 5 && textBoxIncubationBath3EndTime.Text[2].ToString() == ":")
                {
                    int StartHour = int.Parse(textBoxIncubationBath3StartTime.Text.Substring(0, 2));
                    int StartMinute = int.Parse(textBoxIncubationBath3StartTime.Text.Substring(3, 2));
                    int EndHour = int.Parse(textBoxIncubationBath3EndTime.Text.Substring(0, 2));
                    int EndMinute = int.Parse(textBoxIncubationBath3EndTime.Text.Substring(3, 2));
                    int SpanHour = 0;
                    int SpanMinute = 0;
                    SpanHour = 24 - StartHour + EndHour;
                    SpanMinute = EndMinute - StartMinute;

                    if (SpanMinute < 0)
                    {
                        SpanHour = SpanHour - 1;
                        SpanMinute = 60 + SpanMinute;
                    }

                    lblIncubationBath3TimeCalculated.Text = (SpanHour < 10 ? "0" + SpanHour.ToString() : SpanHour.ToString()) + ":" +
                        (SpanMinute < 10 ? "0" + SpanMinute.ToString() : SpanMinute.ToString());
                }
                else
                {
                    lblIncubationBath3TimeCalculated.ForeColor = Color.Red;
                    lblIncubationBath3TimeCalculated.Text = "Error";
                }

                if (lblIncubationBath3TimeCalculated.Text.Contains("-"))
                {
                    // nothing
                }
                else
                {
                    if (lblIncubationBath3TimeCalculated.Text.Contains(":"))
                    {
                        int Hour = 0;
                        int.TryParse(lblIncubationBath3TimeCalculated.Text.Substring(0, lblIncubationBath3TimeCalculated.Text.IndexOf(":")), out Hour);
                        if (Hour < 22 || Hour > 26)
                        {
                            lblIncubationBath3TimeCalculated.ForeColor = Color.Red;
                        }
                        if (Hour == 26)
                        {
                            int Min = 0;
                            int.TryParse(lblIncubationBath3TimeCalculated.Text.Substring(lblIncubationBath3TimeCalculated.Text.IndexOf(":") + 1), out Min);
                            if (Min > 0)
                            {
                                lblIncubationBath3TimeCalculated.ForeColor = Color.Red;
                            }
                        }
                    }
                }
            }
        }
        private void TryToCalculateMPNA1(int RowIndex)
        {
            int MPNColumn = 3;
            int FirstTubeColumn = 4;
            dataGridViewCSSP[MPNColumn, RowIndex].Style.ForeColor = Color.Black;
            dataGridViewCSSP[MPNColumn, RowIndex].Value = "";
            if (dataGridViewCSSP[FirstTubeColumn, RowIndex].Value == null
                || dataGridViewCSSP[FirstTubeColumn + 1, RowIndex].Value == null
                || dataGridViewCSSP[FirstTubeColumn + 2, RowIndex].Value == null)
                return;

            if (string.IsNullOrWhiteSpace(dataGridViewCSSP[FirstTubeColumn, RowIndex].Value.ToString())
                || string.IsNullOrWhiteSpace(dataGridViewCSSP[FirstTubeColumn + 1, RowIndex].Value.ToString())
                || string.IsNullOrWhiteSpace(dataGridViewCSSP[FirstTubeColumn + 2, RowIndex].Value.ToString()))
                return;

            int Tube10 = -1;
            int Tube1_0 = -1;
            int Tube0_1 = -1;

            int.TryParse(dataGridViewCSSP[FirstTubeColumn, RowIndex].Value.ToString(), out Tube10);
            int.TryParse(dataGridViewCSSP[FirstTubeColumn + 1, RowIndex].Value.ToString(), out Tube1_0);
            int.TryParse(dataGridViewCSSP[FirstTubeColumn + 2, RowIndex].Value.ToString(), out Tube0_1);

            CSSPMPNTable csspMPNTable = (from c in csspMPNTableList
                                         where c.Tube10 == Tube10
                                         && c.Tube1_0 == Tube1_0
                                         && c.Tube0_1 == Tube0_1
                                         select c).FirstOrDefault();

            if (csspMPNTable == null)
            {
                dataGridViewCSSP[MPNColumn, RowIndex].Style.ForeColor = Color.Red;
                dataGridViewCSSP[MPNColumn, RowIndex].Value = "Error";
            }
            else
            {
                dataGridViewCSSP[MPNColumn, RowIndex].Value = csspMPNTable.MPN;
                lblDailyDuplicateAcceptableOrUnacceptable.ForeColor = Color.Black;
            }
            CalculateDuplicate();

            ValidateCellA1(3, RowIndex);
        }
        private void TryToSyncArchive()
        {
            lblServerFileDateTime.Text = "";
            lblLocalFileDateTime.Text = "";
            butContinueSendToServer.Visible = false;
            butCancelSendToServer.Visible = false;
            butFileArchiveSkip.Visible = true;
            butFileArchiveCopy.Visible = true;
            butFileArchiveCancel.Visible = true;
            panelSendToServerCompare.BringToFront();
            richTextBoxLabSheetSender.Text = "";
            richTextBoxLabSheetSender.ForeColor = Color.Black;
            richTextBoxLabSheetReceiver.Text = "";
            richTextBoxLabSheetReceiver.ForeColor = Color.Black;
            panelSendToServerCompare.Refresh();
            Application.DoEvents();
            richTextBoxLabSheetSender.Refresh();
            Application.DoEvents();
            richTextBoxLabSheetReceiver.Refresh();
            Application.DoEvents();

            FileInfo fiSamplingPlan = new FileInfo(SamplingPlanName);

            if (!fiSamplingPlan.Exists)
            {
                richTextBoxLabSheetSender.AppendText("Could not find file [" + fiSamplingPlan.FullName + "]\r\n");
                return;
            }

            lblSendingFileName.Text = fiSamplingPlan.Name;
            panelReceiverTop.BackColor = Color.LightGreen;
            panelSenderTop.BackColor = Color.LightBlue;

            lblSendingFileName.Refresh();
            Application.DoEvents();

            string ArchiveFileName = textBoxSharedArchivedDirectory.Text + fiSamplingPlan.FullName.Replace(RootCurrentPath, "");
            FileInfo fiArchive = new FileInfo(ArchiveFileName);

            if (!fiArchive.Exists)
            {
                lblStatus.Text = "Copying - " + fiSamplingPlan.FullName + " to " + fiArchive.FullName;
                richTextBoxLabSheetSender.AppendText("Copying - " + fiSamplingPlan.FullName + "\r\n to " + fiArchive.FullName + "\r\n");
                try
                {
                    File.Copy(fiSamplingPlan.FullName, fiArchive.FullName);
                }
                catch (Exception ex)
                {
                    lblStatus.Text = "Error: " + ex.Message + (ex.InnerException != null ? ex.InnerException.Message : "");
                    richTextBoxLabSheetSender.AppendText("Error:" + ex.Message + (ex.InnerException != null ? ex.InnerException.Message : "") + "\r\n");
                    MessageBox.Show(lblStatus.Text, "Error");
                    return;
                }
                richTextBoxLabSheetSender.AppendText("Copied - " + fiSamplingPlan.FullName + "\r\n to " + fiArchive.FullName + "\r\n");
            }
            else
            {
                if (fiSamplingPlan.LastWriteTimeUtc > fiArchive.LastWriteTimeUtc)
                {
                    lblSenderLocalOrServer.Text = Local;
                    lblReceiverLocalOrServer.Text = Server;
                    bool FromLocal = true;
                    ReplaceFileFromTo(fiSamplingPlan, fiArchive, FromLocal);
                    if (UserActionFileArchiveCancel)
                        return;
                }
                else if (fiSamplingPlan.LastWriteTimeUtc < fiArchive.LastWriteTimeUtc)
                {
                    lblSenderLocalOrServer.Text = Server;
                    lblReceiverLocalOrServer.Text = Local;
                    bool FromLocal = false;
                    ReplaceFileFromTo(fiArchive, fiSamplingPlan, FromLocal);
                    if (UserActionFileArchiveCancel)
                        return;
                }
            }

            fiArchive = new FileInfo(ArchiveFileName);
            if (!fiArchive.Exists)
            {
                lblStatus.Text = "Did not copy file [" + fiArchive.FullName + "]";
                richTextBoxLabSheetSender.AppendText("Did not copy file [" + fiArchive.FullName + "]\r\n");
                MessageBox.Show(lblStatus.Text, "Error");
                return;
            }

            DirectoryInfo di = new DirectoryInfo(fiSamplingPlan.FullName.Replace(".txt", "") + @"\");
            di.Refresh();
            if (!di.Exists)
            {
                lblStatus.Text = "Could not find directory [" + di.FullName + "]";
                richTextBoxLabSheetSender.AppendText("Could not find directory [" + di.FullName + "]\r\n");
                MessageBox.Show(lblStatus.Text, "Error");
                return;
            }
            DirectoryInfo diShared = new DirectoryInfo(textBoxSharedArchivedDirectory.Text + @"\" + fiSamplingPlan.Name.Replace(".txt", "") + @"\");
            if (!diShared.Exists)
            {
                try
                {
                    diShared.Create();
                }
                catch (Exception ex)
                {
                    lblStatus.Text = "Error: " + ex.Message + (ex.InnerException != null ? ex.InnerException.Message : "");
                    richTextBoxLabSheetSender.AppendText("Error:" + ex.Message + (ex.InnerException != null ? ex.InnerException.Message : "") + "\r\n");
                    MessageBox.Show(lblStatus.Text, "Error");
                    return;
                }
            }

            diShared = new DirectoryInfo(diShared.FullName);
            diShared.Refresh();
            if (!diShared.Exists)
            {
                lblStatus.Text = "Could not find directory [" + diShared.FullName + "]";
                richTextBoxLabSheetSender.AppendText("Could not find directory [" + diShared.FullName + "]\r\n");
                MessageBox.Show(lblStatus.Text, "Error");
                return;
            }

            SyncArchives(di, diShared);
        }
        private void UpdatePanelApp()
        {
            InLoadingFile = true;
            CleanAppInputPanel();

            NameCurrent = ((FileItem)comboBoxSubsectorNames.SelectedItem).Name;
            TVItemIDCurrent = ((FileItem)comboBoxSubsectorNames.SelectedItem).TVItemID;
            CSSPWQInputParamCurrent = csspWQInputParamList.Where(c => c.TVItemID == TVItemIDCurrent).FirstOrDefault();
            int month = dateTimePickerRun.Value.Month;
            int day = dateTimePickerRun.Value.Day;
            YearMonthDayCurrent = dateTimePickerRun.Value.Year + "_" + (month > 9 ? month.ToString() : "0" + month) + "_" + (day > 9 ? day.ToString() : "0" + day);

            RunNumberCurrent = (string)comboBoxRunNumber.SelectedItem;

            FileInfo fi = new FileInfo(CurrentPath);
            if (NameCurrent.Contains(" "))
            {
                fi = new FileInfo(CurrentPath + YearMonthDayCurrent.Substring(0, 4) + @"\" + (FileListViewTotalColiformLabSheets ? @"Total Coliform\" : "") + NameCurrent.Substring(0, NameCurrent.IndexOf(" ")) + "_" + YearMonthDayCurrent + "_" + csspWQInputSheetType.ToString() + "_R" + RunNumberCurrent + "_C.txt");
                if (!fi.Exists)
                    fi = new FileInfo(CurrentPath + YearMonthDayCurrent.Substring(0, 4) + @"\" + (FileListViewTotalColiformLabSheets ? @"Total Coliform\" : "") + NameCurrent.Substring(0, NameCurrent.IndexOf(" ")) + "_" + YearMonthDayCurrent + "_" + csspWQInputSheetType.ToString() + "_R" + RunNumberCurrent + "_S.txt");
                if (!fi.Exists)
                    fi = new FileInfo(CurrentPath + YearMonthDayCurrent.Substring(0, 4) + @"\" + (FileListViewTotalColiformLabSheets ? @"Total Coliform\" : "") + NameCurrent.Substring(0, NameCurrent.IndexOf(" ")) + "_" + YearMonthDayCurrent + "_" + csspWQInputSheetType.ToString() + "_R" + RunNumberCurrent + "_R.txt");
                if (!fi.Exists)
                    fi = new FileInfo(CurrentPath + YearMonthDayCurrent.Substring(0, 4) + @"\" + (FileListViewTotalColiformLabSheets ? @"Total Coliform\" : "") + NameCurrent.Substring(0, NameCurrent.IndexOf(" ")) + "_" + YearMonthDayCurrent + "_" + csspWQInputSheetType.ToString() + "_R" + RunNumberCurrent + "_A.txt");
                if (!fi.Exists)
                    fi = new FileInfo(CurrentPath + YearMonthDayCurrent.Substring(0, 4) + @"\" + (FileListViewTotalColiformLabSheets ? @"Total Coliform\" : "") + NameCurrent.Substring(0, NameCurrent.IndexOf(" ")) + "_" + YearMonthDayCurrent + "_" + csspWQInputSheetType.ToString() + "_R" + RunNumberCurrent + "_E.txt");
                if (!fi.Exists)
                    fi = new FileInfo(CurrentPath + YearMonthDayCurrent.Substring(0, 4) + @"\" + (FileListViewTotalColiformLabSheets ? @"Total Coliform\" : "") + NameCurrent.Substring(0, NameCurrent.IndexOf(" ")) + "_" + YearMonthDayCurrent + "_" + csspWQInputSheetType.ToString() + "_R" + RunNumberCurrent + "_F.txt");

            }
            else
            {
                fi = new FileInfo(CurrentPath + YearMonthDayCurrent.Substring(0, 4) + @"\" + (FileListViewTotalColiformLabSheets ? @"Total Coliform\" : "") + NameCurrent + "_" + YearMonthDayCurrent + "_" + csspWQInputSheetType.ToString() + "_R" + RunNumberCurrent + "_C.txt");
                if (!fi.Exists)
                    fi = new FileInfo(CurrentPath + YearMonthDayCurrent.Substring(0, 4) + @"\" + (FileListViewTotalColiformLabSheets ? @"Total Coliform\" : "") + NameCurrent + "_" + YearMonthDayCurrent + "_" + csspWQInputSheetType.ToString() + "_R" + RunNumberCurrent + "_S.txt");
                if (!fi.Exists)
                    fi = new FileInfo(CurrentPath + YearMonthDayCurrent.Substring(0, 4) + @"\" + (FileListViewTotalColiformLabSheets ? @"Total Coliform\" : "") + NameCurrent + "_" + YearMonthDayCurrent + "_" + csspWQInputSheetType.ToString() + "_R" + RunNumberCurrent + "_R.txt");
                if (!fi.Exists)
                    fi = new FileInfo(CurrentPath + YearMonthDayCurrent.Substring(0, 4) + @"\" + (FileListViewTotalColiformLabSheets ? @"Total Coliform\" : "") + NameCurrent + "_" + YearMonthDayCurrent + "_" + csspWQInputSheetType.ToString() + "_R" + RunNumberCurrent + "_A.txt");
                if (!fi.Exists)
                    fi = new FileInfo(CurrentPath + YearMonthDayCurrent.Substring(0, 4) + @"\" + (FileListViewTotalColiformLabSheets ? @"Total Coliform\" : "") + NameCurrent + "_" + YearMonthDayCurrent + "_" + csspWQInputSheetType.ToString() + "_R" + RunNumberCurrent + "_E.txt");
                if (!fi.Exists)
                    fi = new FileInfo(CurrentPath + YearMonthDayCurrent.Substring(0, 4) + @"\" + (FileListViewTotalColiformLabSheets ? @"Total Coliform\" : "") + NameCurrent + "_" + YearMonthDayCurrent + "_" + csspWQInputSheetType.ToString() + "_R" + RunNumberCurrent + "_F.txt");
            }
            lblFilePath.Text = "";
            if (fi.Exists)
            {
                butOpen.Enabled = true;
                lblFilePath.Text = fi.FullName;
                if (!ReadFileFromLocalMachine())
                    return;
                panelAppInput.BringToFront();
                CurrentPanel = panelAppInput;
                panelAppInputIsVisible = true;
                butSendToServer.Enabled = false;
                //butGetLabSheetsStatus.Enabled = false;

                if (fi.FullName.EndsWith("_S.txt"))
                {
                    butSendToServer.Text = "Already saved on server";
                    butSendToServer.Enabled = false;
                }
                else if (fi.FullName.EndsWith("_R.txt"))
                {
                    butSendToServer.Text = "Rejected on server";
                    butSendToServer.Enabled = false;
                }
                else if (fi.FullName.EndsWith("_A.txt"))
                {
                    butSendToServer.Text = "Accepted on server";
                    butSendToServer.Enabled = false;
                }
                else if (fi.FullName.EndsWith("_E.txt"))
                {
                    butSendToServer.Text = "Error no server action";
                    butSendToServer.Enabled = false;
                }
                else if (fi.FullName.EndsWith("_F.txt"))
                {
                    butSendToServer.Text = "Fail no server action";
                    butSendToServer.Enabled = false;
                }
                else
                {
                    butSendToServer.Text = "Send to Server";
                    butSendToServer.Enabled = true;
                }

                lblStatus.Text = "";
                butCreateFile.Visible = false;
            }
            else
            {
                SetupDataGridViewCSSP();
                butCreateFile.Visible = true;
                SetupAppInputFiles();
            }

            InLoadingFile = false;
            if (InternetConnection)
            {
                butGetLabSheetsStatus.Enabled = true;
                //if (comboBoxSubsectorNames.SelectedIndex != 0)
                //{
                //    butGetLabSheetsStatus.Enabled = true;
                //}
            }

        }
        private void ValidateCellA1(int ColumnIndex, int RowIndex)
        {
            switch (ColumnIndex)
            {
                case 2:
                    {
                        ValidateTimeCell(ColumnIndex, RowIndex);
                    }
                    break;
                case 3:
                    {
                        int SiteColumn = 1;
                        int SampleTypeColumn = 10;
                        int MPNColumn = 3;

                        string MPNValueText = dataGridViewCSSP[MPNColumn, RowIndex].Value.ToString();
                        if (MPNValueText == "1")
                        {
                            MPNValueText = "< 2";
                        }
                        if (MPNValueText == "1700")
                        {
                            MPNValueText = "> 1600";
                        }

                        string AfterSampleTypeSpace = GetAfterSampleTypeSpace(dataGridViewCSSP[SampleTypeColumn, RowIndex].Value.ToString());

                        dataGridViewCSSP[0, RowIndex].Value = dataGridViewCSSP[SiteColumn, RowIndex].Value.ToString() + " - " +
                        dataGridViewCSSP[SampleTypeColumn, RowIndex].Value.ToString() + AfterSampleTypeSpace +
                        SpaceStr.Substring(0, SpaceStr.Length - MPNValueText.Length) + MPNValueText.ToString();

                        if (dataGridViewCSSP[0, RowIndex].Value.ToString().ToUpper().Contains("ERROR"))
                        {
                            dataGridViewCSSP[0, RowIndex].Style.ForeColor = Color.Red;
                        }
                        else
                        {
                            dataGridViewCSSP[0, RowIndex].Style.ForeColor = Color.Black;
                        }

                    }
                    break;
                case 4:
                case 5:
                case 6:
                    {
                        ValidatePositiveTubeCell(ColumnIndex, RowIndex);
                    }
                    break;
                case 7:
                    {
                        ValidateSalinityCell(ColumnIndex, RowIndex);
                    }
                    break;
                case 8:
                    {
                        ValidateTemperatureCell(ColumnIndex, RowIndex);
                    }
                    break;
                case 9: // Processed by
                    {
                        int SiteColumn = 1;
                        int SampleTypeColumn = 10;
                        if (dataGridViewCSSP[ColumnIndex, RowIndex].Value == null || string.IsNullOrWhiteSpace(dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString()))
                            return;

                        dataGridViewCSSP[ColumnIndex, RowIndex].Value = dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString().ToUpper();

                        string SiteName = dataGridViewCSSP[SiteColumn, RowIndex].Value.ToString();
                        for (int i = RowIndex + 1, countRow = dataGridViewCSSP.Rows.Count; i < countRow; i++)
                        {
                            //if (dataGridViewCSSP[SiteColumn, i].Value.ToString() == SiteName)
                            //{
                            //    dataGridViewCSSP[ColumnIndex, i].Value = dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString();
                            //}
                        }

                        if (labSheetA1Sheet.LabSheetA1MeasurementList[RowIndex].ProcessedBy != dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString())
                        {
                            labSheetA1Sheet.LabSheetA1MeasurementList[RowIndex].ProcessedBy = dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString();
                            AddLog("CSSP Grid(" + ColumnIndex + "," + RowIndex + ") " + dataGridViewCSSP[SiteColumn, RowIndex].Value.ToString() + (dataGridViewCSSP[SampleTypeColumn, RowIndex].Value.ToString() == "Daily Duplicate" ? " Daily Duplicate" : "") + " - " + " [Processed By] ", dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString());
                        }

                    }
                    break;
                case 10:
                    {
                    }
                    break;
                case 12:
                    {
                        int SiteColumn = 1;
                        int SampleTypeColumn = 10;

                        if (labSheetA1Sheet.LabSheetA1MeasurementList[RowIndex].SiteComment != null && dataGridViewCSSP[ColumnIndex, RowIndex].Value == null)
                        {
                            labSheetA1Sheet.LabSheetA1MeasurementList[RowIndex].SiteComment = "";
                            AddLog("CSSP Grid(" + ColumnIndex + "," + RowIndex + ") " + dataGridViewCSSP[SiteColumn, RowIndex].Value.ToString() + (dataGridViewCSSP[SampleTypeColumn, RowIndex].Value.ToString() == "Daily Duplicate" ? " Daily Duplicate" : "") + " - " + " [Comment] ", "");
                        }
                        else
                        {
                            if (labSheetA1Sheet.LabSheetA1MeasurementList[RowIndex].SiteComment != dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString())
                            {
                                labSheetA1Sheet.LabSheetA1MeasurementList[RowIndex].SiteComment = dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString();
                                AddLog("CSSP Grid(" + ColumnIndex + "," + RowIndex + ") " + dataGridViewCSSP[SiteColumn, RowIndex].Value.ToString() + (dataGridViewCSSP[SampleTypeColumn, RowIndex].Value.ToString() == "Daily Duplicate" ? " Daily Duplicate" : "") + " - " + " [Comment] ", dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString());
                            }
                        }
                    }
                    break;
                default:
                    break;
            }
        }
        private void ValidateCellEC(DataGridViewCellEventArgs e)
        {
            throw new NotImplementedException();
        }
        private void ValidateCellLTB(DataGridViewCellEventArgs e)
        {
            throw new NotImplementedException();
        }
        private void ValidatePositiveTubeCell(int ColumnIndex, int RowIndex)
        {
            int MPNColumn = 3;
            int SiteColumn = 1;
            int SampleTypeColumn = 10;
            const int Tube10 = 4;
            const int Tube1 = 5;
            const int Tube01 = 6;
            lblDailyDuplicateRLogValue.ForeColor = Color.Black;
            lblDailyDuplicateAcceptableOrUnacceptable.ForeColor = Color.Black;
            lblDailyDuplicateRLogValue.Text = "Not calculated";
            lblDailyDuplicateAcceptableOrUnacceptable.Text = "Unknown";
            lblIntertechDuplicateRLogValue.ForeColor = Color.Black;
            lblIntertechDuplicateAcceptableOrUnacceptable.ForeColor = Color.Black;
            lblIntertechDuplicateRLogValue.Text = "Not calculated";
            lblIntertechDuplicateAcceptableOrUnacceptable.Text = "Unknown";
            lblIntertechReadAcceptableOrUnacceptable.ForeColor = Color.Black;
            lblIntertechReadAcceptableOrUnacceptable.Text = "Unknown";

            if (dataGridViewCSSP[ColumnIndex, RowIndex].Value != null)
            {
                dataGridViewCSSP[ColumnIndex, RowIndex].Value = dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString().Trim();

                if (string.IsNullOrWhiteSpace(dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString()))
                {
                    dataGridViewCSSP[ColumnIndex, RowIndex].Value = null;
                }
            }
            else
            {
            }
            dataGridViewCSSP[MPNColumn, RowIndex].Value = "";

            dataGridViewCSSP[ColumnIndex, RowIndex].Style.ForeColor = Color.Black;
            dataGridViewCSSP[MPNColumn, RowIndex].Value = "";
            if (dataGridViewCSSP[ColumnIndex, RowIndex].Value == null)
            {
                ValidateCellA1(MPNColumn, RowIndex);
                return;
            }

            string val = dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString();
            foreach (char c in val)
            {
                if (!char.IsNumber(c))
                {
                    dataGridViewCSSP[MPNColumn, RowIndex].Style.ForeColor = Color.Red;
                    dataGridViewCSSP[MPNColumn, RowIndex].Value = "Error";
                    dataGridViewCSSP[ColumnIndex, RowIndex].Style.ForeColor = Color.Red;
                    ValidateCellA1(MPNColumn, RowIndex);
                    return;
                }
            }

            int valInt = -1;
            int.TryParse(val, out valInt);
            if (valInt > 5 || valInt.ToString() != val)
            {
                dataGridViewCSSP[MPNColumn, RowIndex].Style.ForeColor = Color.Red;
                dataGridViewCSSP[MPNColumn, RowIndex].Value = "Error";
                dataGridViewCSSP[ColumnIndex, RowIndex].Style.ForeColor = Color.Red;
                ValidateCellA1(MPNColumn, RowIndex);
                return;
            }

            TryToCalculateMPNA1(RowIndex);
            ValidateCellA1(MPNColumn, RowIndex);

            switch (ColumnIndex)
            {
                case Tube10:
                    {
                        int TempInt = -1;
                        if (dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString().Length > 0)
                        {
                            int.TryParse(dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString(), out TempInt);
                            if (labSheetA1Sheet.LabSheetA1MeasurementList[RowIndex].Tube10 != TempInt)
                            {
                                labSheetA1Sheet.LabSheetA1MeasurementList[RowIndex].Tube10 = TempInt;
                                AddLog("CSSP Grid(" + ColumnIndex + "," + RowIndex + ") " + dataGridViewCSSP[SiteColumn, RowIndex].Value.ToString() + " " + dataGridViewCSSP[SampleTypeColumn, RowIndex].Value.ToString() + " - " + " [Tube 10] ", dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString());
                            }
                        }
                        else
                        {
                            labSheetA1Sheet.LabSheetA1MeasurementList[RowIndex].Tube10 = null;
                            AddLog("CSSP Grid(" + ColumnIndex + "," + RowIndex + ") " + dataGridViewCSSP[SiteColumn, RowIndex].Value.ToString() + " " + dataGridViewCSSP[SampleTypeColumn, RowIndex].Value.ToString() + " - " + " [Tube 10] ", dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString());
                        }
                    }
                    break;
                case Tube1:
                    {
                        int TempInt = -1;
                        if (dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString().Length > 0)
                        {
                            int.TryParse(dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString(), out TempInt);
                            if (labSheetA1Sheet.LabSheetA1MeasurementList[RowIndex].Tube1_0 != TempInt)
                            {
                                labSheetA1Sheet.LabSheetA1MeasurementList[RowIndex].Tube1_0 = TempInt;
                                AddLog("CSSP Grid(" + ColumnIndex + "," + RowIndex + ") " + dataGridViewCSSP[SiteColumn, RowIndex].Value.ToString() + " " + dataGridViewCSSP[SampleTypeColumn, RowIndex].Value.ToString() + " - " + " [Tube 1] ", dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString());
                            }
                        }
                        else
                        {
                            labSheetA1Sheet.LabSheetA1MeasurementList[RowIndex].Tube1_0 = null;
                            AddLog("CSSP Grid(" + ColumnIndex + "," + RowIndex + ") " + dataGridViewCSSP[SiteColumn, RowIndex].Value.ToString() + " " + dataGridViewCSSP[SampleTypeColumn, RowIndex].Value.ToString() + " - " + " [Tube 10] ", dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString());
                        }
                    }
                    break;
                case Tube01:
                    {
                        int TempInt = -1;
                        if (dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString().Length > 0)
                        {
                            int.TryParse(dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString(), out TempInt);
                            if (labSheetA1Sheet.LabSheetA1MeasurementList[RowIndex].Tube0_1 != TempInt)
                            {
                                labSheetA1Sheet.LabSheetA1MeasurementList[RowIndex].Tube0_1 = TempInt;
                                AddLog("CSSP Grid(" + ColumnIndex + "," + RowIndex + ") " + dataGridViewCSSP[SiteColumn, RowIndex].Value.ToString() + " " + dataGridViewCSSP[SampleTypeColumn, RowIndex].Value.ToString() + " - " + " [Tube 0.1] ", dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString());
                            }
                        }
                        else
                        {
                            labSheetA1Sheet.LabSheetA1MeasurementList[RowIndex].Tube0_1 = null;
                            AddLog("CSSP Grid(" + ColumnIndex + "," + RowIndex + ") " + dataGridViewCSSP[SiteColumn, RowIndex].Value.ToString() + " " + dataGridViewCSSP[SampleTypeColumn, RowIndex].Value.ToString() + " - " + " [Tube 10] ", dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString());
                        }
                    }
                    break;
                default:
                    break;
            }

        }
        private void ValidateSalinityCell(int ColumnIndex, int RowIndex)
        {
            int SiteColumn = 1;
            int SampleTypeColumn = 10;
            dataGridViewCSSP[ColumnIndex, RowIndex].Style.ForeColor = Color.Black;

            if (dataGridViewCSSP[ColumnIndex, RowIndex].Value != null)
            {
                dataGridViewCSSP[ColumnIndex, RowIndex].Value = dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString().Trim();
            }

            if (dataGridViewCSSP[ColumnIndex, RowIndex].Value == null)
            {
                return;
            }

            string val = dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString();
            foreach (char c in val)
            {
                if (!(char.IsNumber(c) || c.ToString() == "."))
                {
                    dataGridViewCSSP[ColumnIndex, RowIndex].Style.ForeColor = Color.Red;
                    return;
                }
            }

            float valFloat = -1.0f;
            float.TryParse(val, out valFloat);
            if (valFloat > 36 || valFloat.ToString() != val)
            {
                dataGridViewCSSP[ColumnIndex, RowIndex].Style.ForeColor = Color.Red;
                return;
            }

            string SiteName = dataGridViewCSSP[SiteColumn, RowIndex].Value.ToString();
            for (int i = RowIndex + 1, countRow = dataGridViewCSSP.Rows.Count; i < countRow; i++)
            {
                //if (dataGridViewCSSP[SiteColumn, i].Value.ToString() == SiteName)
                //{
                //    dataGridViewCSSP[ColumnIndex, i].Value = dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString();
                //}
            }

            float TempFloat = -1;
            if (!string.IsNullOrWhiteSpace(dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString()))
            {
                float.TryParse(dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString(), out TempFloat);
                if (labSheetA1Sheet.LabSheetA1MeasurementList[RowIndex].Salinity != TempFloat)
                {
                    labSheetA1Sheet.LabSheetA1MeasurementList[RowIndex].Salinity = TempFloat;
                    AddLog("CSSP Grid(" + ColumnIndex + "," + RowIndex + ") " + dataGridViewCSSP[SiteColumn, RowIndex].Value.ToString() + (dataGridViewCSSP[SampleTypeColumn, RowIndex].Value.ToString() == "Duplicate" ? " Dupliate" : "") + " - " + " [Salinity] ", dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString());
                }
            }
            else
            {
                labSheetA1Sheet.LabSheetA1MeasurementList[RowIndex].Salinity = null;
                AddLog("CSSP Grid(" + ColumnIndex + "," + RowIndex + ") " + dataGridViewCSSP[SiteColumn, RowIndex].Value.ToString() + (dataGridViewCSSP[SampleTypeColumn, RowIndex].Value.ToString() == "Duplicate" ? " Dupliate" : "") + " - " + " [Salinity] ", dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString());
            }

        }
        private void ValidateTemperatureCell(int ColumnIndex, int RowIndex)
        {
            int SiteColumn = 1;
            int SampleTypeColumn = 10;
            dataGridViewCSSP[ColumnIndex, RowIndex].Style.ForeColor = Color.Black;

            if (dataGridViewCSSP[ColumnIndex, RowIndex].Value != null)
            {
                dataGridViewCSSP[ColumnIndex, RowIndex].Value = dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString().Trim();
            }


            if (dataGridViewCSSP[ColumnIndex, RowIndex].Value == null)
            {
                CalculateTCFirstAndAverage();
                return;
            }

            string val = dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString();
            foreach (char c in val)
            {
                if (!(char.IsNumber(c) || c.ToString() == "."))
                {
                    dataGridViewCSSP[ColumnIndex, RowIndex].Style.ForeColor = Color.Red;
                    CalculateTCFirstAndAverage();
                    return;
                }
            }

            float valFloat = -1.0f;
            float.TryParse(val, out valFloat);
            if (valFloat > 36 || valFloat.ToString() != val)
            {
                dataGridViewCSSP[ColumnIndex, RowIndex].Style.ForeColor = Color.Red;
                CalculateTCFirstAndAverage();
                return;
            }

            string SiteName = dataGridViewCSSP[SiteColumn, RowIndex].Value.ToString();
            for (int i = RowIndex + 1, countRow = dataGridViewCSSP.Rows.Count; i < countRow; i++)
            {
                //if (dataGridViewCSSP[SiteColumn, i].Value.ToString() == SiteName)
                //{
                //    dataGridViewCSSP[ColumnIndex, i].Value = dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString();
                //}
            }

            CalculateTCFirstAndAverage();

            float TempFloat = -1;
            if (!string.IsNullOrWhiteSpace(dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString()))
            {
                float.TryParse(dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString(), out TempFloat);
                if (labSheetA1Sheet.LabSheetA1MeasurementList[RowIndex].Temperature != TempFloat)
                {
                    labSheetA1Sheet.LabSheetA1MeasurementList[RowIndex].Temperature = TempFloat;
                    AddLog("CSSP Grid(" + ColumnIndex + "," + RowIndex + ") " + dataGridViewCSSP[SiteColumn, RowIndex].Value.ToString() + (dataGridViewCSSP[SampleTypeColumn, RowIndex].Value.ToString() == "Duplicate" ? " Dupliate" : "") + " - " + " [Temperature] ", dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString());
                }
            }
            else
            {
                labSheetA1Sheet.LabSheetA1MeasurementList[RowIndex].Temperature = null;
                AddLog("CSSP Grid(" + ColumnIndex + "," + RowIndex + ") " + dataGridViewCSSP[SiteColumn, RowIndex].Value.ToString() + (dataGridViewCSSP[SampleTypeColumn, RowIndex].Value.ToString() == "Duplicate" ? " Dupliate" : "") + " - " + " [Temperature] ", dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString());
            }

        }
        private void ValidateTimeCell(int ColumnIndex, int RowIndex)
        {
            int SiteColumn = 1;
            int SampleTypeColumn = 10;

            textBoxTides.Text = "-- / --";

            if (dataGridViewCSSP[ColumnIndex, RowIndex].Value != null)
            {
                dataGridViewCSSP[ColumnIndex, RowIndex].Value = dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString().Trim();
            }

            dataGridViewCSSP[ColumnIndex, RowIndex].Style.ForeColor = Color.Black;

            if (dataGridViewCSSP[ColumnIndex, RowIndex].Value == null)
            {
                CalculateTCFirstAndAverage();
                return;
            }

            string val = dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString();
            foreach (char c in val)
            {
                if (char.IsNumber(c) || c.ToString() == ":")
                {
                }
                else
                {
                    dataGridViewCSSP[ColumnIndex, RowIndex].Style.ForeColor = Color.Red;
                    CalculateTCFirstAndAverage();
                    return;
                }
            }
            if (val.Length == 4)
            {
                if (!val.Contains(":"))
                {
                    dataGridViewCSSP[ColumnIndex, RowIndex].Value = val.Substring(0, 2) + ":" + val.Substring(2, 2);
                }
                else
                {
                    dataGridViewCSSP[ColumnIndex, RowIndex].Style.ForeColor = Color.Red;
                    CalculateTCFirstAndAverage();
                    return;
                }
            }

            val = dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString();
            if (val.Length == 5)
            {
                int intVal = -1;
                if (!val.Contains(":"))
                {
                    dataGridViewCSSP[ColumnIndex, RowIndex].Style.ForeColor = Color.Red;
                    return;
                }
                if (!int.TryParse(val.Substring(0, 2), out intVal))
                {
                    dataGridViewCSSP[ColumnIndex, RowIndex].Style.ForeColor = Color.Red;
                    return;
                }
                if (!int.TryParse(val.Substring(3, 2), out intVal))
                {
                    dataGridViewCSSP[ColumnIndex, RowIndex].Style.ForeColor = Color.Red;
                    return;
                }
                if (!(int.Parse(val.Substring(0, 2)) >= 0) || !(int.Parse(val.Substring(0, 2)) <= 23))
                {
                    dataGridViewCSSP[ColumnIndex, RowIndex].Style.ForeColor = Color.Red;
                    return;
                }
                if (!(int.Parse(val.Substring(3, 2)) >= 0) || !(int.Parse(val.Substring(3, 2)) <= 59))
                {
                    dataGridViewCSSP[ColumnIndex, RowIndex].Style.ForeColor = Color.Red;
                    return;
                }
            }
            if (val.Length < 4)
            {
                dataGridViewCSSP[ColumnIndex, RowIndex].Style.ForeColor = Color.Red;
                return;
            }
            CalculateTCFirstAndAverage();

            string SiteName = dataGridViewCSSP[SiteColumn, RowIndex].Value.ToString();
            for (int i = RowIndex + 1, countRow = dataGridViewCSSP.Rows.Count; i < countRow; i++)
            {
                //if (dataGridViewCSSP[SiteColumn, i].Value.ToString() == SiteName)
                //{
                //    dataGridViewCSSP[ColumnIndex, i].Value = dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString();
                //    int Hour = 0;
                //    int Minute = 0;

                //    if (dataGridViewCSSP[ColumnIndex, i].Value.ToString().Length > 2)
                //    {
                //        int.TryParse(dataGridViewCSSP[ColumnIndex, i].Value.ToString().Substring(0, 2), out Hour);
                //    }
                //    if (dataGridViewCSSP[ColumnIndex, i].Value.ToString().Length > 4)
                //    {
                //        int.TryParse(dataGridViewCSSP[ColumnIndex, i].Value.ToString().Substring(3, 2), out Minute);
                //    }

                //    labSheetA1Sheet.LabSheetA1MeasurementList[RowIndex].Time = new DateTime(int.Parse(labSheetA1Sheet.RunYear), int.Parse(labSheetA1Sheet.RunMonth), int.Parse(labSheetA1Sheet.RunDay), Hour, Minute, 0);
                //}
            }

            SetButGetTidesEnabledOrNot();

            if (labSheetA1Sheet.LabSheetA1MeasurementList[RowIndex].Time != null)
            {
                if (((DateTime)labSheetA1Sheet.LabSheetA1MeasurementList[RowIndex].Time).ToString("HH:mm") != dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString())
                {
                    int Hour = 0;
                    int Minute = 0;

                    if (dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString().Length > 2)
                    {
                        int.TryParse(dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString().Substring(0, 2), out Hour);
                    }
                    if (dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString().Length > 4)
                    {
                        int.TryParse(dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString().Substring(3, 2), out Minute);
                    }
                    labSheetA1Sheet.LabSheetA1MeasurementList[RowIndex].Time = new DateTime(int.Parse(labSheetA1Sheet.RunYear), int.Parse(labSheetA1Sheet.RunMonth), int.Parse(labSheetA1Sheet.RunDay), Hour, Minute, 0); ;
                }
            }

            AddLog("CSSP Grid(" + ColumnIndex + "," + RowIndex + ") " + dataGridViewCSSP[SiteColumn, RowIndex].Value.ToString() + (dataGridViewCSSP[SampleTypeColumn, RowIndex].Value.ToString() == "Duplicate" ? " Dupliate" : "") + " - " + " [Time] ", dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString());
        }


        #endregion Form Functions

        #endregion Functions

        class AcceptedOrRejected
        {
            public string AcceptedOrRejectedBy { get; set; }
            public DateTime AcceptedOrRejectedDate { get; set; }
            public string RejectReason { get; set; }
        }
    }


}




