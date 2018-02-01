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
        #region Events Form
        private void CSSPWQInputToolForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            while (IsSaving == true)
            {
                Application.DoEvents();
            }

            lblFilePath.Text = "";
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
            while (IsSaving == true)
            {
                Application.DoEvents();
            }

            if (lblFilePath.Text.EndsWith("_C.txt"))
            {
                LogAll();
            }

            SetupAppInputFiles();

            IsSaving = false;
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
            butArchive.Enabled = true;
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
            OpenedFileName = lblFilePath.Text;
        }
        private void butGetTides_Click(object sender, EventArgs e)
        {
            if (!InLoadingFile)
            {
                if (!PreModifying())
                {
                    return;
                }
            }

            while (IsSaving == true)
            {
                Application.DoEvents();
            }

            if (lblFilePath.Text.EndsWith("_C.txt"))
            {
                LogAll();
            }

            IsSaving = false;

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
            while (IsSaving == true)
            {
                Application.DoEvents();
            }

            if (lblFilePath.Text.EndsWith("_C.txt"))
            {
                LogAll();
            }

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
            lblFilePath.Text = "";
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
            if (lblFilePath.Text.EndsWith("_C.txt"))
            {
                LogAll();
            }

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
            if (!InLoadingFile)
            {
                if (checkBox2Coolers.Checked.ToString().ToLower() != Start2Coolers)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        checkBox2Coolers.Checked = Start2Coolers == "true" ? true : false;
                        return;
                    }
                }
            }

            if (checkBox2Coolers.Checked)
            {
                checkBox2Coolers.ForeColor = Color.Green;
                textBoxTCField2.Visible = true;
                textBoxTCLab2.Visible = true;
            }
            else
            {
                checkBox2Coolers.ForeColor = Color.Black;
                textBoxTCField2.Text = "";
                textBoxTCLab2.Text = "";
                textBoxTCField2.Visible = false;
                textBoxTCLab2.Visible = false;
            }
        }
        private void checkBoxIncubationStartSameDay_CheckedChanged(object sender, EventArgs e)
        {
            if (!InLoadingFile)
            {
                if (checkBoxIncubationStartSameDay.Checked.ToString().ToLower() != StartIncubationStartSameDay)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        checkBoxIncubationStartSameDay.Checked = StartIncubationStartSameDay == "true" ? true : false;
                        return;
                    }
                }
            }

            if (checkBoxIncubationStartSameDay.Checked)
            {
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

            RunNumberAndText runNumberAndText = (RunNumberAndText)comboBoxRunNumber.SelectedItem;
            if (runNumberAndText == null)
            {
                ShouldUpdatePanelApp = false;
            }
            else
            {
                RunNumberCurrent = (runNumberAndText.RunNumber < 10 ? "0" + runNumberAndText.RunNumber : runNumberAndText.RunNumber.ToString());
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
        private void dataGridViewCSSP_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (!InLoadingFile)
            {
                if (PreModifying())
                {
                    IsSaving = true;
                    Modifying();
                }
                else
                {
                    string val = dataGridViewCSSP[e.ColumnIndex, e.RowIndex].Value == null ? "" : dataGridViewCSSP[e.ColumnIndex, e.RowIndex].Value.ToString();
                    if (val != StartGridCellText[e.RowIndex][e.ColumnIndex])
                    {
                        dataGridViewCSSP[e.ColumnIndex, e.RowIndex].Value = StartGridCellText[e.RowIndex][e.ColumnIndex];
                    }
                    return;
                }
            }

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
        private void dataGridViewCSSP_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            //if (checkBox2Coolers.Checked.ToString().ToLower() != Start2Coolers)
            //{
            //    if (PreModifying())
            //    {
            //        IsSaving = true;
            //        Modifying();
            //    }
            //    else
            //    {
            //        checkBox2Coolers.Checked = Start2Coolers == "true" ? true : false;
            //        return;
            //    }
            //}

            //dataGridViewCSSP.BackgroundColor = DataGridViewCSSPBackgroundColor;
            //if (csspWQInputSheetType == CSSPWQInputSheetTypeEnum.LTB)
            //{
            //    ValidateCellLTB(e);
            //}
            //else if (csspWQInputSheetType == CSSPWQInputSheetTypeEnum.EC)
            //{
            //    ValidateCellEC(e);
            //}
            //else
            //{
            //    ValidateCellA1(e.ColumnIndex, e.RowIndex);
            //}
            //CalculateDuplicate();
            //Modifying();
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
                        string cellStr = dataGridViewCSSP.Rows[e.RowIndex].Cells[e.ColumnIndex].Value == null ? "" : dataGridViewCSSP.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
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
                                            break;
                                        }
                                    }
                                }
                            }
                            else if (e.ColumnIndex == ProcessByColumn)
                            {
                                if (!(dataGridViewCSSP.Rows[e.RowIndex].Cells[SampleTypeColumn].Value.ToString() == "DailyDuplicate"
                                           || dataGridViewCSSP.Rows[e.RowIndex].Cells[SampleTypeColumn].Value.ToString() == "IntertechDuplicate"
                                           || dataGridViewCSSP.Rows[e.RowIndex].Cells[SampleTypeColumn].Value.ToString() == "IntertechRead"))
                                {
                                    dataGridViewCSSP.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = dataGridViewCSSP.Rows[(e.RowIndex - 1)].Cells[e.ColumnIndex].Value;
                                }
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
                if (!InLoadingFile)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        return;
                    }
                }

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
                                        ReadFileFromLocalMachine();
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
            else if (e.KeyCode == Keys.F3)
            {
                if (!InLoadingFile)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        return;
                    }
                }

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
                if (!InLoadingFile)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        return;
                    }
                }

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
                if (!InLoadingFile)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        return;
                    }
                }

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
        private void dataGridViewCSSP_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            dataGridViewCSSP.Rows[e.RowIndex].Cells[0].Style.BackColor = Color.Aqua;
        }

        private void dataGridViewCSSP_RowLeave(object sender, DataGridViewCellEventArgs e)
        {

            dataGridViewCSSP.Rows[e.RowIndex].Cells[0].Style.BackColor = Color.LightGray;
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
            if (!InLoadingFile)
            {
                if (PreModifying())
                {
                    IsSaving = true;
                    Modifying();
                }
                else
                {
                    dateTimePickerSalinitiesReadDate.Value = new DateTime(int.Parse(StartSalinityReadDateYear), int.Parse(StartSalinityReadDateMonth), int.Parse(StartSalinityReadDateDay));
                    return;
                }
            }
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
        #region Events Focus Leave
        private void checkBox2Coolers_Leave(object sender, EventArgs e)
        {
            string CheckBoxText = (checkBox2Coolers.Checked ? "true" : "false");
            if (labSheetA1Sheet.TCHas2Coolers != CheckBoxText)
            {
                labSheetA1Sheet.TCHas2Coolers = CheckBoxText;
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
            }
        }
        private void richTextBoxRunWeatherComment_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.RunWeatherComment != richTextBoxRunWeatherComment.Text)
            {
                labSheetA1Sheet.RunWeatherComment = richTextBoxRunWeatherComment.Text;
            }
        }
        private void richTextBoxRunComment_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.RunComment != richTextBoxRunComment.Text)
            {
                labSheetA1Sheet.RunWeatherComment = richTextBoxRunComment.Text;
            }
        }
        private void textBoxControlBlank35_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.Blank35 != textBoxControlBlank35.Text)
            {
                labSheetA1Sheet.Blank35 = textBoxControlBlank35.Text;
            }
        }
        private void textBoxControlBath1Blank44_5_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.Bath1Blank44_5 != textBoxControlBath1Blank44_5.Text)
            {
                labSheetA1Sheet.Bath1Blank44_5 = textBoxControlBath1Blank44_5.Text;
            }
        }
        private void textBoxControlBath2Blank44_5_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.Bath2Blank44_5 != textBoxControlBath2Blank44_5.Text)
            {
                labSheetA1Sheet.Bath2Blank44_5 = textBoxControlBath2Blank44_5.Text;
            }
        }
        private void textBoxControlBath3Blank44_5_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.Bath3Blank44_5 != textBoxControlBath3Blank44_5.Text)
            {
                labSheetA1Sheet.Bath3Blank44_5 = textBoxControlBath3Blank44_5.Text;
            }
        }
        private void textBoxControlLot_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.ControlLot != textBoxControlLot.Text)
            {
                labSheetA1Sheet.ControlLot = textBoxControlLot.Text;
            }
        }
        private void textBoxControlNegative35_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.Negative35 != textBoxControlNegative35.Text)
            {
                labSheetA1Sheet.Negative35 = textBoxControlNegative35.Text;
            }
        }
        private void textBoxControlBath1Negative44_5_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.Bath1Negative44_5 != textBoxControlBath1Negative44_5.Text)
            {
                labSheetA1Sheet.Bath1Negative44_5 = textBoxControlBath1Negative44_5.Text;
            }
        }
        private void textBoxControlBath2Negative44_5_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.Bath2Negative44_5 != textBoxControlBath2Negative44_5.Text)
            {
                labSheetA1Sheet.Bath2Negative44_5 = textBoxControlBath2Negative44_5.Text;
            }
        }
        private void textBoxControlBath3Negative44_5_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.Bath3Negative44_5 != textBoxControlBath3Negative44_5.Text)
            {
                labSheetA1Sheet.Bath3Negative44_5 = textBoxControlBath3Negative44_5.Text;
            }
        }
        private void textBoxControlNonTarget35_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.NonTarget35 != textBoxControlNonTarget35.Text)
            {
                labSheetA1Sheet.NonTarget35 = textBoxControlNonTarget35.Text;
            }
        }
        private void textBoxControlBath1NonTarget44_5_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.Bath1NonTarget44_5 != textBoxControlBath1NonTarget44_5.Text)
            {
                labSheetA1Sheet.Bath1NonTarget44_5 = textBoxControlBath1NonTarget44_5.Text;
            }
        }
        private void textBoxControlBath2NonTarget44_5_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.Bath2NonTarget44_5 != textBoxControlBath2NonTarget44_5.Text)
            {
                labSheetA1Sheet.Bath2NonTarget44_5 = textBoxControlBath2NonTarget44_5.Text;
            }
        }
        private void textBoxControlBath3NonTarget44_5_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.Bath3NonTarget44_5 != textBoxControlBath3NonTarget44_5.Text)
            {
                labSheetA1Sheet.Bath3NonTarget44_5 = textBoxControlBath3NonTarget44_5.Text;
            }
        }
        private void textBoxControlPositive35_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.Positive35 != textBoxControlPositive35.Text)
            {
                labSheetA1Sheet.Positive35 = textBoxControlPositive35.Text;
            }
        }
        private void textBoxControlBath1Positive44_5_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.Bath1Positive44_5 != textBoxControlBath1Positive44_5.Text)
            {
                labSheetA1Sheet.Bath1Positive44_5 = textBoxControlBath1Positive44_5.Text;
            }
        }
        private void textBoxControlBath2Positive44_5_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.Bath2Positive44_5 != textBoxControlBath2Positive44_5.Text)
            {
                labSheetA1Sheet.Bath2Positive44_5 = textBoxControlBath2Positive44_5.Text;
            }
        }
        private void textBoxControlBath3Positive44_5_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.Bath3Positive44_5 != textBoxControlBath3Positive44_5.Text)
            {
                labSheetA1Sheet.Bath3Positive44_5 = textBoxControlBath3Positive44_5.Text;
            }
        }
        private void textBoxDailyDuplicatePrecisionCriteria_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.DailyDuplicatePrecisionCriteria != textBoxDailyDuplicatePrecisionCriteria.Text)
            {
                labSheetA1Sheet.DailyDuplicatePrecisionCriteria = textBoxDailyDuplicatePrecisionCriteria.Text;
            }
        }
        private void textBoxIncubationBath1StartTime_Leave(object sender, EventArgs e)
        {
            textBoxIncubationBath1StartTime.ForeColor = Color.Black;
            if (!CheckTimeInTextBox(textBoxIncubationBath1StartTime))
            {
                textBoxIncubationBath1StartTime.ForeColor = Color.Red;
            }

            if (textBoxIncubationBath1StartTime.Text.Length == 5)
            {
                TryToCalculateIncubationTimeSpan();
            }
            else
            {
                textBoxIncubationBath1StartTime.ForeColor = Color.Red;
            }

            if (labSheetA1Sheet.IncubationBath1StartTime != textBoxIncubationBath1StartTime.Text)
            {
                labSheetA1Sheet.IncubationBath1StartTime = textBoxIncubationBath1StartTime.Text;
            }
        }
        private void textBoxIncubationBath2StartTime_Leave(object sender, EventArgs e)
        {
            textBoxIncubationBath2StartTime.ForeColor = Color.Black;
            if (!CheckTimeInTextBox(textBoxIncubationBath2StartTime))
            {
                textBoxIncubationBath2StartTime.ForeColor = Color.Red;
            }

            if (textBoxIncubationBath2StartTime.Text.Length == 5)
            {
                TryToCalculateIncubationTimeSpan();
            }
            else
            {
                textBoxIncubationBath2StartTime.ForeColor = Color.Red;
            }

            if (labSheetA1Sheet.IncubationBath2StartTime != textBoxIncubationBath2StartTime.Text)
            {
                labSheetA1Sheet.IncubationBath2StartTime = textBoxIncubationBath2StartTime.Text;
            }
        }
        private void textBoxIncubationBath3StartTime_Leave(object sender, EventArgs e)
        {
            textBoxIncubationBath3StartTime.ForeColor = Color.Black;
            if (!CheckTimeInTextBox(textBoxIncubationBath3StartTime))
            {
                textBoxIncubationBath3StartTime.ForeColor = Color.Red;
            }

            if (textBoxIncubationBath3StartTime.Text.Length == 5)
            {
                TryToCalculateIncubationTimeSpan();
            }
            else
            {
                textBoxIncubationBath3StartTime.ForeColor = Color.Red;
            }

            if (labSheetA1Sheet.IncubationBath3StartTime != textBoxIncubationBath3StartTime.Text)
            {
                labSheetA1Sheet.IncubationBath3StartTime = textBoxIncubationBath3StartTime.Text;
            }
        }
        private void textBoxIncubationBath1EndTime_Leave(object sender, EventArgs e)
        {
            textBoxIncubationBath1EndTime.ForeColor = Color.Black;
            if (!CheckTimeInTextBox(textBoxIncubationBath1EndTime))
            {
                textBoxIncubationBath1EndTime.ForeColor = Color.Red;
            }

            if (textBoxIncubationBath1EndTime.Text.Length == 5)
            {
                TryToCalculateIncubationTimeSpan();
            }
            else
            {
                textBoxIncubationBath1EndTime.ForeColor = Color.Red;
            }

            if (labSheetA1Sheet.IncubationBath1EndTime != textBoxIncubationBath1EndTime.Text)
            {
                labSheetA1Sheet.IncubationBath1EndTime = textBoxIncubationBath1EndTime.Text;
            }
        }
        private void textBoxIncubationBath2EndTime_Leave(object sender, EventArgs e)
        {
            textBoxIncubationBath2EndTime.ForeColor = Color.Black;
            if (!CheckTimeInTextBox(textBoxIncubationBath2EndTime))
            {
                textBoxIncubationBath2EndTime.ForeColor = Color.Red;
            }

            if (textBoxIncubationBath2EndTime.Text.Length == 5)
            {
                TryToCalculateIncubationTimeSpan();
            }
            else
            {
                textBoxIncubationBath2EndTime.ForeColor = Color.Red;
            }

            if (labSheetA1Sheet.IncubationBath2EndTime != textBoxIncubationBath2EndTime.Text)
            {
                labSheetA1Sheet.IncubationBath2EndTime = textBoxIncubationBath2EndTime.Text;
            }
        }
        private void textBoxIncubationBath3EndTime_Leave(object sender, EventArgs e)
        {
            textBoxIncubationBath3EndTime.ForeColor = Color.Black;
            if (!CheckTimeInTextBox(textBoxIncubationBath3EndTime))
            {
                textBoxIncubationBath3EndTime.ForeColor = Color.Red;
            }

            if (textBoxIncubationBath3EndTime.Text.Length == 5)
            {
                TryToCalculateIncubationTimeSpan();
            }
            else
            {
                textBoxIncubationBath3EndTime.ForeColor = Color.Red;
            }

            if (labSheetA1Sheet.IncubationBath3EndTime != textBoxIncubationBath3EndTime.Text)
            {
                labSheetA1Sheet.IncubationBath3EndTime = textBoxIncubationBath3EndTime.Text;
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
            }
        }
        private void textBoxResultsRecordedBy_Leave(object sender, EventArgs e)
        {
            textBoxResultsRecordedBy.Text = textBoxResultsRecordedBy.Text.ToUpper();

            if (labSheetA1Sheet.ResultsRecordedBy != textBoxResultsRecordedBy.Text)
            {
                labSheetA1Sheet.ResultsRecordedBy = textBoxResultsRecordedBy.Text;
            }
        }
        private void textBoxSalinitiesReadBy_Leave(object sender, EventArgs e)
        {
            textBoxSalinitiesReadBy.Text = textBoxSalinitiesReadBy.Text.ToUpper();

            if (labSheetA1Sheet.SalinitiesReadBy != textBoxSalinitiesReadBy.Text)
            {
                labSheetA1Sheet.SalinitiesReadBy = textBoxSalinitiesReadBy.Text;
            }
        }
        private void textBoxSampleBottleLotNumber_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.SampleBottleLotNumber != textBoxSampleBottleLotNumber.Text)
            {
                labSheetA1Sheet.SampleBottleLotNumber = textBoxSampleBottleLotNumber.Text;
            }
        }
        private void textBoxSampleCrewInitials_Leave(object sender, EventArgs e)
        {
            textBoxSampleCrewInitials.Text = textBoxSampleCrewInitials.Text.ToUpper();

            if (labSheetA1Sheet.SampleCrewInitials != textBoxSampleCrewInitials.Text)
            {
                labSheetA1Sheet.SampleCrewInitials = textBoxSampleCrewInitials.Text;
            }
        }
        private void textBoxTCField1_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.TCField1 != textBoxTCField1.Text)
            {
                labSheetA1Sheet.TCField1 = textBoxTCField1.Text;
            }
        }
        private void textBoxTCField2_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.TCField2 != textBoxTCField2.Text)
            {
                labSheetA1Sheet.TCField2 = textBoxTCField2.Text;
            }
        }
        private void textBoxTCLab1_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.TCLab1 != textBoxTCLab1.Text)
            {
                labSheetA1Sheet.TCLab1 = textBoxTCLab1.Text;
            }
        }
        private void textBoxTCLab2_Leave(object sender, EventArgs e)
        {
            if (labSheetA1Sheet.TCLab2 != textBoxTCLab2.Text)
            {
                labSheetA1Sheet.TCLab2 = textBoxTCLab2.Text;
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
            }
        }
        private void textBoxWaterBath1Number_Leave(object sender, EventArgs e)
        {
            textBoxWaterBath1Number.Text = textBoxWaterBath1Number.Text.ToUpper();

            if (labSheetA1Sheet.WaterBath1 != textBoxWaterBath1Number.Text)
            {
                labSheetA1Sheet.WaterBath1 = textBoxWaterBath1Number.Text;
            }
        }
        private void textBoxWaterBath2Number_Leave(object sender, EventArgs e)
        {
            textBoxWaterBath2Number.Text = textBoxWaterBath2Number.Text.ToUpper();

            if (labSheetA1Sheet.WaterBath2 != textBoxWaterBath2Number.Text)
            {
                labSheetA1Sheet.WaterBath2 = textBoxWaterBath2Number.Text;
            }
        }
        private void textBoxWaterBath3Number_Leave(object sender, EventArgs e)
        {
            textBoxWaterBath3Number.Text = textBoxWaterBath3Number.Text.ToUpper();

            if (labSheetA1Sheet.WaterBath3 != textBoxWaterBath3Number.Text)
            {
                labSheetA1Sheet.WaterBath3 = textBoxWaterBath3Number.Text;
            }
        }
        #endregion Events Focus Leave
        #region Events Focus Enter
        private void checkBox2Coolers_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Indicate if one or two coolers where used during the run.";
        }
        private void comboBoxRunNumber_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Please indicate the run number on this date.";
        }
        private void comboBoxSubsectorNames_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Please select the subsector.";
        }
        private void checkBoxIncubationStartSameDay_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Indicate if the lab analysis was started on the same day or the next day.";
        }
        private void dateTimePickerRun_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Please select the date of the Run. Pressing F2 while having a lab sheet open will let you change the date of that particular lab sheet.";
        }
        private void dateTimePickerSalinitiesReadDate_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Please select the date the salinity was read.";
        }
        private void dateTimePickerResultsReadDate_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Please select the date the results was read.";
        }
        private void dateTimePickerResultsRecordedDate_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Please select the date the results was recorded.";
        }
        private void radioButton1Baths_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Indicate the number of baths used for the run.";
        }
        private void radioButton2Baths_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Indicate the number of baths used for the run.";
        }
        private void radioButton3Baths_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Indicate the number of baths used for the run.";
        }
        private void textBoxControlBath1Blank44_5_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Only allowable characters are '+' , '-' or 'N' for not applicable";
        }
        private void textBoxControlBath1Negative44_5_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Only allowable characters are '+' and '-'";
        }
        private void textBoxControlBath1NonTarget44_5_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Only allowable characters are '+' , '-' or 'N' for not applicable";
        }
        private void textBoxControlBath1Positive44_5_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Only allowable characters are '+' and '-'";
        }
        private void textBoxControlBath2Blank44_5_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Only allowable characters are '+' , '-' or 'N' for not applicable";
        }
        private void textBoxControlBath2Negative44_5_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Only allowable characters are '+' and '-'";
        }
        private void textBoxControlBath2NonTarget44_5_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Only allowable characters are '+' , '-' or 'N' for not applicable";
        }
        private void textBoxControlBath2Positive44_5_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Only allowable characters are '+' and '-'";
        }
        private void textBoxControlBath3Blank44_5_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Only allowable characters are '+' , '-' or 'N' for not applicable";
        }
        private void textBoxControlBath3Negative44_5_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Only allowable characters are '+' and '-'";
        }
        private void textBoxControlBath3NonTarget44_5_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Only allowable characters are '+' , '-' or 'N' for not applicable";
        }
        private void textBoxControlBath3Positive44_5_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Only allowable characters are '+' and '-'";
        }
        private void textBoxControlBlank35_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Only allowable characters are '+' , '-' or 'N' for not applicable.";
        }
        private void textBoxControlLot_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Please enter control lot description.";
        }
        private void textBoxControlNegative35_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Only allowable characters are '+' and '-'.";
        }
        private void textBoxControlNonTarget35_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Only allowable characters are '+' , '-' or 'N' for not applicable.";
        }
        private void textBoxControlPositive35_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Only allowable characters are '+' and '-'.";
        }
        private void textBoxDailyDuplicatePrecisionCriteria_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Please enter the daily duplicate precision criteria for the lab.";
        }
        private void textBoxIncubationBath1EndTime_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Please enter bath #1 incubation end time. All time should be entered with 4 digits. 1234 for 12:34. ':' will be added automatically";
        }
        private void textBoxIncubationBath1StartTime_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Please enter bath #1 incubation start time. All time should be entered with 4 digits. 1234 for 12:34. ':' will be added automatically";
        }
        private void textBoxIncubationBath2EndTime_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Please enter bath #2 incubation end time. All time should be entered with 4 digits. 1234 for 12:34. ':' will be added automatically";
        }
        private void textBoxIncubationBath2StartTime_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Please enter bath #2 incubation start time. All time should be entered with 4 digits. 1234 for 12:34. ':' will be added automatically";
        }
        private void textBoxIncubationBath3EndTime_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Please enter bath #3 incubation end time. All time should be entered with 4 digits. 1234 for 12:34. ':' will be added automatically";
        }
        private void textBoxIncubationBath3StartTime_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Please enter bath #3 incubation start time. All time should be entered with 4 digits. 1234 for 12:34. ':' will be added automatically";
        }
        private void textBoxLot35_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Please enter lot number for control at 35.0 ˚C";
        }
        private void textBoxLot44_5_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Please enter lot number for control at 44.5 ˚C";
        }
        private void textBoxResultsReadBy_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Initials of person who read the results";
        }
        private void textBoxResultsRecordedBy_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Initials of person who recorded the results";
        }
        private void richTextBoxRunComment_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Anything observed during the field trip";
        }
        private void richTextBoxRunWeatherComment_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Anything related to the RunWeatherComment during the sampling";
        }
        private void textBoxSalinitiesReadBy_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Initials of person who measured the salinities in the lab";
        }
        private void textBoxSampleBottleLotNumber_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Anything representing sample bottle lot number";
        }
        private void textBoxSampleCrewInitials_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Initials of Sampling Crew. Lowercase is ok. It will be set to uppercase automatically. Separate initials with comma. Ex: JAR,PG";
        }
        private void textBoxTCField1_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Enter temperature. Only digits and '.' are accepted";
        }
        private void textBoxTCField2_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Enter temperature. Only digits and '.' are accepted";
        }
        private void textBoxTCLab1_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Enter temperature. Only digits and '.' are accepted";
        }
        private void textBoxTCLab2_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Enter temperature. Only digits and '.' are accepted";
        }
        private void textBoxTides_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Allowables are [HR, HT, HF, MR, MT, MF, LR, LT, LF]                 Ex: HT / HT";
        }
        private void textBoxWaterBath1Number_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Please enter the water bath #1 number. Text will be automatically converted to uppercase";
        }
        private void textBoxWaterBath2Number_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Please enter the water bath #1 number. Text will be automatically converted to uppercase";
        }
        private void textBoxWaterBath3Number_Enter(object sender, EventArgs e)
        {
            lblStatus.Text = "Please enter the water bath #1 number. Text will be automatically converted to uppercase";
        }
        #endregion Events Focus Enter
        #region Events KeyDown
        private void checkBox2Coolers_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    textBoxSampleCrewInitials.Focus();
                }
                else
                {
                    textBoxTCField1.Focus();
                }
            }

            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "Indicate if one or two coolers where used during the run.";
            }
        }
        private void radioButton1Baths_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    checkBoxIncubationStartSameDay.Focus();
                }
                else
                {
                    textBoxIncubationBath1StartTime.Focus();
                }
            }

            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "Indicate the number of baths used for the run.";
            }
        }

        private void radioButton2Baths_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    checkBoxIncubationStartSameDay.Focus();
                }
                else
                {
                    textBoxIncubationBath1StartTime.Focus();
                }
            }

            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "Indicate the number of baths used for the run.";
            }

        }

        private void radioButton3Baths_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    checkBoxIncubationStartSameDay.Focus();
                }
                else
                {
                    textBoxIncubationBath1StartTime.Focus();
                }
            }

            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "Indicate the number of baths used for the run.";
            }

        }
        private void checkBoxIncubationStartSameDay_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    if (checkBox2Coolers.Checked)
                    {
                        textBoxTCLab2.Focus();
                    }
                    else
                    {
                        textBoxTCLab1.Focus();
                    }
                }
                else
                {
                    if (radioButton2Baths.Checked)
                    {
                        radioButton2Baths.Focus();
                    }
                    else if (radioButton3Baths.Checked)
                    {
                        radioButton3Baths.Focus();
                    }
                    else
                    {
                        radioButton1Baths.Focus();
                    }
                }
            }

            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "Indicate if the lab analysis was started on the same day or the next day.";
            }

        }
        private void dateTimePickerResultsReadDate_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    textBoxResultsReadBy.Focus();
                }
                else
                {
                    textBoxResultsRecordedBy.Focus();
                }
            }

            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "Please select the date the results was read.";
            }

        }

        private void dateTimePickerResultsRecordedDate_KeyDown(object sender, KeyEventArgs e)
        {
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
                lblStatus.Text = "Please select the date the results was recorded.";
            }

        }
        private void dateTimePickerSalinitiesReadDate_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    textBoxSalinitiesReadBy.Focus();
                }
                else
                {
                    textBoxResultsReadBy.Focus();
                }
            }

            if (e.KeyCode == Keys.F1)
            {
                lblStatus.Text = "Please select the date the salinity was read.";
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
                lblStatus.Text = "Only allowable characters are '+' , '-' or 'N' for not applicable";
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
                lblStatus.Text = "Only allowable characters are '+' , '-' or 'N' for not applicable";
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
                lblStatus.Text = "Only allowable characters are '+' , '-' or 'N' for not applicable";
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
                lblStatus.Text = "Only allowable characters are '+' , '-' or 'N' for not applicable";
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
                lblStatus.Text = "Please enter bath #1 incubation start time. All time should be entered with 4 digits. 1234 for 12:34. ':' will be added automatically";
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
                lblStatus.Text = "Please enter bath #2 incubation start time. All time should be entered with 4 digits. 1234 for 12:34. ':' will be added automatically";
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
                lblStatus.Text = "Please enter bath #3 incubation start time. All time should be entered with 4 digits. 1234 for 12:34. ':' will be added automatically";
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
                lblStatus.Text = "Please enter bath #1 incubation end time. All time should be entered with 4 digits. 1234 for 12:34. ':' will be added automatically";
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
                lblStatus.Text = "Please enter bath #2 incubation end time. All time should be entered with 4 digits. 1234 for 12:34. ':' will be added automatically";
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
                lblStatus.Text = "Please enter bath #3 incubation end time. All time should be entered with 4 digits. 1234 for 12:34. ':' will be added automatically";
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
                    textBoxResultsReadBy.Focus();
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
                    checkBox2Coolers.Focus();
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
                    textBoxResultsReadBy.Focus();
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
                    textBoxTCLab1.Focus();
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
                    textBoxTCLab1.Focus();
                }
                else
                {
                    textBoxTCLab2.Focus();
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
                    textBoxTCField2.Focus();
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
                lblStatus.Text = "Please enter the water bath #1 number. Text will be automatically converted to uppercase";
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
                lblStatus.Text = "Please enter the water bath #1 number. Text will be automatically converted to uppercase";
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
                lblStatus.Text = "Please enter the water bath #1 number. Text will be automatically converted to uppercase";
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

        }
        #endregion Events listBoxFiles
        #region Events radioButtons
        private void radioButton1Baths_CheckedChanged(object sender, EventArgs e)
        {
            if (!InLoadingFile)
            {
                if (PreModifying())
                {
                    IsSaving = true;
                    Modifying();
                }
                else
                {
                    if (StartNumberOfBaths == "1")
                    {
                        radioButton1Baths.Checked = true;
                    }
                    return;
                }
            }

            RadioButtonBathNumberChanged();
        }
        private void radioButton2Baths_CheckedChanged(object sender, EventArgs e)
        {
            if (!InLoadingFile)
            {
                if (PreModifying())
                {
                    IsSaving = true;
                    Modifying();
                }
                else
                {
                    if (StartNumberOfBaths == "2")
                    {
                        radioButton2Baths.Checked = true;
                    }
                    return;
                }
            }

            RadioButtonBathNumberChanged();
        }
        private void radioButton3Baths_CheckedChanged(object sender, EventArgs e)
        {
            if (!InLoadingFile)
            {
                if (PreModifying())
                {
                    IsSaving = true;
                    Modifying();
                }
                else
                {
                    if (StartNumberOfBaths == "3")
                    {
                        radioButton3Baths.Checked = true;
                    }
                    return;
                }
            }

            RadioButtonBathNumberChanged();
        }
        #endregion Events radioButtons
        #region Events TextChanged
        private void lblFilePath_TextChanged(object sender, EventArgs e)
        {
            while (IsSaving == true)
            {
                Application.DoEvents();
            }

            if (OpenedFileName != "")
            {
                if (OpenedFileName != lblFilePath.Text)
                {
                    if (OpenedFileName.EndsWith("_C.txt"))
                    {
                        LogAll();
                    }
                }
            }

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
        private void richTextBoxRunComment_TextChanged(object sender, EventArgs e)
        {
            if (!InLoadingFile)
            {
                if (richTextBoxRunComment.Text != StartRunComment)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        richTextBoxRunComment.Text = StartRunComment;
                    }
                }
            }
        }
        private void richTextBoxRunWeatherComment_TextChanged(object sender, EventArgs e)
        {
            if (!InLoadingFile)
            {
                if (richTextBoxRunWeatherComment.Text != StartRunWeatherComment)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        richTextBoxRunWeatherComment.Text = StartRunWeatherComment;
                    }
                }
            }
        }
        private void textBoxControlBlank35_TextChanged(object sender, EventArgs e)
        {
            textBoxControlBlank35.ForeColor = Color.Black;
            if (textBoxControlBlank35.Text == "-" || textBoxControlBlank35.Text == "+")
            {
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
            {
                if (textBoxControlBlank35.Text != StartControlBlank35)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        textBoxControlBlank35.Text = StartControlBlank35;
                    }
                }
            }
        }
        private void textBoxControlBath1Blank44_5_TextChanged(object sender, EventArgs e)
        {
            textBoxControlBath1Blank44_5.ForeColor = Color.Black;
            if (textBoxControlBath1Blank44_5.Text == "-" || textBoxControlBath1Blank44_5.Text == "+")
            {
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
            {
                if (textBoxControlBath1Blank44_5.Text != StartControlBath1Blank44_5)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        textBoxControlBath1Blank44_5.Text = StartControlBath1Blank44_5;
                    }
                }
            }
        }
        private void textBoxControlBath2Blank44_5_TextChanged(object sender, EventArgs e)
        {
            textBoxControlBath2Blank44_5.ForeColor = Color.Black;
            if (textBoxControlBath2Blank44_5.Text == "-" || textBoxControlBath2Blank44_5.Text == "+")
            {
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
            {
                if (textBoxControlBath2Blank44_5.Text != StartControlBath2Blank44_5)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        textBoxControlBath2Blank44_5.Text = StartControlBath2Blank44_5;
                    }
                }
            }
        }
        private void textBoxControlBath3Blank44_5_TextChanged(object sender, EventArgs e)
        {
            textBoxControlBath3Blank44_5.ForeColor = Color.Black;
            if (textBoxControlBath3Blank44_5.Text == "-" || textBoxControlBath3Blank44_5.Text == "+")
            {
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
            {
                if (textBoxControlBath3Blank44_5.Text != StartControlBath3Blank44_5)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        textBoxControlBath3Blank44_5.Text = StartControlBath3Blank44_5;
                    }
                }
            }
        }
        private void textBoxControlLot_TextChanged(object sender, EventArgs e)
        {
            if (!InLoadingFile)
            {
                if (textBoxControlLot.Text != StartControlLot)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        textBoxControlLot.Text = StartControlLot;
                    }
                }
            }
        }
        private void textBoxControlNegative35_TextChanged(object sender, EventArgs e)
        {
            textBoxControlNegative35.ForeColor = Color.Black;
            if (textBoxControlNegative35.Text == "-" || textBoxControlNegative35.Text == "+")
            {
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
            {
                if (textBoxControlNegative35.Text != StartControlNegative35)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        textBoxControlNegative35.Text = StartControlNegative35;
                    }
                }
            }
        }
        private void textBoxControlBath1Negative44_5_TextChanged(object sender, EventArgs e)
        {
            textBoxControlBath1Negative44_5.ForeColor = Color.Black;
            if (textBoxControlBath1Negative44_5.Text == "-" || textBoxControlBath1Negative44_5.Text == "+")
            {
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
            {
                if (textBoxControlBath1Negative44_5.Text != StartControlBath1Negative44_5)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        textBoxControlBath1Negative44_5.Text = StartControlBath1Negative44_5;
                    }
                }
            }
        }
        private void textBoxControlBath2Negative44_5_TextChanged(object sender, EventArgs e)
        {
            textBoxControlBath2Negative44_5.ForeColor = Color.Black;
            if (textBoxControlBath2Negative44_5.Text == "-" || textBoxControlBath2Negative44_5.Text == "+")
            {
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
            {
                if (textBoxControlBath2Negative44_5.Text != StartControlBath2Negative44_5)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        textBoxControlBath2Negative44_5.Text = StartControlBath2Negative44_5;
                    }
                }
            }
        }
        private void textBoxControlBath3Negative44_5_TextChanged(object sender, EventArgs e)
        {
            textBoxControlBath3Negative44_5.ForeColor = Color.Black;
            if (textBoxControlBath3Negative44_5.Text == "-" || textBoxControlBath3Negative44_5.Text == "+")
            {
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
            {
                if (textBoxControlBath3Negative44_5.Text != StartControlBath3Negative44_5)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        textBoxControlBath3Negative44_5.Text = StartControlBath3Negative44_5;
                    }
                }
            }
        }
        private void textBoxControlNonTarget35_TextChanged(object sender, EventArgs e)
        {
            textBoxControlNonTarget35.ForeColor = Color.Black;
            if (textBoxControlNonTarget35.Text == "+" || textBoxControlNonTarget35.Text == "-" || textBoxControlNonTarget35.Text.ToUpper() == "N")
            {
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
            {
                if (textBoxControlNonTarget35.Text != StartControlNonTarget35)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        textBoxControlNonTarget35.Text = StartControlNonTarget35;
                    }
                }
            }
        }
        private void textBoxControlBath1NonTarget44_5_TextChanged(object sender, EventArgs e)
        {
            textBoxControlBath1NonTarget44_5.ForeColor = Color.Black;
            if (textBoxControlBath1NonTarget44_5.Text == "-" || textBoxControlBath1NonTarget44_5.Text == "+" || textBoxControlBath1NonTarget44_5.Text.ToUpper() == "N")
            {
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
            {
                if (textBoxControlBath1NonTarget44_5.Text != StartControlBath1NonTarget44_5)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        textBoxControlBath1NonTarget44_5.Text = StartControlBath1NonTarget44_5;
                    }
                }
            }
        }
        private void textBoxControlBath2NonTarget44_5_TextChanged(object sender, EventArgs e)
        {
            textBoxControlBath2NonTarget44_5.ForeColor = Color.Black;
            if (textBoxControlBath2NonTarget44_5.Text == "-" || textBoxControlBath2NonTarget44_5.Text == "+" || textBoxControlBath2NonTarget44_5.Text.ToUpper() == "N")
            {
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
            {
                if (textBoxControlBath2NonTarget44_5.Text != StartControlBath2NonTarget44_5)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        textBoxControlBath2NonTarget44_5.Text = StartControlBath2NonTarget44_5;
                    }
                }
            }
        }
        private void textBoxControlBath3NonTarget44_5_TextChanged(object sender, EventArgs e)
        {
            textBoxControlBath3NonTarget44_5.ForeColor = Color.Black;
            if (textBoxControlBath3NonTarget44_5.Text == "-" || textBoxControlBath3NonTarget44_5.Text == "+" || textBoxControlBath3NonTarget44_5.Text.ToUpper() == "N")
            {
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
            {
                if (textBoxControlBath3NonTarget44_5.Text != StartControlBath3NonTarget44_5)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        textBoxControlBath3NonTarget44_5.Text = StartControlBath3NonTarget44_5;
                    }
                }
            }
        }
        private void textBoxControlPositive35_TextChanged(object sender, EventArgs e)
        {
            textBoxControlPositive35.ForeColor = Color.Black;
            if (textBoxControlPositive35.Text == "+" || textBoxControlPositive35.Text == "-")
            {
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
            {
                if (textBoxControlPositive35.Text != StartControlPositive35)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        textBoxControlPositive35.Text = StartControlPositive35;
                    }
                }
            }
        }
        private void textBoxControlBath1Positive44_5_TextChanged(object sender, EventArgs e)
        {
            textBoxControlBath1Positive44_5.ForeColor = Color.Black;
            if (textBoxControlBath1Positive44_5.Text == "+" || textBoxControlBath1Positive44_5.Text == "-")
            {
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
            {
                if (textBoxControlBath1Positive44_5.Text != StartControlBath1Positive44_5)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        textBoxControlBath1Positive44_5.Text = StartControlBath1Positive44_5;
                    }
                }
            }
        }
        private void textBoxControlBath2Positive44_5_TextChanged(object sender, EventArgs e)
        {
            textBoxControlBath2Positive44_5.ForeColor = Color.Black;
            if (textBoxControlBath2Positive44_5.Text == "+" || textBoxControlBath2Positive44_5.Text == "-")
            {
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
            {
                if (textBoxControlBath2Positive44_5.Text != StartControlBath2Positive44_5)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        textBoxControlBath2Positive44_5.Text = StartControlBath2Positive44_5;
                    }
                }
            }
        }
        private void textBoxControlBath3Positive44_5_TextChanged(object sender, EventArgs e)
        {
            textBoxControlBath3Positive44_5.ForeColor = Color.Black;
            if (textBoxControlBath3Positive44_5.Text == "+" || textBoxControlBath3Positive44_5.Text == "-")
            {
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
            {
                if (textBoxControlBath3Positive44_5.Text != StartControlBath3Positive44_5)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        textBoxControlBath3Positive44_5.Text = StartControlBath3Positive44_5;
                    }
                }
            }
        }
        private void textBoxDailyDuplicatePrecisionCriteria_TextChanged(object sender, EventArgs e)
        {
            CalculateDuplicate();

            if (!InLoadingFile)
            {
                if (textBoxDailyDuplicatePrecisionCriteria.Text != StartDailyDuplicatePrecisionCriteria)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        textBoxDailyDuplicatePrecisionCriteria.Text = StartDailyDuplicatePrecisionCriteria;
                    }
                }
            }
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
                    TryToCalculateIncubationTimeSpan();
                }
            }

            if (!InLoadingFile)
            {
                if (textBoxIncubationBath1StartTime.Text != StartIncubationBath1StartTime)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        textBoxIncubationBath1StartTime.Text = StartIncubationBath1StartTime;
                    }
                }
            }
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
                    TryToCalculateIncubationTimeSpan();
                }
            }

            if (!InLoadingFile)
            {
                if (textBoxIncubationBath2StartTime.Text != StartIncubationBath2StartTime)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        textBoxIncubationBath2StartTime.Text = StartIncubationBath2StartTime;
                    }
                }
            }
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
                    TryToCalculateIncubationTimeSpan();
                }
            }

            if (!InLoadingFile)
            {
                if (textBoxIncubationBath3StartTime.Text != StartIncubationBath3StartTime)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        textBoxIncubationBath3StartTime.Text = StartIncubationBath3StartTime;
                    }
                }
            }
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
                    TryToCalculateIncubationTimeSpan();
                }
            }

            if (!InLoadingFile)
            {
                if (textBoxIncubationBath1EndTime.Text != StartIncubationBath1EndTime)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        textBoxIncubationBath1EndTime.Text = StartIncubationBath1EndTime;
                    }
                }
            }
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
                    TryToCalculateIncubationTimeSpan();
                }
            }

            if (!InLoadingFile)
            {
                if (textBoxIncubationBath2EndTime.Text != StartIncubationBath2EndTime)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        textBoxIncubationBath2EndTime.Text = StartIncubationBath2EndTime;
                    }
                }
            }
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
                    TryToCalculateIncubationTimeSpan();
                }
            }

            if (!InLoadingFile)
            {
                if (textBoxIncubationBath3EndTime.Text != StartIncubationBath3EndTime)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        textBoxIncubationBath3EndTime.Text = StartIncubationBath3EndTime;
                    }
                }
            }
        }
        private void textBoxIntertechDuplicatePrecisionCriteria_TextChanged(object sender, EventArgs e)
        {
            CalculateDuplicate();

            if (!InLoadingFile)
            {
                if (textBoxIntertechDuplicatePrecisionCriteria.Text != StartIntertechDuplicatePrecisionCriteria)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        textBoxIntertechDuplicatePrecisionCriteria.Text = StartIntertechDuplicatePrecisionCriteria;
                    }
                }
            }
        }
        private void textBoxLot35_TextChanged(object sender, EventArgs e)
        {
            if (!InLoadingFile)
            {
                if (textBoxLot35.Text != StartLot35)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        textBoxLot35.Text = StartLot35;
                    }
                }
            }
        }
        private void textBoxLot44_5_TextChanged(object sender, EventArgs e)
        {
            if (!InLoadingFile)
            {
                if (textBoxLot44_5.Text != StartLot44_5)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        textBoxLot44_5.Text = StartLot44_5;
                    }
                }
            }
        }
        private void textBoxResultsReadBy_TextChanged(object sender, EventArgs e)
        {
            if (!InLoadingFile)
            {
                if (textBoxResultsReadBy.Text != StartResultsReadBy)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        textBoxResultsReadBy.Text = StartResultsReadBy;
                    }
                }
            }
        }
        private void textBoxResultsRecordedBy_TextChanged(object sender, EventArgs e)
        {
            if (!InLoadingFile)
            {
                if (textBoxResultsRecordedBy.Text != StartResultsRecordedBy)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        textBoxResultsRecordedBy.Text = StartResultsRecordedBy;
                    }
                }
            }
        }
        private void textBoxSalinitiesReadBy_TextChanged(object sender, EventArgs e)
        {
            if (!InLoadingFile)
            {
                if (textBoxSalinitiesReadBy.Text != StartSalinityReadBy)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        textBoxSalinitiesReadBy.Text = StartSalinityReadBy;
                    }
                }
            }
        }
        private void textBoxSampleBottleLotNumber_TextChanged(object sender, EventArgs e)
        {
            if (!InLoadingFile)
            {
                if (textBoxSampleBottleLotNumber.Text != StartSampleBottleLotNumber)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        textBoxSampleBottleLotNumber.Text = StartSampleBottleLotNumber;
                    }
                }
            }
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
            {
                if (textBoxSampleCrewInitials.Text != StartSampleCrewInitials)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        textBoxSampleCrewInitials.Text = StartSampleCrewInitials;
                    }
                }
            }
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
            {
                if (textBoxTCField1.Text != StartTCField1)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        textBoxTCField1.Text = StartTCField1;
                    }
                }
            }
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
            {
                if (textBoxTCField2.Text != StartTCField2)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        textBoxTCField2.Text = StartTCField2;
                    }
                }
            }
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
            {
                if (textBoxTCLab1.Text != StartTCLab1)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        textBoxTCLab1.Text = StartTCLab1;
                    }
                }
            }
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
            {
                if (textBoxTCLab2.Text != StartTCLab2)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        textBoxTCLab2.Text = StartTCLab2;
                    }
                }
            }
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
            {
                if (textBoxTides.Text != StartTide)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        textBoxTides.Text = StartTide;
                    }
                }
            }
        }
        private void textBoxWaterBath1Number_TextChanged(object sender, EventArgs e)
        {
            if (!InLoadingFile)
            {
                if (textBoxWaterBath1Number.Text != StartWaterBath1Number)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        textBoxWaterBath1Number.Text = StartWaterBath1Number;
                    }
                }
            }
        }
        private void textBoxWaterBath2Number_TextChanged(object sender, EventArgs e)
        {
            if (!InLoadingFile)
            {
                if (textBoxWaterBath2Number.Text != StartWaterBath2Number)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        textBoxWaterBath2Number.Text = StartWaterBath2Number;
                    }
                }
            }
        }
        private void textBoxWaterBath3Number_TextChanged(object sender, EventArgs e)
        {
            if (!InLoadingFile)
            {
                if (textBoxWaterBath3Number.Text != StartWaterBath3Number)
                {
                    if (PreModifying())
                    {
                        IsSaving = true;
                        Modifying();
                    }
                    else
                    {
                        textBoxWaterBath3Number.Text = StartWaterBath3Number;
                    }
                }
            }
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
                        }
                        return;
                    }
                    timerGetTides.Enabled = true;
                }
                Modifying();
                if (labSheetA1Sheet.Tides != textBoxTides.Text)
                {
                    labSheetA1Sheet.Tides = textBoxTides.Text;
                }
            }
        }
        #endregion Events WebBrowserCSSP

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
    }


}




