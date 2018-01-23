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
    public partial class CSSPWQInputToolForm
    {
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
                    if (comboBoxSubsectorNames.SelectedIndex != 0)
                    {
                        butGetLabSheetsStatus.Enabled = true;
                    }
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

            StartTide = textBoxTides.Text;
            StartSampleCrewInitials = textBoxSampleCrewInitials.Text;
            Start2Coolers = checkBox2Coolers.Checked.ToString().ToLower();
            StartTCField1 = textBoxTCField1.Text;
            StartTCLab1 = textBoxTCLab1.Text;
            StartTCField2 = textBoxTCField2.Text;
            StartTCLab2 = textBoxTCLab2.Text;
            StartIncubationStartSameDay = checkBoxIncubationStartSameDay.Checked.ToString().ToLower();
            if (radioButton2Baths.Checked)
            {
                StartNumberOfBaths = "2";
            }
            else if (radioButton3Baths.Checked)
            {
                StartNumberOfBaths = "3";
            }
            else
            {
                StartNumberOfBaths = "1";
            }
            StartIncubationBath1StartTime = textBoxIncubationBath1StartTime.Text;
            StartIncubationBath1EndTime = textBoxIncubationBath1EndTime.Text;
            StartWaterBath1Number = textBoxWaterBath1Number.Text;

            if (radioButton2Baths.Checked || radioButton3Baths.Checked)
            {
                StartIncubationBath2StartTime = textBoxIncubationBath2StartTime.Text;
                StartIncubationBath2EndTime = textBoxIncubationBath2EndTime.Text;
                StartWaterBath2Number = textBoxWaterBath2Number.Text;
            }
            else if (radioButton3Baths.Checked)
            {
                StartIncubationBath3StartTime = textBoxIncubationBath3StartTime.Text;
                StartIncubationBath3EndTime = textBoxIncubationBath3EndTime.Text;
                StartWaterBath3Number = textBoxWaterBath3Number.Text;
            }

            StartControlLot = textBoxControlLot.Text;
            StartControlPositive35 = textBoxControlPositive35.Text;
            StartControlNonTarget35 = textBoxControlNonTarget35.Text;
            StartControlNegative35 = textBoxControlNegative35.Text;
            StartControlBlank35 = textBoxControlBlank35.Text;

            StartControlBath1Positive44_5 = textBoxControlBath1Positive44_5.Text;
            StartControlBath1NonTarget44_5 = textBoxControlBath1NonTarget44_5.Text;
            StartControlBath1Negative44_5 = textBoxControlBath1Negative44_5.Text;
            StartControlBath1Blank44_5 = textBoxControlBath1Blank44_5.Text;

            if (radioButton2Baths.Checked || radioButton3Baths.Checked)
            {
                StartControlBath2Positive44_5 = textBoxControlBath2Positive44_5.Text;
                StartControlBath2NonTarget44_5 = textBoxControlBath2NonTarget44_5.Text;
                StartControlBath2Negative44_5 = textBoxControlBath2Negative44_5.Text;
                StartControlBath2Blank44_5 = textBoxControlBath2Blank44_5.Text;
            }

            if (radioButton3Baths.Checked)
            {
                StartControlBath3Positive44_5 = textBoxControlBath3Positive44_5.Text;
                StartControlBath3NonTarget44_5 = textBoxControlBath3NonTarget44_5.Text;
                StartControlBath3Negative44_5 = textBoxControlBath3Negative44_5.Text;
                StartControlBath3Blank44_5 = textBoxControlBath3Blank44_5.Text;
            }

            StartLot35 = textBoxLot35.Text;
            StartLot44_5 = textBoxLot44_5.Text;
            StartRunWeatherComment = richTextBoxRunWeatherComment.Text;
            StartRunComment = richTextBoxRunComment.Text;

            StartSampleBottleLotNumber = textBoxSampleBottleLotNumber.Text;

            StartSalinityReadBy = textBoxSalinitiesReadBy.Text;
            StartSalinityReadDateYear = dateTimePickerSalinitiesReadDate.Value.Year.ToString();
            StartSalinityReadDateMonth = dateTimePickerSalinitiesReadDate.Value.Month.ToString();
            StartSalinityReadDateDay = dateTimePickerSalinitiesReadDate.Value.Day.ToString();

            StartResultsReadBy = textBoxResultsReadBy.Text;
            StartResultsReadDateYear = dateTimePickerResultsReadDate.Value.Year.ToString();
            StartResultsReadDateMonth = dateTimePickerResultsReadDate.Value.Month.ToString();
            StartResultsReadDateDay = dateTimePickerResultsReadDate.Value.Day.ToString();

            StartResultsRecordedBy = textBoxResultsRecordedBy.Text;
            StartResultsRecordedDateYear = dateTimePickerResultsRecordedDate.Value.Year.ToString();
            StartResultsRecordedDateMonth = dateTimePickerResultsRecordedDate.Value.Month.ToString();
            StartResultsRecordedDateDay = dateTimePickerResultsRecordedDate.Value.Day.ToString();

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
            for (int i = 1; i < 11; i++)
            {
                comboBoxRunNumber.Items.Add(new RunNumberAndText() { RunNumber = i, RunNumberText = $"Run { i } on this date" });
            }

            comboBoxRunNumber.DisplayMember = "RunNumberText";
            comboBoxRunNumber.ValueMember = "RunNumber";

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
        private void LogAll()
        {
            if (StartTide != textBoxTides.Text)
            {
                AddLog("Tides", textBoxTides.Text);
            }

            if (StartSampleCrewInitials != textBoxSampleCrewInitials.Text)
            {
                AddLog("Sample Crew Initials", textBoxSampleCrewInitials.Text);
            }

            if (Start2Coolers.ToLower() != checkBox2Coolers.Checked.ToString().ToLower())
            {
                AddLog("2 Coolers", checkBox2Coolers.Checked.ToString().ToLower());
            }

            if (StartTCField1 != textBoxTCField1.Text)
            {
                AddLog("TC Field #1", textBoxTCField1.Text);
            }

            if (StartTCLab1 != textBoxTCLab1.Text)
            {
                AddLog("TC Lab #1", textBoxTCLab1.Text);
            }

            if (checkBox2Coolers.Checked)
            {
                if (StartTCField2 != textBoxTCField2.Text)
                {
                    AddLog("TC Field #2", textBoxTCField2.Text);
                }

                if (StartTCLab2 != textBoxTCLab2.Text)
                {
                    AddLog("TC Lab #2", textBoxTCLab2.Text);
                }
            }

            if (StartIncubationStartSameDay.ToLower() != checkBoxIncubationStartSameDay.Checked.ToString().ToLower())
            {
                AddLog("Incubation Start Same Day", checkBoxIncubationStartSameDay.Checked.ToString().ToLower());
            }

            if (radioButton2Baths.Checked)
            {
                if (StartNumberOfBaths != "2")
                {
                    AddLog("Water Bath Count", "2");
                }
            }
            else if (radioButton3Baths.Checked)
            {
                if (StartNumberOfBaths != "3")
                {
                    AddLog("Water Bath Count", "3");
                }
            }
            else
            {
                if (StartNumberOfBaths != "1")
                {
                    AddLog("Water Bath Count", "1");
                }
            }

            if (StartIncubationBath1StartTime != textBoxIncubationBath1StartTime.Text)
            {
                AddLog("Incubation Bath 1 Start Time", textBoxIncubationBath1StartTime.Text);
            }

            if (StartIncubationBath1EndTime != textBoxIncubationBath1EndTime.Text)
            {
                AddLog("Incubation Bath 1 End Time", textBoxIncubationBath1EndTime.Text);
            }

            if (StartWaterBath1Number != textBoxWaterBath1Number.Text)
            {
                AddLog("Water Bath 1", textBoxWaterBath1Number.Text);
            }

            if (radioButton2Baths.Checked || radioButton3Baths.Checked)
            {
                if (StartIncubationBath2StartTime != textBoxIncubationBath2StartTime.Text)
                {
                    AddLog("Incubation Bath 2 Start Time", textBoxIncubationBath2StartTime.Text);
                }

                if (StartIncubationBath2EndTime != textBoxIncubationBath2EndTime.Text)
                {
                    AddLog("Incubation Bath 2 End Time", textBoxIncubationBath2EndTime.Text);
                }

                if (StartWaterBath2Number != textBoxWaterBath2Number.Text)
                {
                    AddLog("Water Bath 2", textBoxWaterBath2Number.Text);
                }
            }

            if (radioButton3Baths.Checked)
            {
                if (StartIncubationBath3StartTime != textBoxIncubationBath3StartTime.Text)
                {
                    AddLog("Incubation Bath 3 Start Time", textBoxIncubationBath3StartTime.Text);
                }

                if (StartIncubationBath3EndTime != textBoxIncubationBath3EndTime.Text)
                {
                    AddLog("Incubation Bath 3 End Time", textBoxIncubationBath3EndTime.Text);
                }

                if (StartWaterBath3Number != textBoxWaterBath3Number.Text)
                {
                    AddLog("Water Bath 3", textBoxWaterBath3Number.Text);
                }
            }

            if (StartControlLot != textBoxControlLot.Text)
            {
                AddLog("Control Lot", textBoxControlLot.Text);
            }

            if (StartControlPositive35 != textBoxControlPositive35.Text)
            {
                AddLog("Positive 35", textBoxControlPositive35.Text);
            }

            if (StartControlNonTarget35 != textBoxControlNonTarget35.Text)
            {
                AddLog("Non Target 35", textBoxControlNonTarget35.Text);
            }

            if (StartControlNegative35 != textBoxControlNegative35.Text)
            {
                AddLog("Negative 35", textBoxControlNegative35.Text);
            }

            if (StartControlBlank35 != textBoxControlBlank35.Text)
            {
                AddLog("Control Blank 35", textBoxControlBlank35.Text);
            }

            // Bath #1
            if (StartControlBath1Positive44_5 != textBoxControlBath1Positive44_5.Text)
            {
                AddLog("Bath 1 Positive 44.5", textBoxControlBath1Positive44_5.Text);
            }

            if (StartControlBath1NonTarget44_5 != textBoxControlBath1NonTarget44_5.Text)
            {
                AddLog("Bath 1 Non Target 44.5", textBoxControlBath1NonTarget44_5.Text);
            }

            if (StartControlBath1Negative44_5 != textBoxControlBath1Negative44_5.Text)
            {
                AddLog("Bath1 Negative 44.5", textBoxControlBath1Negative44_5.Text);
            }

            if (StartControlBath1Blank44_5 != textBoxControlBath1Blank44_5.Text)
            {
                AddLog("Control Bath 1 Blank 44.5", textBoxControlBath1Blank44_5.Text);
            }

            if (radioButton2Baths.Checked || radioButton3Baths.Checked)
            {
                // Bath #2
                if (StartControlBath2Positive44_5 != textBoxControlBath2Positive44_5.Text)
                {
                    AddLog("Bath 2 Positive 44.5", textBoxControlBath2Positive44_5.Text);
                }

                if (StartControlBath2NonTarget44_5 != textBoxControlBath2NonTarget44_5.Text)
                {
                    AddLog("Bath 2 Non Target 44.5", textBoxControlBath2NonTarget44_5.Text);
                }

                if (StartControlBath2Negative44_5 != textBoxControlBath2Negative44_5.Text)
                {
                    AddLog("Bath2 Negative 44.5", textBoxControlBath2Negative44_5.Text);
                }

                if (StartControlBath2Blank44_5 != textBoxControlBath2Blank44_5.Text)
                {
                    AddLog("Control Bath 2 Blank 44.5", textBoxControlBath2Blank44_5.Text);
                }
            }

            if (radioButton3Baths.Checked)
            {
                // Bath #3
                if (StartControlBath3Positive44_5 != textBoxControlBath3Positive44_5.Text)
                {
                    AddLog("Bath 3 Positive 44.5", textBoxControlBath3Positive44_5.Text);
                }

                if (StartControlBath3NonTarget44_5 != textBoxControlBath3NonTarget44_5.Text)
                {
                    AddLog("Bath 3 Non Target 44.5", textBoxControlBath3NonTarget44_5.Text);
                }

                if (StartControlBath3Negative44_5 != textBoxControlBath3Negative44_5.Text)
                {
                    AddLog("Bath3 Negative 44.5", textBoxControlBath3Negative44_5.Text);
                }

                if (StartControlBath3Blank44_5 != textBoxControlBath3Blank44_5.Text)
                {
                    AddLog("Control Bath 3 Blank 44.5", textBoxControlBath3Blank44_5.Text);
                }
            }

            if (StartLot35 != textBoxLot35.Text)
            {
                AddLog("Lot 35", textBoxLot35.Text);
            }

            if (StartLot44_5 != textBoxLot44_5.Text)
            {
                AddLog("Lot 44.5", textBoxLot44_5.Text);
            }

            if (StartRunWeatherComment != richTextBoxRunWeatherComment.Text)
            {
                AddLog("Run Weather Comment", richTextBoxRunWeatherComment.Text);
            }

            if (StartRunComment != richTextBoxRunComment.Text)
            {
                AddLog("Run Comment", richTextBoxRunComment.Text);
            }

            if (StartSampleBottleLotNumber != textBoxSampleBottleLotNumber.Text)
            {
                AddLog("Sample Bottle Lot Number", textBoxSampleBottleLotNumber.Text);
            }

            if (StartSalinityReadBy != textBoxSalinitiesReadBy.Text)
            {
                AddLog("Salinities Read By", textBoxSalinitiesReadBy.Text);
            }

            if (StartSalinityReadDateYear != dateTimePickerSalinitiesReadDate.Value.Year.ToString()
                || StartSalinityReadDateMonth != dateTimePickerSalinitiesReadDate.Value.Month.ToString()
                || StartSalinityReadDateDay != dateTimePickerSalinitiesReadDate.Value.Day.ToString())
            {
                AddLog("Results Salinities Date", dateTimePickerSalinitiesReadDate.Value.Year.ToString() + 
                    "\t" + dateTimePickerSalinitiesReadDate.Value.Month.ToString() + 
                    "\t" + dateTimePickerSalinitiesReadDate.Value.Day.ToString());
            }

            if (StartResultsReadBy != textBoxResultsReadBy.Text)
            {
                AddLog("Results Read By", textBoxResultsReadBy.Text);
            }

            if (StartResultsReadDateYear != dateTimePickerResultsReadDate.Value.Year.ToString()
                || StartResultsReadDateMonth != dateTimePickerResultsReadDate.Value.Month.ToString()
                || StartResultsReadDateDay != dateTimePickerResultsReadDate.Value.Day.ToString())
            {
                AddLog("Results Results Date", dateTimePickerResultsReadDate.Value.Year.ToString() + 
                    "\t" + dateTimePickerResultsReadDate.Value.Month.ToString() + 
                    "\t" + dateTimePickerResultsReadDate.Value.Day.ToString());
            }

            if (StartResultsRecordedBy != textBoxResultsRecordedBy.Text)
            {
                AddLog("Results Recorded By", textBoxResultsRecordedBy.Text);
            }

            if (StartResultsRecordedDateYear != dateTimePickerResultsRecordedDate.Value.Year.ToString()
                || StartResultsRecordedDateMonth != dateTimePickerResultsRecordedDate.Value.Month.ToString()
                || StartResultsRecordedDateDay != dateTimePickerResultsRecordedDate.Value.Day.ToString())
            {
                AddLog("Results Results Date", dateTimePickerResultsRecordedDate.Value.Year.ToString() +
                    "\t" + dateTimePickerResultsRecordedDate.Value.Month.ToString() +
                    "\t" + dateTimePickerResultsRecordedDate.Value.Day.ToString());
            }

            int SiteColumn = 1;
            int SampleTypeColumn = 10;

            for (int row = 0, countRow = dataGridViewCSSP.Rows.Count; row < countRow; row++)
            {
                for (int col = 0, countCol = dataGridViewCSSP.Columns.Count; col < countCol; col++)
                {
                    switch (col)
                    {
                        //case 1:
                        //    {
                        //        if (StartGridCellText[row].Site != (dataGridViewCSSP[col, row].Value == null ? "" : dataGridViewCSSP[col, row].Value.ToString()))
                        //        {
                        //            AddLog("CSSP Grid(" + col + "," + row + ") " + dataGridViewCSSP[SiteColumn, row].Value.ToString() + (dataGridViewCSSP[SampleTypeColumn, row].Value.ToString() == "Daily Duplicate" ? " Daily Duplicate" : "") + " - " + " [Processed By] ", dataGridViewCSSP[col, row].Value.ToString());
                        //        }
                        //    }
                        //    break;
                        case 2:
                            {
                                if (StartGridCellText[row].Time != (dataGridViewCSSP[col, row].Value == null ? "" : dataGridViewCSSP[col, row].Value.ToString()))
                                {
                                    AddLog("CSSP Grid(" + col + "," + row + ") " + dataGridViewCSSP[SiteColumn, row].Value.ToString() + " " + dataGridViewCSSP[SampleTypeColumn, row].Value.ToString() + " - " + " [Time] ", dataGridViewCSSP[col, row].Value.ToString());
                                }
                            }
                            break;
                        //case 3:
                        //    {
                        //        if (StartGridCellText[row].MPN != (dataGridViewCSSP[col, row].Value == null ? "" : dataGridViewCSSP[col, row].Value.ToString()))
                        //        {
                        //            AddLog("CSSP Grid(" + col + "," + row + ") " + dataGridViewCSSP[SiteColumn, row].Value.ToString() + " " + dataGridViewCSSP[SampleTypeColumn, row].Value.ToString() + " - " + " [MPN] ", dataGridViewCSSP[col, row].Value.ToString());
                        //        }
                        //    }
                        //    break;
                        case 4:
                            {
                                if (StartGridCellText[row].Tube10 != (dataGridViewCSSP[col, row].Value == null ? "" : dataGridViewCSSP[col, row].Value.ToString()))
                                {
                                    AddLog("CSSP Grid(" + col + "," + row + ") " + dataGridViewCSSP[SiteColumn, row].Value.ToString() + " " + dataGridViewCSSP[SampleTypeColumn, row].Value.ToString() + " - " + " [Tube 10] ", dataGridViewCSSP[col, row].Value.ToString());
                                }
                            }
                            break;
                        case 5:
                            {
                                if (StartGridCellText[row].Tube1_0 != (dataGridViewCSSP[col, row].Value == null ? "" : dataGridViewCSSP[col, row].Value.ToString()))
                                {
                                    AddLog("CSSP Grid(" + col + "," + row + ") " + dataGridViewCSSP[SiteColumn, row].Value.ToString() + " " + dataGridViewCSSP[SampleTypeColumn, row].Value.ToString() + " - " + " [Tube 1] ", dataGridViewCSSP[col, row].Value.ToString());
                                }
                            }
                            break;
                        case 6:
                            {
                                if (StartGridCellText[row].Tube1_0 != (dataGridViewCSSP[col, row].Value == null ? "" : dataGridViewCSSP[col, row].Value.ToString()))
                                {
                                    AddLog("CSSP Grid(" + col + "," + row + ") " + dataGridViewCSSP[SiteColumn, row].Value.ToString() + " " + dataGridViewCSSP[SampleTypeColumn, row].Value.ToString() + " - " + " [Tube 0.1] ", dataGridViewCSSP[col, row].Value.ToString());
                                }
                            }
                            break;
                        case 7:
                            {
                                if (StartGridCellText[row].Sal != (dataGridViewCSSP[col, row].Value == null ? "" : dataGridViewCSSP[col, row].Value.ToString()))
                                {
                                    AddLog("CSSP Grid(" + col + "," + row + ") " + dataGridViewCSSP[SiteColumn, row].Value.ToString() + " " + dataGridViewCSSP[SampleTypeColumn, row].Value.ToString() + " - " + " [Sal] ", dataGridViewCSSP[col, row].Value.ToString());
                                }
                            }
                            break;
                        case 8:
                            {
                                if (StartGridCellText[row].Temp != (dataGridViewCSSP[col, row].Value == null ? "" : dataGridViewCSSP[col, row].Value.ToString()))
                                {
                                    AddLog("CSSP Grid(" + col + "," + row + ") " + dataGridViewCSSP[SiteColumn, row].Value.ToString() + " " + dataGridViewCSSP[SampleTypeColumn, row].Value.ToString() + " - " + " [Temp] ", dataGridViewCSSP[col, row].Value.ToString());
                                }
                            }
                            break;
                        case 9:
                            {
                                if (StartGridCellText[row].ProcessBy != (dataGridViewCSSP[col, row].Value == null ? "" : dataGridViewCSSP[col, row].Value.ToString()))
                                {
                                    AddLog("CSSP Grid(" + col + "," + row + ") " + dataGridViewCSSP[SiteColumn, row].Value.ToString() + " " + dataGridViewCSSP[SampleTypeColumn, row].Value.ToString() + " - " + " [Process by] ", dataGridViewCSSP[col, row].Value.ToString());
                                }
                            }
                            break;
                        //case 10:
                        //    {
                        //        if (StartGridCellText[row].SampleType != (dataGridViewCSSP[col, row].Value == null ? "" : dataGridViewCSSP[col, row].Value.ToString()))
                        //        {
                        //            AddLog("CSSP Grid(" + col + "," + row + ") " + dataGridViewCSSP[SiteColumn, row].Value.ToString() + " " + dataGridViewCSSP[SampleTypeColumn, row].Value.ToString() + " - " + " [Sample type] ", dataGridViewCSSP[col, row].Value.ToString());
                        //        }
                        //    }
                        //    break;
                        //case 11:
                        //    {
                        //        if (StartGridCellText[row].ID != (dataGridViewCSSP[col, row].Value == null ? "" : dataGridViewCSSP[col, row].Value.ToString()))
                        //        {
                        //            AddLog("CSSP Grid(" + col + "," + row + ") " + dataGridViewCSSP[SiteColumn, row].Value.ToString() + " " + dataGridViewCSSP[SampleTypeColumn, row].Value.ToString() + " - " + " [TVItemID] ", dataGridViewCSSP[col, row].Value.ToString());
                        //        }
                        //    }
                        //    break;
                        case 12:
                            {
                                if (StartGridCellText[row].Comment != (dataGridViewCSSP[col, row].Value == null ? "" : dataGridViewCSSP[col, row].Value.ToString()))
                                {
                                    AddLog("CSSP Grid(" + col + "," + row + ") " + dataGridViewCSSP[SiteColumn, row].Value.ToString() + " " + dataGridViewCSSP[SampleTypeColumn, row].Value.ToString() + " - " + " [Comment] ", dataGridViewCSSP[col, row].Value.ToString());
                                }
                            }
                            break;
                        default:
                            break;
                    }
                }
            }

            DoSave();
        }

        // more functions in Functions_M_Z.cs
    }
}
