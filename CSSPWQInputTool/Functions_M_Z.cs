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
        // more functions in Functions_A_L.cs

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

            lblStatus.Text = "Modified";
            butSendToEnvironmentCanada.Text = "Saving ...";
            butSendToEnvironmentCanada.Enabled = false;
            butGetTides.Enabled = false;
            butViewFCForm.Enabled = false;
            if (!timerSave.Enabled)
            {
                timerSave.Enabled = true;
                timerSave.Start();
            }
        }
        private bool PreModifying()
        {
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
                    return false;
            }

            return true;
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
                butDeleteLabSheet.Enabled = false;
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

            foreach (RunNumberAndText runNumberAndText in comboBoxRunNumber.Items)
            {
                string r = (runNumberAndText.RunNumber < 10 ? "0" + runNumberAndText.RunNumber : runNumberAndText.RunNumber.ToString());

                if ("R" + r == RunText)
                {
                    comboBoxRunNumber.SelectedValue = runNumberAndText;
                    break;
                }
            }

            DateTime dateTimeRun = new DateTime(int.Parse(Year), int.Parse(Month), int.Parse(Day));

            dateTimePickerRun.Value = dateTimeRun;

            NoUpdate = false;

            UpdatePanelApp();

            OpenedFileName = lblFilePath.Text;

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
        }
        private bool ReadSamplingPlan()
        {
            InLoadingFile = true;

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
                    case "Backup Directory":
                        {
                            BackupDirectory = LineTxt.Substring("Backup Directory\t".Length).Trim();
                            textBoxSharedArchivedDirectory.Text = BackupDirectory;
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
                butSendToEnvironmentCanada.Text = "Already saved on server";
                butSendToEnvironmentCanada.Enabled = false;
            }
            else if (fi.FullName.EndsWith("_R.txt"))
            {
                butSendToEnvironmentCanada.Text = "Rejected on server";
                butSendToEnvironmentCanada.Enabled = true;
            }
            else if (fi.FullName.EndsWith("_A.txt"))
            {
                butSendToEnvironmentCanada.Text = "Accepted on server";
                butSendToEnvironmentCanada.Enabled = false;
            }
            else
            {
                butSendToEnvironmentCanada.Text = "Send to Environment Canada";
                butSendToEnvironmentCanada.Enabled = true;
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
                butSendToEnvironmentCanada.Text = "No Internet Connection";
                butSendToEnvironmentCanada.Enabled = false;
                MessageBox.Show("No internet connection", "Internet connection");
                return;
            }

            if (!EverythingEntered())
            {
                return;
            }
            butSendToEnvironmentCanada.Text = "Working ...";
            butSendToEnvironmentCanada.Enabled = false;
            lblStatus.Text = "Sending lab sheet to server ... Working ...";
            lblStatus.Refresh();
            Application.DoEvents();
            string retStr = PostLabSheet();
            if (string.IsNullOrWhiteSpace(retStr))
            {
                butSendToEnvironmentCanada.Text = "Lab sheet sent ok";
                butSendToEnvironmentCanada.Enabled = false;
                lblStatus.Text = "Lab sheet sent ok";

                File.Copy(lblFilePath.Text, lblFilePath.Text.Replace("_C.txt", "_S.txt"));
                File.Delete(lblFilePath.Text);
                OpenedFileName = lblFilePath.Text.Replace("_C.txt", "_S.txt");
                lblFilePath.Text = lblFilePath.Text.Replace("_C.txt", "_S.txt");
            }
            else
            {
                butSendToEnvironmentCanada.Text = "Error sending lab sheet";
                if (InternetConnection)
                {
                    butSendToEnvironmentCanada.Enabled = true;
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
            butHome.Enabled = false;
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
            butSendToEnvironmentCanada.Enabled = false;
        }
        private void SetupCSSPWQInputTool()
        {
            panelAddInputBottomLeftDuplicate.Visible = false;
            CreateCSSPSamplingPlanFilePath();
            textBoxAccessCode.Text = "";
            panelApp.BringToFront();
            CurrentPanel = panelApp;
            panelButtonBar.Visible = true;
            FillInternetConnectionVariable();
            lblFilePath.Text = "";
            butCreateLabSheet.Visible = false;
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

            StartGridCellText = new List<List<string>>();
            int SampleTypeColNumber = 10;
            for (int row = 0, countRow = dataGridViewCSSP.Rows.Count; row < countRow; row++)
            {
                List<string> gridCellTextList = new List<string>();

                for (int col = 0, countCol = dataGridViewCSSP.Columns.Count; col < countCol; col++)
                {
                    switch (col)
                    {
                        case 0:
                            {
                                gridCellTextList.Add("");
                            }
                            break;
                        case 1:
                            {
                                gridCellTextList.Add(dataGridViewCSSP[col, row].Value == null ? "" : dataGridViewCSSP[col, row].Value.ToString());
                            }
                            break;
                        case 2:
                            {
                                if (dataGridViewCSSP[SampleTypeColNumber, row].Value != null
                                    && (dataGridViewCSSP[SampleTypeColNumber, row].Value.ToString() == SampleTypeEnum.DailyDuplicate.ToString()
                                    || dataGridViewCSSP[SampleTypeColNumber, row].Value.ToString() == SampleTypeEnum.IntertechDuplicate.ToString()
                                    || dataGridViewCSSP[SampleTypeColNumber, row].Value.ToString() == SampleTypeEnum.IntertechRead.ToString()))
                                {
                                    dataGridViewCSSP[col, row].Style.BackColor = Color.Gray;
                                    dataGridViewCSSP[col, row].ReadOnly = true;
                                }
                                gridCellTextList.Add(dataGridViewCSSP[col, row].Value == null ? "" : dataGridViewCSSP[col, row].Value.ToString());
                            }
                            break;
                        case 3:
                            {
                                gridCellTextList.Add(dataGridViewCSSP[col, row].Value == null ? "" : dataGridViewCSSP[col, row].Value.ToString());
                            }
                            break;
                        case 4:
                            {
                                gridCellTextList.Add(dataGridViewCSSP[col, row].Value == null ? "" : dataGridViewCSSP[col, row].Value.ToString());
                            }
                            break;
                        case 5:
                            {
                                gridCellTextList.Add(dataGridViewCSSP[col, row].Value == null ? "" : dataGridViewCSSP[col, row].Value.ToString());
                            }
                            break;
                        case 6:
                            {
                                gridCellTextList.Add(dataGridViewCSSP[col, row].Value == null ? "" : dataGridViewCSSP[col, row].Value.ToString());
                            }
                            break;
                        case 7:
                            {
                                if (dataGridViewCSSP[SampleTypeColNumber, row].Value != null
                                    && (dataGridViewCSSP[SampleTypeColNumber, row].Value.ToString() == SampleTypeEnum.DailyDuplicate.ToString()
                                    || dataGridViewCSSP[SampleTypeColNumber, row].Value.ToString() == SampleTypeEnum.IntertechDuplicate.ToString()
                                    || dataGridViewCSSP[SampleTypeColNumber, row].Value.ToString() == SampleTypeEnum.IntertechRead.ToString()))
                                {
                                    dataGridViewCSSP[col, row].Style.BackColor = Color.Gray;
                                    dataGridViewCSSP[col, row].ReadOnly = true;
                                }
                                gridCellTextList.Add(dataGridViewCSSP[col, row].Value == null ? "" : dataGridViewCSSP[col, row].Value.ToString());
                            }
                            break;
                        case 8:
                            {
                                if (dataGridViewCSSP[SampleTypeColNumber, row].Value != null
                                    && (dataGridViewCSSP[SampleTypeColNumber, row].Value.ToString() == SampleTypeEnum.DailyDuplicate.ToString()
                                    || dataGridViewCSSP[SampleTypeColNumber, row].Value.ToString() == SampleTypeEnum.IntertechDuplicate.ToString()
                                    || dataGridViewCSSP[SampleTypeColNumber, row].Value.ToString() == SampleTypeEnum.IntertechRead.ToString()))
                                {
                                    dataGridViewCSSP[col, row].Style.BackColor = Color.Gray;
                                    dataGridViewCSSP[col, row].ReadOnly = true;
                                }
                                gridCellTextList.Add(dataGridViewCSSP[col, row].Value == null ? "" : dataGridViewCSSP[col, row].Value.ToString());
                            }
                            break;
                        case 9:
                            {
                                //if (dataGridViewCSSP[SampleTypeColNumber, row].Value != null
                                //    && (dataGridViewCSSP[SampleTypeColNumber, row].Value.ToString() == SampleTypeEnum.DailyDuplicate.ToString()
                                //    || dataGridViewCSSP[SampleTypeColNumber, row].Value.ToString() == SampleTypeEnum.IntertechDuplicate.ToString()
                                //    || dataGridViewCSSP[SampleTypeColNumber, row].Value.ToString() == SampleTypeEnum.IntertechRead.ToString()))
                                //{
                                //    dataGridViewCSSP[col, row].Style.BackColor = Color.Gray;
                                //    dataGridViewCSSP[col, row].ReadOnly = true;
                                //}
                                gridCellTextList.Add(dataGridViewCSSP[col, row].Value == null ? "" : dataGridViewCSSP[col, row].Value.ToString());
                            }
                            break;
                        case 10:
                            {
                                gridCellTextList.Add(dataGridViewCSSP[col, row].Value == null ? "" : dataGridViewCSSP[col, row].Value.ToString());
                            }
                            break;
                        case 11:
                            {
                                gridCellTextList.Add(dataGridViewCSSP[col, row].Value == null ? "" : dataGridViewCSSP[col, row].Value.ToString());
                            }
                            break;
                        case 12:
                            {
                                gridCellTextList.Add(dataGridViewCSSP[col, row].Value == null ? "" : dataGridViewCSSP[col, row].Value.ToString());
                            }
                            break;
                        default:
                            break;
                    }

                }

                StartGridCellText.Add(gridCellTextList);
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
            butSendToEnvironmentCanada.Enabled = false;
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

            int rn = ((RunNumberAndText)comboBoxRunNumber.SelectedItem).RunNumber;
            RunNumberCurrent = (rn < 10 ? "0" + rn : rn.ToString());

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
                butDeleteLabSheet.Enabled = true;
                lblFilePath.Text = fi.FullName;
                if (!ReadFileFromLocalMachine())
                    return;
                panelAppInput.BringToFront();
                butHome.Enabled = true;
                CurrentPanel = panelAppInput;
                panelAppInputIsVisible = true;
                butSendToEnvironmentCanada.Enabled = false;

                if (fi.FullName.EndsWith("_S.txt"))
                {
                    butSendToEnvironmentCanada.Text = "Already saved on server";
                    butSendToEnvironmentCanada.Enabled = false;
                }
                else if (fi.FullName.EndsWith("_R.txt"))
                {
                    butSendToEnvironmentCanada.Text = "Rejected on server";
                    butSendToEnvironmentCanada.Enabled = false;
                }
                else if (fi.FullName.EndsWith("_A.txt"))
                {
                    butSendToEnvironmentCanada.Text = "Accepted on server";
                    butSendToEnvironmentCanada.Enabled = false;
                }
                else if (fi.FullName.EndsWith("_E.txt"))
                {
                    butSendToEnvironmentCanada.Text = "Error no server action";
                    butSendToEnvironmentCanada.Enabled = false;
                }
                else if (fi.FullName.EndsWith("_F.txt"))
                {
                    butSendToEnvironmentCanada.Text = "Fail no server action";
                    butSendToEnvironmentCanada.Enabled = false;
                }
                else
                {
                    butSendToEnvironmentCanada.Text = "Send to Environment Canada";
                    butSendToEnvironmentCanada.Enabled = true;
                }

                lblStatus.Text = "";
                butCreateLabSheet.Visible = false;
            }
            else
            {
                SetupDataGridViewCSSP();
                butCreateLabSheet.Visible = true;
                SetupAppInputFiles();
            }

            InLoadingFile = false;
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
                        if (dataGridViewCSSP[ColumnIndex, RowIndex].Value != null)
                        {
                            foreach (char c in dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString())
                            {
                                if (!(char.IsLetter(c)))
                                {
                                    dataGridViewCSSP[ColumnIndex, RowIndex].Value = "";
                                    return;
                                }
                            }
                        }

                        int SiteColumn = 1;
                        int SampleTypeColumn = 10;
                        if (dataGridViewCSSP[ColumnIndex, RowIndex].Value == null || string.IsNullOrWhiteSpace(dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString()))
                            return;

                        dataGridViewCSSP[ColumnIndex, RowIndex].Value = dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString().ToUpper();

                        string SiteName = dataGridViewCSSP[SiteColumn, RowIndex].Value.ToString();
                        for (int i = RowIndex + 1, countRow = dataGridViewCSSP.Rows.Count; i < countRow; i++)
                        {
                            if (dataGridViewCSSP[SiteColumn, i].Value.ToString() == SiteName)
                            {
                                if (dataGridViewCSSP[SampleTypeColumn, i].Value.ToString() == SampleTypeEnum.DailyDuplicate.ToString())
                                {
                                    dataGridViewCSSP[ColumnIndex, i].Value = dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString();
                                }
                            }
                        }

                        if (labSheetA1Sheet.LabSheetA1MeasurementList[RowIndex].ProcessedBy != dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString())
                        {
                            labSheetA1Sheet.LabSheetA1MeasurementList[RowIndex].ProcessedBy = dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString();
                        }

                    }
                    break;
                case 10:
                    {
                    }
                    break;
                case 12:
                    {
                        //int SiteColumn = 1;
                        //int SampleTypeColumn = 10;

                        if (labSheetA1Sheet.LabSheetA1MeasurementList[RowIndex].SiteComment != null && dataGridViewCSSP[ColumnIndex, RowIndex].Value == null)
                        {
                            labSheetA1Sheet.LabSheetA1MeasurementList[RowIndex].SiteComment = "";
                        }
                        else
                        {
                            if (labSheetA1Sheet.LabSheetA1MeasurementList[RowIndex].SiteComment != dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString())
                            {
                                labSheetA1Sheet.LabSheetA1MeasurementList[RowIndex].SiteComment = dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString();
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
            if (dataGridViewCSSP[ColumnIndex, RowIndex].Value != null)
            {
                if (dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString().Length != 1)
                {
                    dataGridViewCSSP[ColumnIndex, RowIndex].Value = "";
                    return;
                }

                foreach (char c in dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString())
                {
                    if (!char.IsDigit(c))
                    {
                        dataGridViewCSSP[ColumnIndex, RowIndex].Value = "";
                        return;
                    }
                }

                int theVal = -1;
                if (int.TryParse(dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString(), out theVal))
                {
                    if (!(theVal >= 0 && theVal <= 5))
                    {
                        dataGridViewCSSP[ColumnIndex, RowIndex].Value = "";
                        return;
                    }
                }
            }

            int MPNColumn = 3;
            //int SiteColumn = 1;
            //int SampleTypeColumn = 10;
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

                if (dataGridViewCSSP[ColumnIndex, RowIndex].Value == null || string.IsNullOrWhiteSpace(dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString()))
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
                            }
                        }
                        else
                        {
                            labSheetA1Sheet.LabSheetA1MeasurementList[RowIndex].Tube10 = null;
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
                            }
                        }
                        else
                        {
                            labSheetA1Sheet.LabSheetA1MeasurementList[RowIndex].Tube1_0 = null;
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
                            }
                        }
                        else
                        {
                            labSheetA1Sheet.LabSheetA1MeasurementList[RowIndex].Tube0_1 = null;
                        }
                    }
                    break;
                default:
                    break;
            }

        }
        private void ValidateSalinityCell(int ColumnIndex, int RowIndex)
        {
            if (dataGridViewCSSP[ColumnIndex, RowIndex].Value != null)
            {
                foreach (char c in dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString())
                {
                    if (!(char.IsDigit(c) || char.IsPunctuation(c)))
                    {
                        dataGridViewCSSP[ColumnIndex, RowIndex].Value = "";
                        return;
                    }
                }

                float theVal = -99;
                if (float.TryParse(dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString(), out theVal))
                {
                    if (!(theVal >= 0.0f && theVal <= 36.0f))
                    {
                        dataGridViewCSSP[ColumnIndex, RowIndex].Value = "";
                        return;
                    }
                }
            }

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
                if (dataGridViewCSSP[SiteColumn, i].Value.ToString() == SiteName)
                {
                    if (dataGridViewCSSP[SampleTypeColumn, i].Value.ToString() == SampleTypeEnum.DailyDuplicate.ToString()
                        || dataGridViewCSSP[SampleTypeColumn, i].Value.ToString() == SampleTypeEnum.IntertechDuplicate.ToString()
                        || dataGridViewCSSP[SampleTypeColumn, i].Value.ToString() == SampleTypeEnum.IntertechRead.ToString())
                    {
                        dataGridViewCSSP[ColumnIndex, i].Value = dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString();
                    }
                }
            }

            float TempFloat = -1;
            if (!string.IsNullOrWhiteSpace(dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString()))
            {
                float.TryParse(dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString(), out TempFloat);
                if (labSheetA1Sheet.LabSheetA1MeasurementList[RowIndex].Salinity != TempFloat)
                {
                    labSheetA1Sheet.LabSheetA1MeasurementList[RowIndex].Salinity = TempFloat;
                }
            }
            else
            {
                labSheetA1Sheet.LabSheetA1MeasurementList[RowIndex].Salinity = null;
            }

        }
        private void ValidateTemperatureCell(int ColumnIndex, int RowIndex)
        {
            if (dataGridViewCSSP[ColumnIndex, RowIndex].Value != null)
            {
                foreach (char c in dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString())
                {
                    if (!(char.IsDigit(c) || char.IsPunctuation(c)))
                    {
                        dataGridViewCSSP[ColumnIndex, RowIndex].Value = "";
                        return;
                    }
                }

                float theVal = -99;
                if (float.TryParse(dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString(), out theVal))
                {
                    if (!(theVal >= -12.0f && theVal <= 36.0f))
                    {
                        dataGridViewCSSP[ColumnIndex, RowIndex].Value = "";
                        return;
                    }
                }
            }

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
                if (!(char.IsNumber(c) || c.ToString() == "." || c.ToString() == "-"))
                {
                    dataGridViewCSSP[ColumnIndex, RowIndex].Style.ForeColor = Color.Red;
                    return;
                }
            }

            float valFloat = -1.0f;
            float.TryParse(val, out valFloat);
            if (valFloat > 36 || valFloat.ToString() != val || valFloat < -12)
            {
                dataGridViewCSSP[ColumnIndex, RowIndex].Style.ForeColor = Color.Red;
                return;
            }

            string SiteName = dataGridViewCSSP[SiteColumn, RowIndex].Value.ToString();
            for (int i = RowIndex + 1, countRow = dataGridViewCSSP.Rows.Count; i < countRow; i++)
            {
                if (dataGridViewCSSP[SiteColumn, i].Value.ToString() == SiteName)
                {
                    if (dataGridViewCSSP[SampleTypeColumn, i].Value.ToString() == SampleTypeEnum.DailyDuplicate.ToString()
                         || dataGridViewCSSP[SampleTypeColumn, i].Value.ToString() == SampleTypeEnum.IntertechDuplicate.ToString()
                         || dataGridViewCSSP[SampleTypeColumn, i].Value.ToString() == SampleTypeEnum.IntertechRead.ToString())
                    {
                        dataGridViewCSSP[ColumnIndex, i].Value = dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString();
                    }
                }
            }

            float TempFloat = -1;
            if (!string.IsNullOrWhiteSpace(dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString()))
            {
                float.TryParse(dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString(), out TempFloat);
                if (labSheetA1Sheet.LabSheetA1MeasurementList[RowIndex].Temperature != TempFloat)
                {
                    labSheetA1Sheet.LabSheetA1MeasurementList[RowIndex].Temperature = TempFloat;
                }
            }
            else
            {
                labSheetA1Sheet.LabSheetA1MeasurementList[RowIndex].Temperature = null;
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
            else
            {
                return;
            }

            dataGridViewCSSP[ColumnIndex, RowIndex].Style.ForeColor = Color.Black;

            string val = dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString();
            foreach (char c in val)
            {
                if (char.IsNumber(c) || c.ToString() == ":")
                {
                }
                else
                {
                    dataGridViewCSSP[ColumnIndex, RowIndex].Style.ForeColor = Color.Red;
                    dataGridViewCSSP[ColumnIndex, RowIndex].Value = "";
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
                    dataGridViewCSSP[ColumnIndex, RowIndex].Value = "";
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
                    dataGridViewCSSP[ColumnIndex, RowIndex].Value = "";
                    return;
                }
                if (!int.TryParse(val.Substring(0, 2), out intVal))
                {
                    dataGridViewCSSP[ColumnIndex, RowIndex].Style.ForeColor = Color.Red;
                    dataGridViewCSSP[ColumnIndex, RowIndex].Value = "";
                    return;
                }
                if (!int.TryParse(val.Substring(3, 2), out intVal))
                {
                    dataGridViewCSSP[ColumnIndex, RowIndex].Style.ForeColor = Color.Red;
                    dataGridViewCSSP[ColumnIndex, RowIndex].Value = "";
                    return;
                }
                if (!(int.Parse(val.Substring(0, 2)) >= 0) || !(int.Parse(val.Substring(0, 2)) <= 23))
                {
                    dataGridViewCSSP[ColumnIndex, RowIndex].Style.ForeColor = Color.Red;
                    dataGridViewCSSP[ColumnIndex, RowIndex].Value = "";
                    return;
                }
                if (!(int.Parse(val.Substring(3, 2)) >= 0) || !(int.Parse(val.Substring(3, 2)) <= 59))
                {
                    dataGridViewCSSP[ColumnIndex, RowIndex].Style.ForeColor = Color.Red;
                    dataGridViewCSSP[ColumnIndex, RowIndex].Value = "";
                    return;
                }
            }
            if (val.Length < 4)
            {
                dataGridViewCSSP[ColumnIndex, RowIndex].Style.ForeColor = Color.Red;
                dataGridViewCSSP[ColumnIndex, RowIndex].Value = "";
                return;
            }

            string SiteName = dataGridViewCSSP[SiteColumn, RowIndex].Value.ToString();
            for (int i = RowIndex + 1, countRow = dataGridViewCSSP.Rows.Count; i < countRow; i++)
            {
                if (dataGridViewCSSP[SiteColumn, i].Value.ToString() == SiteName)
                {
                    if (dataGridViewCSSP[SampleTypeColumn, i].Value.ToString() == SampleTypeEnum.DailyDuplicate.ToString()
                       || dataGridViewCSSP[SampleTypeColumn, i].Value.ToString() == SampleTypeEnum.IntertechDuplicate.ToString()
                       || dataGridViewCSSP[SampleTypeColumn, i].Value.ToString() == SampleTypeEnum.IntertechRead.ToString())
                    {
                        dataGridViewCSSP[ColumnIndex, i].Value = dataGridViewCSSP[ColumnIndex, RowIndex].Value.ToString();
                    }

                    int Hour = 0;
                    int Minute = 0;

                    if (dataGridViewCSSP[ColumnIndex, i].Value.ToString().Length > 2)
                    {
                        int.TryParse(dataGridViewCSSP[ColumnIndex, i].Value.ToString().Substring(0, 2), out Hour);
                    }
                    if (dataGridViewCSSP[ColumnIndex, i].Value.ToString().Length > 4)
                    {
                        int.TryParse(dataGridViewCSSP[ColumnIndex, i].Value.ToString().Substring(3, 2), out Minute);
                    }

                    labSheetA1Sheet.LabSheetA1MeasurementList[RowIndex].Time = new DateTime(int.Parse(labSheetA1Sheet.RunYear), int.Parse(labSheetA1Sheet.RunMonth), int.Parse(labSheetA1Sheet.RunDay), Hour, Minute, 0);
                }
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

        }


    }
}
