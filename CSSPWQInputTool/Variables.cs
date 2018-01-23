﻿using System;
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
        private string OpenedFileName = "";

        // ------------------------ Start Lab Sheet Variables ----------------------
        private string StartTide = "";
        private string StartSampleCrewInitials = "";
        private string Start2Coolers = "";
        private string StartTCField1 = "";
        private string StartTCLab1 = "";
        private string StartTCField2 = "";
        private string StartTCLab2 = "";
        private string StartIncubationStartSameDay = "";
        private string StartNumberOfBaths = "";
        private string StartIncubationBath1StartTime = "";
        private string StartIncubationBath1EndTime = "";
        private string StartWaterBath1Number = "";
        private string StartIncubationBath2StartTime = "";
        private string StartIncubationBath2EndTime = "";
        private string StartWaterBath2Number = "";
        private string StartIncubationBath3StartTime = "";
        private string StartIncubationBath3EndTime = "";
        private string StartWaterBath3Number = "";
        private string StartControlLot = "";
        private string StartControlPositive35 = "";
        private string StartControlNonTarget35 = "";
        private string StartControlNegative35 = "";
        private string StartControlBlank35 = "";
        private string StartControlBath1Positive44_5 = "";
        private string StartControlBath1NonTarget44_5 = "";
        private string StartControlBath1Negative44_5 = "";
        private string StartControlBath1Blank44_5 = "";
        private string StartControlBath2Positive44_5 = "";
        private string StartControlBath2NonTarget44_5 = "";
        private string StartControlBath2Negative44_5 = "";
        private string StartControlBath2Blank44_5 = "";
        private string StartControlBath3Positive44_5 = "";
        private string StartControlBath3NonTarget44_5 = "";
        private string StartControlBath3Negative44_5 = "";
        private string StartControlBath3Blank44_5 = "";
        private string StartLot35 = "";
        private string StartLot44_5 = "";
        private string StartRunWeatherComment = "";
        private string StartRunComment = "";
        private string StartSampleBottleLotNumber = "";
        private string StartSalinityReadBy = "";
        private string StartSalinityReadDateYear = "";
        private string StartSalinityReadDateMonth = "";
        private string StartSalinityReadDateDay = "";
        private string StartResultsReadBy = "";
        private string StartResultsReadDateYear = "";
        private string StartResultsReadDateMonth = "";
        private string StartResultsReadDateDay = "";
        private string StartResultsRecordedBy = "";
        private string StartResultsRecordedDateYear = "";
        private string StartResultsRecordedDateMonth = "";
        private string StartResultsRecordedDateDay = "";
        private List<GridCellText> StartGridCellText = new List<GridCellText>();
    }
}
