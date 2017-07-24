using System;
using System.Text;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Globalization;
using Microsoft.QualityTools.Testing.Fakes;
using CSSPWQInputTool.Fakes;
using System.Windows.Forms;
using System.Diagnostics;
using System.Drawing;
using System.Threading;
using CSSPModelsDLL.Models;
using CSSPEnumsDLL.Enums;

namespace CSSPWQInputTool.Tests.app
{
    /// <summary>
    /// Summary description for FormTest
    /// </summary>
    [TestClass]
    public class FormTest
    {
        #region Variables
        //private string userName = "Charles";
        //private string userName = "leblancc";
        private string r = "OgW2S3EHhQ(6!Z$odV7eAGnim/#YIClk9vF&1@5xDUa)wPLu*BN.t,c8%JRMbK^yqzXpfTj4sr0:d";
        private bool InternetConnection = false;
        private string FormTitle = "";
        private List<CSSPWQInputParam> csspWQInputParamList = new List<CSSPWQInputParam>();
        private CSSPWQInputApp csspWQInputApp = new CSSPWQInputApp();
        private Color ButBackColor = Color.Black;
        private CSSPWQInputTypeEnum csspWQInputTypeCurrent = CSSPWQInputTypeEnum.Subsector;
        private CSSPWQInputSheetTypeEnum csspWQInputSheetType = CSSPWQInputSheetTypeEnum.A1;
        private string CurrentPath = "";
        private string NameCurrent = "";
        private int TVItemIDCurrent = 0;
        private string YearMonthDayCurrent = "";
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
        private int VersionOfSamplingPlanFile = 1;
        private int VersionOfResultFile = 1;
        private Panel CurrentPanel = null;
        private string Initials = "";
        private bool IsOnDuplicate = false;
        private bool AppIsWide = false;
        private List<string> AllowableTideString = new List<string>()
        {
            "/", "--", "HR", "HT", "HF", "MR", "MT", "MF", "LR", "LT", "LF",
        };
        private StringBuilder sbPrevCommands = new StringBuilder();
        private StringBuilder sbNewCommands = new StringBuilder();
        #endregion Variables

        #region Properties
        public List<CultureInfo> cultureListGood { get; set; }
        public CSSPWQInputToolForm csspWQInputToolForm { get; set; }
        public PrivateObject privateObject { get; set; }
        public ShimCSSPWQInputToolForm shimCSSPWQInputToolForm { get; set; }
        public Panel panelPassword = new Panel();
        public Panel panelPasswordCenter = new Panel();
        public Panel panelAccessCode = new Panel();
        public Panel panelApp = new Panel();
        public Panel panelAppInputFiles = new Panel();
        public Panel panelAppInputFilesTop = new Panel();
        public Panel panelAppInput = new Panel();
        public Panel panelAddInputMiddle = new Panel();
        public Panel panelAppInputTop = new Panel();
        public Panel panelControl = new Panel();
        public Panel panelTC = new Panel();
        public Panel panelAppInputTopIncubation = new Panel();
        public Panel panelAppInputTopTideCrew = new Panel();
        public Panel panelAppInputBottom = new Panel();
        public Panel panelAddInputBottomRight = new Panel();
        public Panel panelLineForSignature = new Panel();
        public Panel panelAddInputBottomLeft = new Panel();
        public Panel panelAddInputBottomLeftDuplicate = new Panel();
        public Panel panelAppTop = new Panel();
        public Panel panelChangeDateOfCurrentDoc = new Panel();
        public Panel panelStatusBar = new Panel();
        public Panel panelButtonBar = new Panel();

        public TextBox textBoxInitials = new TextBox();
        public TextBox textBoxAccessCode = new TextBox();
        public TextBox textBoxLot44_5 = new TextBox();
        public TextBox textBoxLot35 = new TextBox();
        public TextBox textBoxControlBlank35 = new TextBox();
        public TextBox textBoxControlBath1Blank44_5 = new TextBox();
        public TextBox textBoxControlBath2Blank44_5 = new TextBox();
        public TextBox textBoxControlBath3Blank44_5 = new TextBox();
        public TextBox textBoxControlBath1Negative44_5 = new TextBox();
        public TextBox textBoxControlBath2Negative44_5 = new TextBox();
        public TextBox textBoxControlBath3Negative44_5 = new TextBox();
        public TextBox textBoxControlNegative35 = new TextBox();
        public TextBox textBoxControlBath1NonTarget44_5 = new TextBox();
        public TextBox textBoxControlBath2NonTarget44_5 = new TextBox();
        public TextBox textBoxControlBath3NonTarget44_5 = new TextBox();
        public TextBox textBoxControlNonTarget35 = new TextBox();
        public TextBox textBoxControlLot = new TextBox();
        public TextBox textBoxControlBath1Positive44_5 = new TextBox();
        public TextBox textBoxControlBath2Positive44_5 = new TextBox();
        public TextBox textBoxControlBath3Positive44_5 = new TextBox();
        public TextBox textBoxControlPositive35 = new TextBox();
        public TextBox textBoxTCLab1 = new TextBox();
        public TextBox textBoxTCField1 = new TextBox();
        public TextBox textBoxTCLab2 = new TextBox();
        public TextBox textBoxTCField2 = new TextBox();
        public TextBox textBoxWaterBathNumber = new TextBox();
        public TextBox textBoxIncubationEndTime = new TextBox();
        public TextBox textBoxIncubationBath1StartTime = new TextBox();
        public TextBox textBoxIncubationBath2StartTime = new TextBox();
        public TextBox textBoxIncubationBath3StartTime = new TextBox();
        public TextBox textBoxTides = new TextBox();
        public TextBox textBoxSampleCrewInitials = new TextBox();
        public TextBox textBoxResultsReadBy = new TextBox();
        public TextBox textBoxResultsRecordedBy = new TextBox();
        public TextBox textBoxSalinitiesReadBy = new TextBox();
        public TextBox textBoxSampleBottleLotNumber = new TextBox();
        public TextBox textBoxDailyDuplicatePrecisionCriteria = new TextBox();

        public Button butBrowseSamplingPlanFile = new Button();
        public Button butOpen = new Button();
        public Button butGetTides = new Button();
        public Button butViewFCForm = new Button();
        public Button butCreateFile = new Button();
        public Button butChangeDateCancel = new Button();
        public Button butChangeDate = new Button();
        public Button butSendToServer = new Button();
        public Button butSubsector = new Button();
        public Button butMunicipality = new Button();
        public Button butArchive = new Button();
        public Button butEC = new Button();
        public Button butLTB = new Button();
        public Button butA1 = new Button();
        public Button butLogoff = new Button();

        public ListBox listBoxFiles = new ListBox();

        public ComboBox comboBoxFileSubsector = new ComboBox();
        public ComboBox comboBoxSubsectorNames = new ComboBox();

        public RichTextBox richTextBoxFile = new RichTextBox();
        public RichTextBox richTextBoxRunWeatherComment = new RichTextBox();
        public RichTextBox richTextBoxRunComment = new RichTextBox();

        public DataGridView dataGridViewCSSP = new DataGridView();

        public CheckBox checkBox2Coolers = new CheckBox();

        public Label lblIncubationTimeCalculated = new Label();
        public Label lblSampleCrewInitials = new Label();
        public Label lblSamplingPlanFileName = new Label();

        public WebBrowser webBrowserCSSP = new WebBrowser();

        public DateTimePicker dateTimePickerResultsReadDate = new DateTimePicker();
        public DateTimePicker dateTimePickerResultsRecordedDate = new DateTimePicker();
        public DateTimePicker dateTimePickerSalinitiesReadDate = new DateTimePicker();
        public DateTimePicker dateTimePickerDuplicateDataEntryDate = new DateTimePicker();
        public DateTimePicker dateTimePickerRun = new DateTimePicker();
        public DateTimePicker dateTimePickerChangeDate = new DateTimePicker();

        public System.Windows.Forms.Timer timerSave = new System.Windows.Forms.Timer();
        public System.Windows.Forms.Timer timerGetTides = new System.Windows.Forms.Timer();

        public OpenFileDialog openFileDialogCSSP = new OpenFileDialog();

        public Process processCSSP = new Process();

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
      
        #endregion Properties

        #region Constructors
        public FormTest()
        {
            cultureListGood = new List<CultureInfo>() { new CultureInfo("en-CA"), new CultureInfo("fr-CA") };
            csspWQInputToolForm = new CSSPWQInputToolForm();
            privateObject = new PrivateObject(csspWQInputToolForm);
            SetupTest(cultureListGood[0]);

            privateObject.Invoke("butBrowseSamplingPlanFile_Click", butBrowseSamplingPlanFile, new EventArgs());

            //bool RetBool = (bool)privateObject.Invoke("ReadSamplingPlan");
            //if (RetBool)
            //{
            //    panelAccessCode.Visible = true;
            //}
            //textBoxInitials.Focus();
            //textBoxInitials.Text = "AA";
            //Initials = textBoxInitials.Text;
            //textBoxAccessCode.Focus();
            //textBoxAccessCode.Text = "abcdef";

            LoadVariable();
        }
        #endregion Constructors

        #region Initialize and Cleanup
        //
        // You can use the following additional attributes as you write your tests:
        //
        // Use ClassInitialize to run code before running the first test in the class
        // [ClassInitialize()]
        // public static void MyClassInitialize(TestContext testContext) { }
        //
        // Use ClassCleanup to run code after all tests in a class have run
        // [ClassCleanup()]
        // public static void MyClassCleanup() { }
        //
        // Use TestInitialize to run code before running each test 
        // [TestInitialize()]
        // public void MyTestInitialize() { }
        //
        // Use TestCleanup to run code after each test has run
        // [TestCleanup()]
        // public void MyTestCleanup() { }
        //
        #endregion Initialize and Cleanup

        #region Testing Constructors
        [TestMethod]
        public void CSSPWQInputToolForm_Constructor()
        {
            Assert.IsNotNull(csspWQInputToolForm);
            Assert.AreEqual(@"OgW2S3EHhQ(6!Z$odV7eAGnim/#YIClk9vF&1@5xDUa)wPLu*BN.t,c8%JRMbK^yqzXpfTj4sr0:d", r);
            Assert.AreEqual(true, InternetConnection);
            Assert.AreEqual("CSSP Water Quality Input Tool", FormTitle);
            Assert.AreEqual(1, csspWQInputParamList.Count);
            Assert.IsNotNull(csspWQInputApp);
            Assert.AreEqual(butMunicipality.BackColor, ButBackColor);
            Assert.AreEqual(CSSPWQInputTypeEnum.Subsector, csspWQInputTypeCurrent);
            Assert.AreEqual(CSSPWQInputSheetTypeEnum.A1, csspWQInputSheetType);
            Assert.AreEqual(@"C:\CSSPLabSheets\SamplingPlan_Testing\2015\", CurrentPath);
            Assert.AreEqual(@"NB-01-020-002 (Charlo)", NameCurrent);
            Assert.AreEqual(560, TVItemIDCurrent);
            Assert.AreEqual(DateTime.Now.ToString("yyyy_MM_dd"), YearMonthDayCurrent);
            Assert.IsNotNull(CSSPWQInputParamCurrent);
            Assert.IsNotNull(dataGridViewCellStyleDefault);
            Assert.IsNotNull(dataGridViewCellStyleEdit);
            Assert.IsNotNull(dataGridViewCellStyleEditRowCell);
            Assert.IsNotNull(dataGridViewCellStyleEditError);
            Assert.IsNotNull(csspMPNTableList);
            Assert.AreEqual(false, InLoadingFile);
            Assert.AreEqual(@"C:\CSSPLabSheets\SamplingPlan_Testing.txt", SamplingPlanName);
            Assert.AreEqual(false, NoUpdate);
            Assert.AreEqual(0, TideToTryIndex);
            Assert.AreEqual(true, panelAppInputIsVisible);
            Assert.AreEqual(lblSampleCrewInitials.BackColor, ControlBackColor);
            Assert.AreEqual(textBoxSampleCrewInitials.BackColor, TextBoxBackColor);
            Assert.AreEqual(dataGridViewCSSP.BackgroundColor, DataGridViewCSSPBackgroundColor);
            Assert.AreEqual(1, VersionOfSamplingPlanFile);
            Assert.AreEqual(1, VersionOfResultFile);
            Assert.AreEqual(panelAppInput.Name, CurrentPanel.Name);
            Assert.AreEqual("AA", Initials);
            Assert.AreEqual(false, IsOnDuplicate);
            Assert.AreEqual(false, AppIsWide);
            List<string> TestAllowableTideStringList = new List<string>() { "/", "--", "HR", "HT", "HF", "MR", "MT", "MF", "LR", "LT", "LF", };
            for (int i = 0, count = TestAllowableTideStringList.Count; i < count; i++)
            {
                Assert.AreEqual(TestAllowableTideStringList[i], AllowableTideString[i]);
            }
            Assert.IsNotNull(sbPrevCommands);
            Assert.AreEqual("", sbPrevCommands.ToString());
            Assert.IsNotNull(sbNewCommands);
            Assert.AreEqual("", sbNewCommands.ToString());

            Assert.AreEqual(@"C:\CSSPLabSheets\SamplingPlan_Testing.txt", lblSamplingPlanFileName.Text);
            Assert.AreEqual(DockStyle.Fill, panelPassword.Dock);
            Assert.AreEqual(DockStyle.Fill, panelApp.Dock);
            Assert.AreEqual(false, panelButtonBar.Visible);
            Assert.AreEqual(new Point(panelPassword.Width / 2 - panelPasswordCenter.Size.Width / 2, panelPassword.Height / 2 - panelPasswordCenter.Size.Height / 2), panelPasswordCenter.Location);
            Assert.AreEqual(AnchorStyles.None, panelPasswordCenter.Anchor);
            Assert.AreEqual(DockStyle.Fill, panelAppInput.Dock);
            Assert.AreEqual(DockStyle.Fill, panelAppInputFiles.Dock);
            Assert.AreEqual(Color.LightGreen, butSubsector.BackColor);
            Assert.AreEqual(Color.LightGreen, butA1.BackColor);
            Assert.AreEqual(true, webBrowserCSSP.ScriptErrorsSuppressed);

            Assert.IsNotNull(csspWQInputToolForm.csspFCFormWriter);
            Assert.IsNotNull(csspWQInputToolForm.csspLabSheetParser);
        }
        #endregion Testing Construtors

        #region Testing Events
        #endregion Testing Events

        #region Testing Methods
        [TestMethod]
        public void CSSPWQInputToolForm_CreateCode_Good()
        {
            string Value = "0.093";
            string retStr = (string)privateObject.Invoke("CreateCode", Value);
            string retStr2 = (string)privateObject.Invoke("GetCodeString", retStr);
            Assert.AreEqual(Value, retStr2);
        }
        [TestMethod]
        public void CSSPWQInputToolForm_AddLog_Good()
        {
            string Element = "elem";
            string NewValue = "newValue";
            privateObject.Invoke("AddLog", Element, NewValue);
            Assert.AreEqual((DateTime.Now + "\t" + Initials + "\t" + Element + "\t" + NewValue) + "\r\n", sbNewCommands.ToString());
        }
        [TestMethod]
        public void CSSPWQInputToolForm_CalculateDuplicate_Good()
        {
            privateObject.Invoke("CalculateDuplicate");
            Assert.AreEqual(false, true);
        }
        #endregion Testing Methods

        #region Testing others
        [TestMethod]
        public void Create_Random_String_Test()
        {
            List<int> intList = new List<int>();
            Random rd = new Random();
            //                    1         2         3         4         5         6         7         8         9         1
            //          012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789
            string r = "OgW2S3EHhQ(6!Z$odV7eAGnim/#YIClk9vF&1@5xDUa)wPLu*BN.t,c8%JRMbK^yqzXpfTj4sr0";
            string str = "12345";
            foreach (char c in str)
            {
                int pos = r.IndexOf(c);
                int first = rd.Next(pos + 1, pos + 9);
                int second = rd.Next(2, 9);
                int tot = (first * second) + pos;
                intList.Add(tot);
                intList.Add(first);
            }

            string pw = "";
            foreach (int i in intList)
            {
                pw = pw + i + ",";
            }
        }
        #endregion Testing others

        #region Functions private
        private void SetupTest(CultureInfo culture)
        {
            csspWQInputToolForm = new CSSPWQInputToolForm();

            Assert.AreEqual(csspWQInputToolForm.currentCulture, Thread.CurrentThread.CurrentCulture);
            Assert.AreEqual(csspWQInputToolForm.currentUICulture, Thread.CurrentThread.CurrentCulture);

            // Panel
            panelPassword = ((Panel)privateObject.GetField("panelPassword"));
            panelPasswordCenter = ((Panel)privateObject.GetField("panelPasswordCenter"));
            panelAccessCode = ((Panel)privateObject.GetField("panelAccessCode"));
            panelApp = ((Panel)privateObject.GetField("panelApp"));
            panelAppInputFiles = ((Panel)privateObject.GetField("panelAppInputFiles"));
            panelAppInputFilesTop = ((Panel)privateObject.GetField("panelAppInputFilesTop"));
            panelAppInput = ((Panel)privateObject.GetField("panelAppInput"));
            panelAddInputMiddle = ((Panel)privateObject.GetField("panelAddInputMiddle"));
            panelAppInputTop = ((Panel)privateObject.GetField("panelAppInputTop"));
            panelControl = ((Panel)privateObject.GetField("panelControl"));
            panelTC = ((Panel)privateObject.GetField("panelTC"));
            panelAppInputTopIncubation = ((Panel)privateObject.GetField("panelAppInputTopIncubation"));
            panelAppInputTopTideCrew = ((Panel)privateObject.GetField("panelAppInputTopTideCrew"));
            panelAppInputBottom = ((Panel)privateObject.GetField("panelAppInputBottom"));
            panelAddInputBottomRight = ((Panel)privateObject.GetField("panelAddInputBottomRight"));
            panelLineForSignature = ((Panel)privateObject.GetField("panelLineForSignature"));
            panelAddInputBottomLeft = ((Panel)privateObject.GetField("panelAddInputBottomLeft"));
            panelAddInputBottomLeftDuplicate = ((Panel)privateObject.GetField("panelAddInputBottomLeftDuplicate"));
            panelAppTop = ((Panel)privateObject.GetField("panelAppTop"));
            panelChangeDateOfCurrentDoc = ((Panel)privateObject.GetField("panelChangeDateOfCurrentDoc"));
            panelStatusBar = ((Panel)privateObject.GetField("panelStatusBar"));
            panelButtonBar = ((Panel)privateObject.GetField("panelButtonBar"));

            // TextBox
            textBoxInitials = ((TextBox)privateObject.GetField("textBoxInitials"));
            textBoxAccessCode = ((TextBox)privateObject.GetField("textBoxAccessCode"));
            textBoxLot44_5 = ((TextBox)privateObject.GetField("textBoxLot44_5"));
            textBoxLot35 = ((TextBox)privateObject.GetField("textBoxLot35"));
            textBoxControlBlank35 = ((TextBox)privateObject.GetField("textBoxControlBlank35"));
            textBoxControlBath1Negative44_5 = ((TextBox)privateObject.GetField("textBoxControlBath1Negative44_5"));
            textBoxControlBath2Negative44_5 = ((TextBox)privateObject.GetField("textBoxControlBath2Negative44_5"));
            textBoxControlBath3Negative44_5 = ((TextBox)privateObject.GetField("textBoxControlBath3Negative44_5"));
            textBoxControlNegative35 = ((TextBox)privateObject.GetField("textBoxControlNegative35"));
            textBoxControlBath1NonTarget44_5 = ((TextBox)privateObject.GetField("textBoxControlBath1NonTarget44_5"));
            textBoxControlBath2NonTarget44_5 = ((TextBox)privateObject.GetField("textBoxControlBath2NonTarget44_5"));
            textBoxControlBath3NonTarget44_5 = ((TextBox)privateObject.GetField("textBoxControlBath3NonTarget44_5"));
            textBoxControlNonTarget35 = ((TextBox)privateObject.GetField("textBoxControlNonTarget35"));
            textBoxControlLot = ((TextBox)privateObject.GetField("textBoxControlLot"));
            textBoxControlBath1Positive44_5 = ((TextBox)privateObject.GetField("textBoxControlBath1Positive44_5"));
            textBoxControlBath2Positive44_5 = ((TextBox)privateObject.GetField("textBoxControlBath2Positive44_5"));
            textBoxControlBath3Positive44_5 = ((TextBox)privateObject.GetField("textBoxControlBath3Positive44_5"));
            textBoxControlPositive35 = ((TextBox)privateObject.GetField("textBoxControlPositive35"));
            textBoxTCLab1 = ((TextBox)privateObject.GetField("textBoxTCLab1"));
            textBoxTCField1 = ((TextBox)privateObject.GetField("textBoxTCField1"));
            textBoxTCLab2 = ((TextBox)privateObject.GetField("textBoxTCLab2"));
            textBoxTCField2 = ((TextBox)privateObject.GetField("textBoxTCField2"));
            textBoxWaterBathNumber = ((TextBox)privateObject.GetField("textBoxWaterBathNumber"));
            textBoxIncubationEndTime = ((TextBox)privateObject.GetField("textBoxIncubationEndTime"));
            textBoxIncubationBath1StartTime = ((TextBox)privateObject.GetField("textBoxIncubationBath1StartTime"));
            textBoxIncubationBath2StartTime = ((TextBox)privateObject.GetField("textBoxIncubationBath2StartTime"));
            textBoxIncubationBath3StartTime = ((TextBox)privateObject.GetField("textBoxIncubationBath3StartTime"));
            textBoxTides = ((TextBox)privateObject.GetField("textBoxTides"));
            textBoxSampleCrewInitials = ((TextBox)privateObject.GetField("textBoxSampleCrewInitials"));
            textBoxResultsReadBy = ((TextBox)privateObject.GetField("textBoxResultsReadBy"));
            textBoxResultsRecordedBy = ((TextBox)privateObject.GetField("textBoxResultsRecordedBy"));
            textBoxSalinitiesReadBy = ((TextBox)privateObject.GetField("textBoxSalinitiesReadBy"));
            textBoxSampleBottleLotNumber = ((TextBox)privateObject.GetField("textBoxSampleBottleLotNumber"));
            textBoxDailyDuplicatePrecisionCriteria = ((TextBox)privateObject.GetField("textBoxDailyDuplicatePrecisionCriteria"));

            // Button
            butBrowseSamplingPlanFile = ((Button)privateObject.GetField("butBrowseSamplingPlanFile"));
            butOpen = ((Button)privateObject.GetField("butOpen"));
            butGetTides = ((Button)privateObject.GetField("butGetTides"));
            butViewFCForm = ((Button)privateObject.GetField("butViewFCForm"));
            butCreateFile = ((Button)privateObject.GetField("butCreateFile"));
            butChangeDateCancel = ((Button)privateObject.GetField("butChangeDateCancel"));
            butChangeDate = ((Button)privateObject.GetField("butChangeDate"));
            butSendToServer = ((Button)privateObject.GetField("butSendToServer"));
            butSubsector = ((Button)privateObject.GetField("butSubsector"));
            butMunicipality = ((Button)privateObject.GetField("butMunicipality"));
            butArchive = ((Button)privateObject.GetField("butArchive"));
            butEC = ((Button)privateObject.GetField("butEC"));
            butLTB = ((Button)privateObject.GetField("butLTB"));
            butA1 = ((Button)privateObject.GetField("butA1"));
            butLogoff = ((Button)privateObject.GetField("butLogoff"));

            // ListBox
            listBoxFiles = ((ListBox)privateObject.GetField("listBoxFiles"));

            // ComboBox
            comboBoxFileSubsector = ((ComboBox)privateObject.GetField("comboBoxFileSubsector"));
            comboBoxSubsectorNames = ((ComboBox)privateObject.GetField("comboBoxSubsectorNames"));

            // RichTextBox
            richTextBoxFile = ((RichTextBox)privateObject.GetField("richTextBoxFile"));
            richTextBoxRunWeatherComment = ((RichTextBox)privateObject.GetField("richTextBoxRunWeatherComment"));
            richTextBoxRunComment = ((RichTextBox)privateObject.GetField("richTextBoxRunComment"));

            // DataGridView
            dataGridViewCSSP = ((DataGridView)privateObject.GetField("dataGridViewCSSP"));

            // Checkbox
            checkBox2Coolers = ((CheckBox)privateObject.GetField("checkBox2Coolers"));

            // Label
            lblIncubationTimeCalculated = ((Label)privateObject.GetField("lblIncubationTimeCalculated"));
            lblSampleCrewInitials = ((Label)privateObject.GetField("lblSampleCrewInitials"));
            lblSamplingPlanFileName = ((Label)privateObject.GetField("lblSamplingPlanFileName"));

            // WebBrowser
            webBrowserCSSP = ((WebBrowser)privateObject.GetField("webBrowserCSSP"));

            // DateTimePicker
            dateTimePickerResultsReadDate = ((DateTimePicker)privateObject.GetField("dateTimePickerResultsReadDate"));
            dateTimePickerResultsRecordedDate = ((DateTimePicker)privateObject.GetField("dateTimePickerResultsRecordedDate"));
            dateTimePickerSalinitiesReadDate = ((DateTimePicker)privateObject.GetField("dateTimePickerSalinitiesReadDate"));
            dateTimePickerRun = ((DateTimePicker)privateObject.GetField("dateTimePickerRun"));
            dateTimePickerChangeDate = ((DateTimePicker)privateObject.GetField("dateTimePickerChangeDate"));

            // Timer
            timerSave = ((System.Windows.Forms.Timer)privateObject.GetField("timerSave"));
            timerGetTides = ((System.Windows.Forms.Timer)privateObject.GetField("timerGetTides"));

            // OpenFileDialog
            openFileDialogCSSP = ((OpenFileDialog)privateObject.GetField("openFileDialogCSSP"));

            // Process
            processCSSP = ((Process)privateObject.GetField("processCSSP"));



        }
        private void LoadVariable()
        {
            r = (string)privateObject.GetField("r");
            InternetConnection = (bool)privateObject.GetField("InternetConnection");
            FormTitle = (string)privateObject.GetField("FormTitle");
            csspWQInputParamList = ((List<CSSPWQInputParam>)privateObject.GetField("csspWQInputParamList"));
            csspWQInputApp = (CSSPWQInputApp)privateObject.GetField("csspWQInputApp");
            ButBackColor = (Color)privateObject.GetField("ButBackColor");
            csspWQInputTypeCurrent = (CSSPWQInputTypeEnum)privateObject.GetField("csspWQInputTypeCurrent");
            csspWQInputSheetType = (CSSPWQInputSheetTypeEnum)privateObject.GetField("csspWQInputSheetType");
            CurrentPath = (string)privateObject.GetField("CurrentPath");
            NameCurrent = (string)privateObject.GetField("NameCurrent");
            TVItemIDCurrent = (int)privateObject.GetField("TVItemIDCurrent");
            YearMonthDayCurrent = (string)privateObject.GetField("YearMonthDayCurrent");
            CSSPWQInputParamCurrent = (CSSPWQInputParam)privateObject.GetField("CSSPWQInputParamCurrent");
            dataGridViewCellStyleDefault = (DataGridViewCellStyle)privateObject.GetField("dataGridViewCellStyleDefault");
            dataGridViewCellStyleEdit = (DataGridViewCellStyle)privateObject.GetField("dataGridViewCellStyleEdit");
            dataGridViewCellStyleEditRowCell = (DataGridViewCellStyle)privateObject.GetField("dataGridViewCellStyleEditRowCell");
            dataGridViewCellStyleEditError = (DataGridViewCellStyle)privateObject.GetField("dataGridViewCellStyleEditError");
            csspMPNTableList = (List<CSSPMPNTable>)privateObject.GetField("csspMPNTableList");
            InLoadingFile = (bool)privateObject.GetField("InLoadingFile");
            SamplingPlanName = (string)privateObject.GetField("SamplingPlanName");
            NoUpdate = (bool)privateObject.GetField("NoUpdate");
            TideToTryIndex = (int)privateObject.GetField("TideToTryIndex");
            panelAppInputIsVisible = (bool)privateObject.GetField("panelAppInputIsVisible");
            ControlBackColor = (Color)privateObject.GetField("ControlBackColor");
            TextBoxBackColor = (Color)privateObject.GetField("TextBoxBackColor");
            DataGridViewCSSPBackgroundColor = (Color)privateObject.GetField("DataGridViewCSSPBackgroundColor");
            VersionOfSamplingPlanFile = (int)privateObject.GetField("VersionOfSamplingPlanFile");
            VersionOfResultFile = (int)privateObject.GetField("VersionOfResultFile");
            CurrentPanel = (Panel)privateObject.GetField("CurrentPanel");
            Initials = (string)privateObject.GetField("Initials");
            AppIsWide = (bool)privateObject.GetField("AppIsWide");
            AllowableTideString = (List<string>)privateObject.GetField("AllowableTideString");

            textBoxInitials.Text = "AA";
            textBoxAccessCode.Focus();
        }
        private void SetupShims()
        {
            shimCSSPWQInputToolForm = new ShimCSSPWQInputToolForm();
        }
        #endregion Functions private

    }
}
