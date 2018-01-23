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
        #region Constructors
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
        #endregion Constructors

    }
}
