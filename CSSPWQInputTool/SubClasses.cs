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
        class AcceptedOrRejected
        {
            public string AcceptedOrRejectedBy { get; set; }
            public DateTime AcceptedOrRejectedDate { get; set; }
            public string RejectReason { get; set; }
        }

        class RunNumberAndText
        {
            public string RunNumberText { get; set; }
            public int RunNumber { get; set; }
        }
    }
}
