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

        //public class GridCellText
        //{
        //    public GridCellText()
        //    {

        //    }

        //    public string Site { get; set; }
        //    public string Time { get; set; }
        //    public string MPN { get; set; }
        //    public string Tube10 { get; set; }
        //    public string Tube1_0 { get; set; }
        //    public string Tube0_1 { get; set; }
        //    public string Sal { get; set; }
        //    public string Temp { get; set; }
        //    public string ProcessBy { get; set; }
        //    public string SampleType { get; set; }
        //    public string ID { get; set; }
        //    public string Comment { get; set; }
        //}

        class RunNumberAndText
        {
            public string RunNumberText { get; set; }
            public int RunNumber { get; set; }
        }
    }
}
