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
        public LabSheetA1Sheet labSheetA1Sheet { get; set; }
        public CSSPFCFormWriter csspFCFormWriter { get; set; }
        public CSSPLabSheetParser csspLabSheetParser { get; set; }
        public CultureInfo currentCulture { get; set; }
        public CultureInfo currentUICulture { get; set; }

    }
}
