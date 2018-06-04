/*
* Excel handling class
* Excel reference needs to be added to the project:
* Project -> Add Reference -> .NET -> Microsoft.Office.Interop.Excel -> Choose 2.0 Version
*
* Iteration through Excel Worksheet is possible using GetNextCell and ToNewLine method
* At the end of usage file must be closed with Close function
*/
namespace HandleExcel {
    using Excel = Microsoft.Office.Interop.Excel;
    using System;

    public enum Align {
        Left,
        Right,
        Center
    }

    public enum Border {
        Top,
        Bottom,
        Right,
        Left,
        All
    }

    public enum HeaderPosition {
        Left,
        Right,
        Center,
        Default
    }

    public enum FooterPosition {
        Left,
        Right,
        Center,
        Default
    }

    public enum Orientations {
        Portrait,
        Landscape,
        Default
    }

    public sealed class HandleExcel : IDisposable {
        private Excel.Application XlApplication;
        private Excel.Workbook XlWorkbook;
        private Excel.Worksheet XlWorksheet;
        private Excel.Range XlRange;
        private bool closeFlag;
        private bool is_console_app;
        private int X;
        private int Y;

        public int ActiveRow { get { return X; } }
        public int ActiveColumn { get { return Y; } }
        public string FileName { get { return XlWorkbook.Name; } }
        public int GetRowCount { get { return XlWorksheet.UsedRange.Rows.Count; } }
        public int GetColumnCount { get { return XlWorksheet.UsedRange.Columns.Count; } }

        public HandleExcel(string Path = "", string Password = "") {
            // Excel Dll only works with en-US locale. If the locale is different, this line enforces en-US.
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
            is_console_app = Console.OpenStandardInput(1) != System.IO.Stream.Null;
            XlApplication = new Excel.Application();
            XlApplication.Visible = false;
            XlApplication.DisplayAlerts = false;
            if (!string.IsNullOrEmpty(Password)) {
                XlWorkbook = XlApplication.Workbooks.Open(Path, Password: Password);
            } else if (!string.IsNullOrEmpty(Path)) {
                XlWorkbook = XlApplication.Workbooks.Open(Path);
            } else {
                XlWorkbook = XlApplication.Workbooks.Add();
            }
            XlWorksheet = XlWorkbook.Worksheets[1] as Excel.Worksheet;
            XlRange = XlWorksheet.get_Range("A1");
            X = 1;
            Y = 1;
            closeFlag = true;
        }

        public void AddWorksheet(string Name = "") {
            Excel.Worksheet temp = XlWorksheet;
            XlWorksheet = XlWorkbook.Worksheets.Add() as Excel.Worksheet;
            if (!string.IsNullOrEmpty(Name)) XlWorksheet.Name = Name;
            XlWorksheet = temp;
        }

        public void DeleteWorksheet(string Name = "") {
            if (!string.IsNullOrEmpty(Name)) XlWorksheet = XlWorkbook.Worksheets[Name] as Excel.Worksheet;
            XlWorksheet.Delete();
        }

        public string[] GetWorksheetNames() {
            var t = new System.Collections.Generic.List<string>();
            foreach (Excel.Worksheet s in XlWorkbook.Worksheets) {
                t.Add(s.Name);
            }
            return t.ToArray();
        }

        public void Close() {
            if (closeFlag) {
                closeFlag = false;
                XlRange = null;
                XlWorksheet = null;
                XlWorkbook.Close();
                XlWorkbook = null;
                XlApplication.Quit();
                XlApplication = null;
                GC.Collect();
            }
        }

        public void DeleteRange(int Row = 0, int Column = 0, int EndRow = 0, int EndColumn = 0) {
            if (Row > 0) { X = Row; }
            if (Column > 0) { Y = Column; }
            string r = IndexToLetter(Y) + X.ToString();
            if (EndRow > 0 && EndColumn > 0) {
                r = r + ":" + IndexToLetter(EndColumn) + EndRow.ToString();
            }
            XlRange = XlApplication.Range[r];
            XlRange.Delete(-4162);
        }

        public string GetCell(int Row = 0, int Column = 0) {
            if (Row > 0) { X = Row; }
            if (Column > 0) { Y = Column; }
            string t = (string)(XlWorksheet.Cells[X, Y] as Excel.Range).Text;
            if (t != null) return t;
            return string.Empty;
        }

        public string GetNextCell() {
            return GetCell(X, ++Y);
        }

        public void InsertRow(int row = 1) {
            (XlWorksheet.Rows[row] as Excel.Range).Insert();
        }

        public void RenameTab(string Name) {
            XlWorksheet.Name = Name;
        }

        public void RunMacro(string Name) {
            XlApplication.Run(Name);
        }

        public bool Save() {
            try {
                XlWorkbook.Save();
            } catch (Exception) {
                MsgboxOrWrite("File not writable, please close Excel first!" + Environment.NewLine + "A fájl nem írható, zárd be az Excelt mentés előtt!", "Close Excel!", System.Windows.MessageBoxImage.Information);
                return false;
            }
            return true;
        }

        public bool SaveAs(string Path, bool forceXLS = true) {
            try {
                if (forceXLS) {
                    XlWorkbook.SaveAs(Path, FileFormat: Excel.XlFileFormat.xlWorkbookNormal, AccessMode: Excel.XlSaveAsAccessMode.xlNoChange);
                } else {
                    XlWorkbook.SaveAs(Path, AccessMode: Excel.XlSaveAsAccessMode.xlNoChange);
                }
            } catch (Exception) {
                MsgboxOrWrite("File not writable, please close Excel first!" + Environment.NewLine + "A fájl nem írható, zárd be az Excelt mentés előtt!", "Close Excel!", System.Windows.MessageBoxImage.Information);
                return false;
            }
            return true;
        }

        public void SelectWorksheet(string Name) {
            XlWorksheet = XlWorkbook.Worksheets[Name] as Excel.Worksheet;
            XlWorksheet.Activate();
            X = 1;
            Y = 1;
        }

        public void SetCell(string Val, int Row = 0, int Column = 0) {
            if (Row > 0) { X = Row; }
            if (Column > 0) { Y = Column; }
            if (string.IsNullOrEmpty(Val)) {
                XlWorksheet.Cells[X, Y] = string.Empty;
            } else {
                XlWorksheet.Cells[X, Y] = Val;
            }
        }

        public void SetCell(int Val, int Row = 0, int Column = 0) {
            if (Row > 0) { X = Row; }
            if (Column > 0) { Y = Column; }
            XlWorksheet.Cells[X, Y] = Val;
        }

        public void SetCell(int? Val, int Row = 0, int Column = 0) {
            if (Row > 0) { X = Row; }
            if (Column > 0) { Y = Column; }
            XlWorksheet.Cells[X, Y] = Val.GetValueOrDefault();
        }

        public void SetFormula(string formula, int Row = 0, int Column = 0, int EndRow = 0, int EndColumn = 0) {
            if (Row > 0) { X = Row; }
            if (Column > 0) { Y = Column; }
            string r = IndexToLetter(Y) + X.ToString();
            if (EndRow > 0 && EndColumn > 0) {
                r = r + ":" + IndexToLetter(EndColumn) + EndRow.ToString();
            }
            XlRange = XlApplication.Range[r];
            XlRange.Formula = formula;
        }

        public void SetFormat(int Row = 0, int Column = 0, int EndRow = 0, int EndColumn = 0, bool? Bold = null, bool? Italic = null, int Size = 0, bool AutoFit = false, bool? WrapText = null, Align Align = Align.Left, int BGColor = -1, string FontName = "") {
            if (Row > 0) { X = Row; }
            if (Column > 0) { Y = Column; }
            string r = IndexToLetter(Y) + X.ToString();
            if (EndRow > 0 && EndColumn > 0) {
                r = r + ":" + IndexToLetter(EndColumn) + EndRow.ToString();
            }

            switch (Align) {
                case Align.Left:
                    XlRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    break;
                case Align.Right:
                    XlRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    break;
                case Align.Center:
                    XlRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    break;
            }

            if (Size > 0) XlRange.Font.Size = Size;
            if (AutoFit) XlRange.EntireColumn.AutoFit();
            if (BGColor != -1) XlRange.Interior.ColorIndex = BGColor;
            if (!string.IsNullOrEmpty(FontName)) XlRange.Font.Name = FontName;

            if (WrapText != null) XlRange.WrapText = WrapText;
            XlRange = XlApplication.Range[r];
            if (Bold != null) XlRange.Font.Bold = Bold;
            if (Italic != null) XlRange.Font.Italic = Italic;
        }

        public void SetBorder(int Row = 0, int Column = 0, int EndRow = 0, int EndColumn = 0, int BorderWeight = 0, Border Border = Border.All) {
            if (Row > 0) { X = Row; }
            if (Column > 0) { Y = Column; }
            string r = IndexToLetter(Y) + X.ToString();
            if (EndRow > 0 && EndColumn > 0) {
                r = r + ":" + IndexToLetter(EndColumn) + EndRow.ToString();
            }
            XlRange = XlApplication.Range[r];
            switch (Border) {
                case Border.Top:
                    XlRange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = BorderWeight;
                    break;
                case Border.Bottom:
                    XlRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = BorderWeight;
                    break;
                case Border.Right:
                    XlRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = BorderWeight;
                    break;
                case Border.Left:
                    XlRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = BorderWeight;
                    break;
                case Border.All:
                    XlRange.Borders.Weight = BorderWeight;
                    break;
            }
        }

        public void SetNextCell(string Val) {
            Y++;
            SetCell(Val);
        }

        public void SetNextCell(int Val) {
            Y++;
            SetCell(Val);
        }

        public void SetNextCell(int? Val) {
            Y++;
            SetCell(Val);
        }

        public void SetNumberFormat(string NumberFormat, int Row = 0, int Column = 0, int EndRow = 0, int EndColumn = 0) {
            if (Row > 0) { X = Row; }
            if (Column > 0) { Y = Column; }
            string r = IndexToLetter(Y) + X.ToString();
            if (EndRow > 0 && EndColumn > 0) {
                r = r + ":" + IndexToLetter(EndColumn) + EndRow.ToString();
            }
            XlRange = XlApplication.get_Range(r);
            XlRange.NumberFormat = NumberFormat;
        }

        public void MergeCells(int Row = 0, int Column = 0, int EndRow = 0, int EndColumn = 0, bool Merge = true) {
            if (Row > 0) { X = Row; }
            if (Column > 0) { Y = Column; }
            string r = IndexToLetter(Y) + X.ToString();
            if (EndRow > 0 && EndColumn > 0) {
                r = r + ":" + IndexToLetter(EndColumn) + EndRow.ToString();
            }
            XlRange = XlApplication.Range[r];
            if (Merge) {
                XlRange.Merge();
            } else {
                XlRange.UnMerge();
            }
        }

        public void PageSetup(Orientations Orientation = Orientations.Default, double LeftMargin = 50.0, double RightMargin = 50.0, double TopMargin = 55.0, double BottomMargin = 55.0, FooterPosition FooterPosition = FooterPosition.Default, string FooterText = "", HeaderPosition HeaderPosition = HeaderPosition.Default, string HeaderText = "") {
            Excel.PageSetup pageSetup = XlWorksheet.PageSetup;
            if (Orientation == Orientations.Portrait) {
                pageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;
            } else if (Orientation == Orientations.Landscape) {
                pageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
            }
            pageSetup.LeftMargin = LeftMargin;
            pageSetup.RightMargin = RightMargin;
            pageSetup.TopMargin = TopMargin;
            pageSetup.BottomMargin = BottomMargin;
            if (!string.IsNullOrEmpty(FooterText)) {
                switch (FooterPosition) {
                    case FooterPosition.Left:
                        pageSetup.LeftFooter = FooterText;
                        break;
                    case FooterPosition.Right:
                        pageSetup.RightFooter = FooterText;
                        break;
                    case FooterPosition.Center:
                        pageSetup.CenterFooter = FooterText;
                        break;
                }
            }
            if (!string.IsNullOrEmpty(HeaderText)) {
                switch (HeaderPosition) {
                    case HeaderPosition.Left:
                        pageSetup.LeftHeader = HeaderText;
                        break;
                    case HeaderPosition.Right:
                        pageSetup.RightHeader = HeaderText;
                        break;
                    case HeaderPosition.Center:
                        pageSetup.CenterHeader = HeaderText;
                        break;
                }
            }
        }

        public void SetColumnWidth(int Column, int Width, int EndColumn = 0) {
            Y = Column;
            string r = IndexToLetter(Y) + X.ToString();
            if (EndColumn > 0) {
                r = r + ":" + IndexToLetter(EndColumn) + X.ToString();
            }
            XlRange = XlApplication.get_Range(r);
            XlRange.EntireColumn.ColumnWidth = Width;
        }

        public void CopyFromDt(System.Data.DataTable dt) {
            string[,] ArrayDt = new string[dt.Rows.Count + 1 + 1, dt.Columns.Count - 1 + 1];
            int EndRow = dt.Rows.Count;
            int EndColumn = dt.Columns.Count;
            string r = IndexToLetter(Y) + X.ToString() + ":" + IndexToLetter(EndColumn) + (EndRow + 1).ToString();
            for (int i = 0; i < dt.Columns.Count; i++) {
                ArrayDt[0, i] = dt.Columns[i].Caption.ToString();
            }
            XlRange = XlApplication.Range[r];
            for (int j = 0; j < dt.Rows.Count; j++) {
                for (int k = 0; k < dt.Columns.Count; k++) {
                    if (dt.Rows[j][k] == DBNull.Value) {
                        ArrayDt[j + 1, k] = string.Empty;
                    } else {
                        ArrayDt[j + 1, k] = dt.Rows[j][k].ToString();
                    }
                }
            }
            try {
                XlRange.Value = ArrayDt;
            } catch (Exception ex) {
                if (!ex.Message.Contains("HRESULT")) {
                    MsgboxOrWrite(ex.Message, "Exception", System.Windows.MessageBoxImage.Error);
                }
            }
        }

        public System.Data.DataTable CopyToDt(int rownum = 0, int colnum = 0) {
            if (rownum == 0) { rownum = this.GetRowCount; }
            if (colnum == 0) { colnum = this.GetColumnCount; }
            string r = IndexToLetter(1) + 1.ToString();
            if (rownum > 0 && colnum > 0) r += ":" + IndexToLetter(colnum) + rownum.ToString();
            XlRange = XlApplication.Range[r];
            System.Data.DataTable dt = new System.Data.DataTable();
            object[,] ArrayRange = (object[,])XlRange.Value;
            for (int i = 1; i <= ArrayRange.GetLength(0); i++) {
                for (int j = 1; j <= ArrayRange.GetLength(1); j++) {
                    if (i == 1) {
                        if (ArrayRange[i, j] == null) {
                            dt.Columns.Add("COLUMN" + j);
                        } else {
                            dt.Columns.Add(ArrayRange[1, j].ToString());
                        }
                    } else {
                        if (j == 1) {
                            dt.Rows.Add();
                        }
                        if (ArrayRange[i, j] == null) {
                            dt.Rows[i - 2][j - 1] = string.Empty;
                        } else {
                            dt.Rows[i - 2][j - 1] = ArrayRange[i, j].ToString();
                        }
                    }
                }
            }
            return dt;
        }

        public void Print(int Copies = 1, bool Preview = false) {
            try {
                XlWorksheet.PrintOut(Copies: Copies, Preview: Preview);
            } catch (Exception) {
                XlWorksheet.PrintOutEx(Copies: Copies, Preview: Preview);
            }
        }

        public void ToNextCell(int Columns = 1) {
            Y += Columns;
        }

        public void ToNewLine(int Rows = 1) {
            X += Rows;
            Y = 1;
        }

        public static string IndexToLetter(int colIndex) {
            if (colIndex < 1) return "A";
            int div = colIndex;
            string colLetter = string.Empty;
            while (div > 0) {
                int modnum = (div - 1) % 26;
                colLetter = (char)(65 + modnum) + colLetter;
                div = (div - modnum) / 26;
            }
            return colLetter;
        }

        #region Internal functions
        ~HandleExcel() { Close(); }
        public void Dispose() { Close(); }

        private void MsgboxOrWrite(string text, string title = "", System.Windows.MessageBoxImage icon = System.Windows.MessageBoxImage.None) {
            if (is_console_app) {
                Console.WriteLine(text);
            } else {
                System.Windows.MessageBox.Show(text, title, System.Windows.MessageBoxButton.OK, icon);
            }
        }
        #endregion
    }
}
