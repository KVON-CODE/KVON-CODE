
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Path = System.IO.Path;
using SpreadsheetLight;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Security.Cryptography;

namespace PayrollAnalyzer
{
    /// <summary>
    /// CORE FORENSIC ENGINE: Developed to identify time-series discrepancies and 
    /// potential compliance violations across high-volume payroll datasets.
    /// </summary>
    public partial class MainWindow : Window
    {
        // GLOBAL STATE PERSISTENCE: Tracking source paths for forensic logging.
        public string thefilepath = "";
        public string thefilename = "";

        /// <summary>
        /// DATA INGESTION LAYER: Interfaces with legacy Excel formats via OLEDB.
        /// Extracts raw data for normalization and auditing.
        /// </summary>
        public DataTable ReadExcel(string fileName, string fileExt)
        {
            // LEGACY COMPATIBILITY: Utilizing JET engine for high-fidelity data extraction from .xls/xlsx.
            string CONNECTION_STRING = "Provider=Microsoft.Jet.OLEDB.4.0;;Data Source=" +
                                fileName +
                                ";Extended Properties=\"Excel 8.0;HDR=NO';";

            OleDbConnection objConnection = new OleDbConnection(CONNECTION_STRING);
            DataSet dsImport = new DataSet();

            objConnection.Open();
            DataTable dtSchema = objConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

            // SCHEMA DISCOVERY: Dynamically identifying the target worksheet for processing.
            string firstSheetName = dtSchema.Rows[0]["TABLE_NAME"].ToString();
            MessageBox.Show(firstSheetName);
            new OleDbDataAdapter("SELECT * FROM [" + firstSheetName + "]", objConnection).Fill(dsImport);

            if (objConnection != null)
            {
                objConnection.Close();
                objConnection.Dispose();
            }

            return (dsImport.Tables.Count > 0) ? dsImport.Tables[0] : null;
        }

        // OBJECT-ORIENTED DATA MODEL: Encapsulating daily time-series data for chronological analysis.
        xDayArr ListDays = new xDayArr();

        public MainWindow()
        {
            InitializeComponent();
            ListDays.DayList = new List<xDay>();
        }

        class xDayArr
        {
            public List<xDay> DayList { get; set; }
        }

        /// <summary>
        /// DATA ENTITY: Represents a single 24-hour audit window.
        /// Includes minute-level tracking (TSpan) and hourly attribution (xHours).
        /// </summary>
        class xDay
        {
            public string xDate { get; set; }
            public List<string> xHours { get; set; }
            public List<string> XDays { get; set; }
            public DateTime XDT { get; set; }
            public int TSpan { get; set; }
        }

        private void LoadHours(xDay loadDay, string time1, string time2, string time3, string time4)
        {
            // MERIDIAN VALIDATION: Preparing time strings for AM/PM standardization.
            if (time1.Contains("PM")) { }
            if (time2.Contains("PM")) { }
            if (time3.Contains("PM")) { }
            if (time4.Contains("PM")) { }
        }

        private dynamic TimeMath(DateTime Time1, DateTime Time2)
        {
            // TEMPORAL DELTA CALCULATION: Quantifying work duration for specific shifts.
            var timemath = Time2 - Time1;
            return timemath;
        }

        /// <summary>
        /// HOURLY ATTRIBUTION ENGINE: Distributes continuous shift data into discrete 
        /// 1-hour audit blocks while managing the "Midnight Crossover" edge case.
        /// </summary>
        private void LoadBlocks(DateTime Time1, DateTime Time2, String pType)
        {
            try
            {
                // DATA PARSING: Deconstructing time objects into manageable integer units.
                int timehour1 = Int32.Parse(Time1.ToString("HH"));
                int timehour2 = Int32.Parse(Time2.ToString("HH"));
                int timemin1 = Int32.Parse(Time1.ToString("mm"));
                int timemin2 = Int32.Parse(Time2.ToString("mm"));

                string timehourstr1 = Time1.ToString("hh:mm tt");
                string timehourstr2 = Time2.ToString("hh:mm tt");

                // CROSS-DAY LINKING: Ensuring sequential shifts across midnight remain chronologically accurate.
                var TheDay = ListDays.DayList.First(r => r.XDT == Time1.Date);
                var listpos1 = ListDays.DayList.FindIndex(r => r == TheDay);
                var TheDay2 = ListDays.DayList[listpos1];
                if ((listpos1 + 1) < ListDays.DayList.Count) TheDay2 = ListDays.DayList[listpos1 + 1];

                var timemath = Time2 - Time1;
                var useday = TheDay;
                var timemathleft = timemath;

                // LOGIC BRANCH: Handling single-hour shifts vs multi-hour/multi-day shifts.
                if (timehour1 == timehour2)
                {
                    useday.xHours[timehour1] += " " + timehourstr1 + "-" + timehourstr2 + " " + pType;
                    useday.TSpan += (timemin2 - timemin1);
                }
                else
                {
                    var timeend = Time1.Add(timemathleft);
                    useday.xHours[timehour1] += " " + timehourstr1 + " " + pType;
                    useday.TSpan += (60 - timemin1);

                    // YEARLY CALENDAR INITIALIZATION: Maintaining a full-year audit context.
                    var beginOfYear = new DateTime(Time1.Year, 01, 01, 0, 0, 0, DateTimeKind.Local);
                    var endOfYear = beginOfYear.AddYears(1);
                    var daysOfYear = endOfYear.Subtract(beginOfYear).TotalDays;
                    DateTime tstart = Time1.AddMinutes(-(Time1.Minute)).AddHours(1);
                    DateTime tend = Time2.AddMinutes(-(Time2.Minute)).AddHours(-1);
                    DateTime tcur = tstart;
                    bool done = false;
                    int p = 0;
                    bool yearend = false;
                    DateTime tempcur = tstart;

                    // INCREMENTAL PROCESSING LOOP: Iterating through each hour of a multi-hour shift.
                    do
                    {
                        tempcur = tstart.AddHours(p);
                        if (tempcur.Year > Time1.Year) { yearend = true; done = true; }
                        if (tempcur > tend) { done = true; }
                        else
                        {
                            // SHIFT TRANSITION: Automatically migrating data to the next day object if midnight is crossed.
                            if ((tempcur.DayOfYear > Time1.DayOfYear))
                            {
                                if ((TheDay2 != TheDay) && (tempcur.DayOfYear <= daysOfYear)) { useday = TheDay2; }
                                else { yearend = true; done = true; }
                            }

                            if (yearend == false) { useday.xHours[tempcur.Hour] += pType; useday.TSpan += 60; }
                            p++;
                        }
                    } while (done == false);

                    if (Time2.DayOfYear > Time1.DayOfYear)
                    {
                        if ((TheDay2 != TheDay) && (Time2.DayOfYear <= daysOfYear)) { useday = TheDay2; }
                    }
                    if (yearend == false) { useday.xHours[timehour2] += timehourstr2 + " " + pType; useday.TSpan += (timemin2); }
                }
            }
            catch (Exception ex)
            {
                // FORENSIC EXCEPTION LOGGING: Capturing failed rows for manual audit.
                MessageBox.Show(Time1.ToString() + Time2.ToString());
            }
        }

        private void Log(string message)
        {
            // SYSTEM LOGGING: Replaces pop-up UI with a persistent diagnostic console.
            LogBox.AppendText($"[{DateTime.Now.ToString("HH:mm:ss")}] {message}\r\n");
            LogBox.ScrollToEnd(); // Automatically follows the latest data
        }
        private void Log(string message, bool clear)
        {
            if (clear = true) LogBox.Clear();
            // SYSTEM LOGGING: Replaces pop-up UI with a persistent diagnostic console.
            LogBox.AppendText($"[{DateTime.Now.ToString("HH:mm:ss")}] {message}\r\n");
            LogBox.ScrollToEnd(); // Automatically follows the latest data
        }

        /// <summary>
        /// ADP FILE INTERFACE: Specialized logic for parsing standard ADP payroll exports.
        /// Implements custom Regex to clean system-generated "dirty" strings.
        /// </summary>
        private void LOADADP_Click(object sender, RoutedEventArgs e)
        {
            string QUOTA = "";
            string filePath = string.Empty;
            string fileExt = string.Empty;
            List<xDay> xDays = new List<xDay>();
            OpenFileDialog file = new OpenFileDialog();

            if (file.ShowDialog() == true)
            {
                filePath = file.FileName;
                fileExt = Path.GetExtension(filePath);
                thefilepath = Path.GetDirectoryName(filePath);
                Log("========================================================================================", true);
                Log("ADP File Selected: " + Path.GetFileName(file.FileName));
                Log("========================================================================================");
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {
                    try
                    {
                        DataTable dtExcel = new DataTable();
                        SLDocument sl = new SLDocument(filePath);
                        SLWorksheetStatistics stats = sl.GetWorksheetStatistics();
                        CURADP.Content = filePath;
                        thefilename = Path.GetFileName(filePath);

                        bool use = false;
                        xDay tempDate = new xDay();
                        var curDate = tempDate;
                        int curYear = int.Parse(YearText.Text);

                        // DATE RANGE PROTECTION: Establishing the yearly audit boundaries.
                        var beginOfYear = new DateTime(curYear, 01, 01, 0, 0, 0, DateTimeKind.Local);
                        var endOfYear = beginOfYear.AddYears(1);
                        var daysOfYear = endOfYear.Subtract(beginOfYear).TotalDays;
                        sl.RemoveAllPageBreaks();
                        var datenotfirst = 0;

                        DateTime startday = beginOfYear.AddDays(-1);
                        int firstdate = 0;
                        ListDays.DayList.Clear();

                        // AUDIT GRID INITIALIZATION: Creating a blank 365/366-day matrix.
                        for (int i = 1; i <= ((Int32)daysOfYear); i++)
                        {
                            xDay newDate = new xDay();
                            DateTime thisday = startday.AddDays(i);
                            newDate.xHours = new List<string> { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
                            newDate.xDate = startday.Date.ToString("MM/dd");
                            newDate.XDT = thisday;
                            ListDays.DayList.Add(newDate);
                        }

                        if (curYear > 2000 && curYear < 3000)
                        {
                            string fday = "";
                            string fmonth = "";
                            string fDay1 = "";
                            int nextday = 0;
                            bool datefound = false;

                            // MAIN IMPORT LOOP: Iterating through ADP source rows.
                            for (int i = 1; i < (stats.NumberOfRows + 1); i++)
                            {
                                string fDay = sl.GetCellValueAsString("A" + i);

                                // HEADER DETECTION: Identifying start of data blocks within complex system exports.
                                if (fDay.Contains("Date"))
                                {
                                    if (datenotfirst != 1)
                                    {
                                        datefound = true;
                                        nextday = 0;
                                        use = false;
                                        fday = ""; fmonth = ""; fDay1 = "";
                                        curDate = tempDate;
                                        datenotfirst = 1;
                                    }
                                    else { 
                                        Log("Date carryover: " + sl.GetCellValueAsString("A" + (i + 1)));
                                    }
                                }
                                else
                                {
                                    // DATA SANITIZATION (REGEX): Stripping non-essential characters from source cells.
                                    fDay1 = Regex.Replace(sl.GetCellValueAsString("A" + i), "[^0-9/]", "");
                                    string v1 = sl.GetCellValueAsString("B" + i);
                                    string v2 = sl.GetCellValueAsString("C" + i);
                                    string v3 = sl.GetCellValueAsString("D" + i);
                                    string v4 = sl.GetCellValueAsString("E" + i);

                                    // SHIFT PATTERN MATCHING: Identifying time entries via standard Regex masks.
                                    if (Regex.Matches(fDay1, "[0-1]?[0-9]/[0-3]?[0-9]").Count > 0)
                                    {
                                        firstdate = 1;
                                        try
                                        {
                                            string ft1, ft2, ft3, ft4 = "";
                                            ft1 = Regex.Replace(sl.GetCellValueAsString("B" + i), "[^0-9:AP]", "");
                                            ft2 = Regex.Replace(sl.GetCellValueAsString("C" + i), "[^0-9:AP]", "");
                                            ft3 = Regex.Replace(sl.GetCellValueAsString("D" + i), "[^0-9:AP]", "");
                                            ft4 = Regex.Replace(sl.GetCellValueAsString("E" + i), "[^0-9:AP]", "");

                                            fday = fDay1.Split('/')[1];
                                            fmonth = fDay1.Split('/')[0];
                                            var fdate = new DateTime(curYear, Int32.Parse(fmonth), Int32.Parse(fday), 0, 0, 0);

                                            curDate = ListDays.DayList.First(r => r.XDT == fdate);

                                            if (curDate.XDT.DayOfYear > 0 && curDate.XDT.Day <= 366)
                                            {
                                                use = true;
                                                nextday = 0;

                                                // INPUT PAIR VALIDATION: Ensuring both Start and End times are present for a shift.
                                                if ((v1.Length > 3 && v2.Length > 3) || (v3.Length > 3 && v4.Length > 3))
                                                {
                                                    var cultureInfo = new CultureInfo("en-US");
                                                    var regmatch1 = Regex.Matches(ft1, "[0-1]?[0-9]:[0-5][0-9][AP]");
                                                    var regmatch2 = Regex.Matches(ft2, "[0-1]?[0-9]:[0-5][0-9][AP]");
                                                    var regmatch3 = Regex.Matches(ft3, "[0-1]?[0-9]:[0-5][0-9][AP]");
                                                    var regmatch4 = Regex.Matches(ft4, "[0-1]?[0-9]:[0-5][0-9][AP]");

                                                    // BLOCK 1 PROCESSING (Columns B-C)
                                                    if (regmatch1.Count > 0 && regmatch2.Count > 0)
                                                    {
                                                        string rm1 = regmatch1[0].ToString() + "M";
                                                        string rm2 = regmatch2[0].ToString() + "M";

                                                        string FT1str = fmonth + "/" + fday + "/" + curYear + " " + rm1;
                                                        string FT2str = fmonth + "/" + fday + "/" + curYear + " " + rm2;

                                                        DateTime FT1 = DateTime.Parse(FT1str);
                                                        DateTime FT2 = DateTime.Parse(FT2str);

                                                        // AUTOMATIC CHRONOLOGICAL ADJUSTMENT: Correcting for logic where End < Start.
                                                        if (rm1.Contains("A") && rm2.Contains("A") && DateTime.Compare(FT1, FT2) >= 0) { nextday++; FT2 = FT2.AddDays(nextday); }
                                                        else if (rm1.Contains("P") && rm2.Contains("A")) { nextday++; FT2 = FT2.AddDays(nextday); }

                                                        LoadBlocks(FT1, FT2, "DUTY");
                                                    }

                                                    // BLOCK 2 PROCESSING (Columns D-E)
                                                    if (regmatch3.Count > 0 && regmatch4.Count > 0)
                                                    {
                                                        string rm3 = regmatch3[0].ToString() + "M";
                                                        string rm4 = regmatch4[0].ToString() + "M";
                                                        string FT3str = fmonth + "/" + fday + "/" + curYear + " " + rm3;
                                                        string FT4str = fmonth + "/" + fday + "/" + curYear + " " + rm4;
                                                        DateTime FT3 = DateTime.Parse(FT3str);
                                                        DateTime FT4 = DateTime.Parse(FT4str);

                                                        if (nextday > 0) { FT3 = FT3.AddDays(nextday); FT4 = FT4.AddDays(nextday); }

                                                        if (rm3.Contains("A") && rm4.Contains("A") && DateTime.Compare(FT3, FT4) >= 0) { nextday++; FT4 = FT4.AddDays(nextday); }
                                                        else if (rm3.Contains("P") && rm4.Contains("A")) { nextday++; FT4 = FT4.AddDays(nextday); }

                                                        LoadBlocks(FT3, FT4, "DUTY");
                                                    }
                                                }
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            QUOTA += "\r\nDAY\r\n" + fDay + v1 + v2 + v3 + v4;
                                        }
                                    }
                                    // RECURSIVE PATTERN DETECTION: Capturing multi-line entries under a single date header.
                                    else if (use == true && firstdate == 1 && !fDay.Contains("Date"))
                                    {
                                        try
                                        {
                                            string ft1, ft2, ft3, ft4 = "";
                                            ft1 = Regex.Replace(sl.GetCellValueAsString("B" + i), "[^0-9:AP]", "");
                                            ft2 = Regex.Replace(sl.GetCellValueAsString("C" + i), "[^0-9:AP]", "");
                                            ft3 = Regex.Replace(sl.GetCellValueAsString("D" + i), "[^0-9:AP]", "");
                                            ft4 = Regex.Replace(sl.GetCellValueAsString("E" + i), "[^0-9:AP]", "");

                                            var regmatch1 = Regex.Matches(ft1, "[0-1]?[0-9]:[0-5][0-9][AP]");
                                            var regmatch2 = Regex.Matches(ft2, "[0-1]?[0-9]:[0-5][0-9][AP]");
                                            var regmatch3 = Regex.Matches(ft3, "[0-1]?[0-9]:[0-5][0-9][AP]");
                                            var regmatch4 = Regex.Matches(ft4, "[0-1]?[0-9]:[0-5][0-9][AP]");

                                            if (regmatch1.Count > 0 && regmatch2.Count > 0)
                                            {
                                                string rm1 = regmatch1[0].ToString() + "M";
                                                string rm2 = regmatch2[0].ToString() + "M";

                                                string FT1str = fmonth + "/" + fday + "/" + curYear + " " + rm1;
                                                string FT2str = fmonth + "/" + fday + "/" + curYear + " " + rm2;
                                                DateTime FT1 = DateTime.Parse(FT1str);
                                                DateTime FT2 = DateTime.Parse(FT2str);

                                                if (nextday > 0) { FT1 = FT1.AddDays(nextday); FT2 = FT2.AddDays(nextday); }

                                                if (rm1.Contains("A") && rm2.Contains("A") && DateTime.Compare(FT1, FT2) >= 0) { nextday++; FT2 = FT2.AddDays(nextday); }
                                                else if (rm1.Contains("P") && rm2.Contains("A")) { nextday++; FT2 = FT2.AddDays(nextday); }

                                                LoadBlocks(FT1, FT2, "DUTY");
                                            }
                                            if (regmatch3.Count > 0 && regmatch4.Count > 0)
                                            {
                                                string rm3 = regmatch3[0].ToString() + "M";
                                                string rm4 = regmatch4[0].ToString() + "M";
                                                string FT3str = fmonth + "/" + fday + "/" + curYear + " " + rm3;
                                                string FT4str = fmonth + "/" + fday + "/" + curYear + " " + rm4;
                                                DateTime FT3 = DateTime.Parse(FT3str);
                                                DateTime FT4 = DateTime.Parse(FT4str);
                                                if (nextday > 0) { FT3 = FT3.AddDays(nextday); FT4 = FT4.AddDays(nextday); }

                                                if (rm3.Contains("A") && rm4.Contains("A") && DateTime.Compare(FT3, FT4) >= 0) { nextday++; FT4 = FT4.AddDays(nextday); }
                                                else if (rm3.Contains("P") && rm4.Contains("A")) { nextday++; FT4 = FT4.AddDays(nextday); }

                                                LoadBlocks(FT3, FT4, "DUTY");
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            QUOTA += "\r\nBLANK\r\n" + fDay + v1 + v2 + v3 + v4;
                                        }
                                    }
                                }
                            }
                            Log("SUCCESS: ADP data parsed. Total rows analyzed: " + stats.NumberOfRows);
                            Log("ISSUE LIST: " + QUOTA);

                            // AUDIT LOG GENERATION: Exporting processing results for independent verification.
                            string NEWFILENAME = "COMBINE-ADP-DETAILS-" + thefilename + ".txt";
                            File.WriteAllText(thefilepath + "/" + NEWFILENAME, QUOTA);

                        }
                        else { MessageBox.Show("Enter a valid year!"); }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("EXCEPTION: " + ex.Message.ToString());
                    }
                }
                else { MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButton.OK); }
            }
        }

        private int AMORPM(string AOP)
        {
            // MERIDIAN IDENTIFIER: Logic to differentiate morning vs evening shifts for forensic sorting.
            var AM = Regex.Matches(AOP, "[0-9]?[0-9]:[0-5][0-9][A]");
            var PM = Regex.Matches(AOP, "[0-9]?[0-9]:[0-5][0-9][P]");

            if (AM.Count > 0) return 1;
            else if (PM.Count > 0) return 2;
            else return 0;
        }

        /// <summary>
        /// DATA VISUALIZATION ENGINE: Generates a high-level forensic dashboard in Excel.
        /// Features color-coded conflict detection (Overlaps/Fraud markers).
        /// </summary>
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            SLDocument newdocument = new SLDocument();

            // DASHBOARD LAYOUT: Initializing the 24-hour visualization grid.
            newdocument.SetWorksheetDefaultColumnWidth(15);
            newdocument.SetWorksheetDefaultRowHeight(25);
            newdocument.SetCellValue("A1", "Date");
            newdocument.SetCellValue("B1", "12:00 AM");
            newdocument.SetCellValue("C1", "1:00 AM");
            newdocument.SetCellValue("D1", "2:00 AM");
            newdocument.SetCellValue("E1", "3:00 AM");
            newdocument.SetCellValue("F1", "4:00 AM");
            newdocument.SetCellValue("G1", "5:00 AM");
            newdocument.SetCellValue("H1", "6:00 AM");
            newdocument.SetCellValue("I1", "7:00 AM");
            newdocument.SetCellValue("J1", "8:00 AM");
            newdocument.SetCellValue("K1", "9:00 AM");
            newdocument.SetCellValue("L1", "10:00 AM");
            newdocument.SetCellValue("M1", "11:00 AM");
            newdocument.SetCellValue("N1", "12:00 PM");
            newdocument.SetCellValue("O1", "1:00 PM");
            newdocument.SetCellValue("P1", "2:00 PM");
            newdocument.SetCellValue("Q1", "3:00 PM");
            newdocument.SetCellValue("R1", "4:00 PM");
            newdocument.SetCellValue("S1", "5:00 PM");
            newdocument.SetCellValue("T1", "6:00 PM");
            newdocument.SetCellValue("U1", "7:00 PM");
            newdocument.SetCellValue("V1", "8:00 PM");
            newdocument.SetCellValue("W1", "9:00 PM");
            newdocument.SetCellValue("X1", "10:00 PM");
            newdocument.SetCellValue("Y1", "11:00 PM");
            newdocument.SetCellValue("Z1", "");
            newdocument.SetCellValue("AA1", "Hours");
            newdocument.SetCellValue("AB1", "Day");

            int rowstart = 1;
            SLConditionalFormatting cf;
            SLStyle style = newdocument.CreateStyle();
            style.Alignment.WrapText = true;

            // DASHBOARD COMPILATION: Iterating through processed days to populate the visual grid.
            foreach (var x1 in ListDays.DayList)
            {
                rowstart++;
                newdocument.SetCellValue("A" + rowstart, x1.XDT.ToString("MM/dd/yyyy"));

                List<string> letter = new List<string>();
                int thishour = 0;

                // COLUMN-SPECIFIC MAPPING: Distributing shift info across the A-Z axis.
                for (char c = 'B'; c <= 'Y'; c++)
                {
                    string thiscell = c.ToString() + rowstart.ToString();
                    string thisinfo = x1.xHours[thishour];

                    style.SetFontColor(System.Drawing.Color.Black);

                    // FORENSIC COLOR CODING: Visually flagging anomalies.
                    // RED: Multiple employment types in same hour (Potential Fraud/Overlap).
                    if (thisinfo.Contains("DUTY") && thisinfo.Contains("Detail")) { style.SetFontColor(System.Drawing.Color.Red); }

                    // PURPLE: Duplicate primary employment entries (System Error/Time Padding).
                    else if (Regex.Matches(thisinfo, "DUTY").Count > 1) { style.SetFontColor(System.Drawing.Color.Purple); }

                    // BLUE: Standard primary duty hours.
                    else if (thisinfo.Contains("DUTY")) { style.SetFontColor(System.Drawing.Color.Blue); }

                    // ORANGE/SLATE: Secondary employment (Detail) flagging.
                    else if (Regex.Matches(thisinfo, "Detail").Count > 1) { style.SetFontColor(System.Drawing.Color.Orange); }
                    else if (thisinfo.Contains("Detail")) { style.SetFontColor(System.Drawing.Color.SlateGray); }
                    else { style.SetFontColor(System.Drawing.Color.Black); }

                    // STRING BEAUTIFICATION: Removing metadata before final report generation.
                    thisinfo = thisinfo.Replace(" AM ", "").Replace(" PM ", "").Replace(" PM", "").Replace(" AM", "").Replace("PM", "").Replace("AM", "");

                    newdocument.SetCellValue(thiscell, thisinfo);
                    newdocument.SetCellStyle(thiscell, style);
                    thishour++;
                }

                // TOTAL HOURS CALCULATION: Quantifying daily output for audit summaries.
                newdocument.SetCellValue("AA" + rowstart, new TimeSpan(0, x1.TSpan, 0).ToString(@"hh\:mm"));

                // CRITICAL OVERTIME FLAG: Marking shifts exceeding 16.5 hours in RED (Safety/Compliance check).
                if (new TimeSpan(0, x1.TSpan, 0) > new TimeSpan(16, 35, 0))
                {
                    style.SetFontColor(System.Drawing.Color.Red);
                    newdocument.SetCellStyle("AA" + rowstart, style);
                }
                newdocument.SetCellValue("AB" + rowstart, x1.XDT.ToString("ddd"));
            }
            try
            {
                string NEWFILENAME = "COMBINE-ADP-DETAILS-" + thefilename;
                newdocument.SaveAs(thefilepath + "\\" + NEWFILENAME);
                newdocument.Dispose();
                Log("========================================================================================");
                Log("Audit Spreadsheet Saved: " + thefilepath + "\\" + NEWFILENAME);
                Log("========================================================================================");
            }
            catch
            {
                MessageBox.Show("Error while trying to save. See if file is open by another program. Then try again.");
                newdocument.Dispose();
            }
        }

        /// <summary>
        /// DETAIL FILE INTERFACE: Specialized parser for secondary employment data.
        /// Supports multiple formatting standards found in manual "Detail" logs.
        /// </summary>
        private void LOADDET_Click(object sender, RoutedEventArgs e)
        {
            string filePath = string.Empty;
            string fileExt = string.Empty;
            List<xDay> xDays = new List<xDay>();
            OpenFileDialog file = new OpenFileDialog();

            if (file.ShowDialog() == true)
            {
                filePath = file.FileName;
                fileExt = Path.GetExtension(filePath);
                thefilepath = Path.GetDirectoryName(filePath);
                CURDET.Content = filePath;
                Log("========================================================================================");
                Log("Detail File Selected: " + Path.GetFileName(file.FileName));
                Log("========================================================================================");
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {
                    try
                    {
                        DataTable dtExcel = new DataTable();
                        SLDocument sl = new SLDocument(filePath);
                        SLWorksheetStatistics stats = sl.GetWorksheetStatistics();

                        xDay tempDate = new xDay();
                        var curDate = tempDate;
                        int curYear = int.Parse(YearText.Text);
                        var beginOfYear = new DateTime(curYear, 01, 01, 0, 0, 0, DateTimeKind.Local);
                        var endOfYear = beginOfYear.AddYears(1);
                        var daysOfYear = endOfYear.Subtract(beginOfYear).TotalDays;
                        DateTime startday = beginOfYear.AddDays(-1);

                        if (curYear > 2000 && curYear < 3000)
                        {
                            int nextday = 0;
                            bool datefound = false;

                            // FORMAT DISCOVERY: Identifying which "Detail" template is being processed.
                            string howto = sl.GetCellValueAsString("A1");
                            Log("Type of detail format: " + howto);

                            for (int i = 2; i < (stats.NumberOfRows + 1); i++)
                            {
                                try
                                {
                                    string d1 = sl.GetCellValueAsString("A" + i);

                                    // FORMAT 1: Combined DateTime string (e.g., "1/1 10:00 - 1/1 14:00")
                                    if (howto.Contains("DATETIME-DATETIME"))
                                    {
                                        var fdt1 = d1.Split('-')[0];
                                        var fdt2 = d1.Split('-')[1];
                                        var datetime1 = DateTime.Parse(fdt1);
                                        var datetime2 = DateTime.Parse(fdt2);
                                        DateTime FT1 = datetime1;
                                        DateTime FT2 = datetime2;

                                        curDate = ListDays.DayList.First(r => r.XDT == FT1.Date);
                                        nextday = 0;
                                        if (curDate != tempDate) { LoadBlocks(FT1, FT2, "Detail"); }
                                    }
                                    // FORMAT 2: Service Date with explicit range (e.g., "1/1", "10:00-14:00")
                                    else if (howto.Contains("servdate"))
                                    {
                                        DateTime fDay = sl.GetCellValueAsDateTime("A" + i);
                                        string fTime = Regex.Replace(sl.GetCellValueAsString("B" + i), "[^0-9:-]", "");
                                        string fTime1 = fTime.Split('-')[0];
                                        string fTime2 = fTime.Split('-')[1].Replace("24:", "0:");

                                        var fdate = fDay;
                                        curDate = ListDays.DayList.First(r => r.XDT == fdate);
                                        nextday = 0;
                                        if (curDate != tempDate)
                                        {
                                            DateTime FT1 = fDay.Add(TimeSpan.Parse(fTime1));
                                            DateTime FT2 = fDay.Add(TimeSpan.Parse(fTime2));

                                            // CHRONOLOGICAL REPAIR: Logic for shifts ending in early AM hours.
                                            if (FT1.Hour < 12 && FT2.Hour < 12 && DateTime.Compare(FT1, FT2) >= 0) { nextday++; FT2 = FT2.AddDays(nextday); }
                                            else if (FT1.Hour >= 12 && FT2.Hour < 12) { nextday++; FT2 = FT2.AddDays(nextday); }

                                            LoadBlocks(FT1, FT2, "Detail");
                                        }
                                    }
                                    // FORMAT 3: Distinct start/end columns.
                                    else if (howto.Contains("Actual"))
                                    {
                                        DateTime FT1 = sl.GetCellValueAsDateTime("A" + i);
                                        DateTime FT2 = sl.GetCellValueAsDateTime("B" + i);
                                        curDate = ListDays.DayList.First(r => r.XDT == FT1.Date);
                                        nextday = 0;
                                        if (curDate != tempDate) { LoadBlocks(FT1, FT2, "Detail"); }
                                    }
                                    else { MessageBox.Show(howto.ToString() + " -  not valid detail first row."); }
                                }
                                catch { }
                            }
                            Log("Detail input complete");
                           
                        }
                        else { MessageBox.Show("Enter a valid year!"); }
                    }
                    catch (Exception ex) { }
                }
                else { MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButton.OK); }
            }
        }
    }
}