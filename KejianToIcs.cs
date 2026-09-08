using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Drawing.Drawing2D;

// 课表 docx -> Outlook .ics 转换器（WinForms）
namespace KejianToIcs
{
    public class CourseEntry
    {
        public string Name = "";
        public int Weekday = 1;      // 1=周一 ... 7=周日
        public string StartTime = ""; // HH:mm
        public string EndTime = "";   // HH:mm
        public int WeekStart = 1;
        public int WeekEnd = 17;
        public string Location = "";
        public string Teacher = "";
        public string Code = "";
    }

    // ---------------- docx 解析 ----------------
    public static class DocxParser
    {
        static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        // 读取 docx 中第一个含“星期”表头的表格，返回矩阵（每行是字符串数组）
        public static List<string[]> LoadMatrix(string docxPath)
        {
            string xml = ReadDocumentXml(docxPath);
            XDocument doc = XDocument.Parse(xml);
            var tables = doc.Descendants(W + "tbl").ToList();
            if (tables.Count == 0) throw new Exception("文档里没有找到表格。");

            List<string[]> matrix = new List<string[]>();
            foreach (var tbl in tables)
            {
                matrix.Clear();
                foreach (var tr in tbl.Elements(W + "tr"))
                {
                    var cells = new List<string>();
                    foreach (var tc in tr.Elements(W + "tc"))
                    {
                        int span = 1;
                        var tcPr = tc.Element(W + "tcPr");
                        if (tcPr != null)
                        {
                            var gs = tcPr.Element(W + "gridSpan");
                            if (gs != null)
                            {
                                var v = gs.Attribute(W + "val");
                                if (v != null) { int s; if (int.TryParse(v.Value, out s) && s > 0) span = s; }
                            }
                        }
                        string txt = Normalize(string.Concat(tc.Descendants(W + "t").Select(t => t.Value)));
                        for (int k = 0; k < span; k++) cells.Add(txt);
                    }
                    matrix.Add(cells.ToArray());
                }
                foreach (var row in matrix)
                {
                    if (row.Any(c => c.Contains("星期一"))) return matrix;
                }
            }
            throw new Exception("没有找到含“星期一…星期日”表头的课表。");
        }

        static string ReadDocumentXml(string docxPath)
        {
            using (var zip = ZipFile.OpenRead(docxPath))
            {
                var entry = zip.GetEntry("word/document.xml");
                if (entry == null) throw new Exception("不是有效的 .docx（缺少 word/document.xml）。");
                using (var r = new StreamReader(entry.Open(), Encoding.UTF8))
                {
                    return r.ReadToEnd();
                }
            }
        }

        static string Normalize(string s)
        {
            if (s == null) return "";
            s = s.Replace('\u00A0', ' ').Replace('\u3000', ' ');
            s = s.Replace("\r", " ").Replace("\n", " ");
            s = Regex.Replace(s, @"\s+", " ");
            return s.Trim();
        }

        static Dictionary<int, int> FindWeekdayColumns(List<string[]> matrix)
        {
            string[] names = { "星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日" };
            int header = -1;
            for (int i = 0; i < matrix.Count; i++)
            {
                if (matrix[i].Any(c => c.Contains("星期一"))) { header = i; break; }
            }
            if (header < 0) throw new Exception("找不到表头行。");

            var map = new Dictionary<int, int>();
            var head = matrix[header];
            for (int col = 0; col < head.Length; col++)
            {
                for (int w = 0; w < names.Length; w++)
                {
                    if (head[col].Contains(names[w])) { map[col] = w + 1; break; }
                }
            }
            return map;
        }

        static string GetCell(List<string[]> m, int row, int col)
        {
            if (row < 0 || row >= m.Count) return "";
            if (col < 0 || col >= m[row].Length) return "";
            return m[row][col];
        }

        public static List<CourseEntry> Parse(string docxPath)
        {
            var m = LoadMatrix(docxPath);
            var cols = FindWeekdayColumns(m);
            var result = new List<CourseEntry>();
            for (int r = 0; r < m.Count; r++)
            {
                foreach (var kv in cols)
                {
                    string cell = GetCell(m, r, kv.Key);
                    if (string.IsNullOrEmpty(cell)) continue;
                    if (cell.Contains("星期")) continue;
                    if (kv.Key >= m[r].Length) continue;
                    if (!cell.Contains("教学班代码") && Regex.IsMatch(cell, @"^第?\d+节")) continue;
                    result.AddRange(ParseCell(cell, kv.Value));
                }
            }
            return result;
        }

        static List<CourseEntry> ParseCell(string text, int weekday)
        {
            var list = new List<CourseEntry>();
            string t = Normalize(text);
            if (t.Length == 0) return list;

            if (!t.Contains("教学班代码"))
            {
                list.Add(new CourseEntry { Name = t, Weekday = weekday, WeekStart = 1, WeekEnd = 17 });
                return list;
            }

            string[] parts = t.Split(new string[] { "教学班代码" }, StringSplitOptions.None);
            string nextName = parts[0].Trim();
            for (int i = 1; i < parts.Length; i++)
            {
                string seg = parts[i];
                string code = "", tailName = "";
                int ws = 0, we = 0;
                string t1 = "", t2 = "", after = "";

                var m = Regex.Match(seg,
                    @"^[：:]?\s*([A-Za-z0-9_#]+)\s*\((\d{1,2})\s*[~\-]\s*(\d{1,2})周\)\s*\((\d+)-(\d+)节\s*(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})\)\s*(.*?)(?=人数|$)");

                if (m.Success)
                {
                    code = m.Groups[1].Value;
                    int.TryParse(m.Groups[2].Value, out ws);
                    int.TryParse(m.Groups[3].Value, out we);
                    t1 = m.Groups[6].Value;
                    t2 = m.Groups[7].Value;
                    after = m.Groups[8].Value;

                    int qi = seg.IndexOf("人数", StringComparison.Ordinal);
                    if (qi >= 0)
                    {
                        var nm = Regex.Match(seg.Substring(qi), @"人数\s*[：:]\s*\d+/\d+(.*)$");
                        if (nm.Success) tailName = nm.Groups[1].Value.Trim();
                    }
                }

                var entry = new CourseEntry();
                entry.Name = nextName;
                entry.Code = code;
                entry.Weekday = weekday;
                entry.WeekStart = ws > 0 ? ws : 1;
                entry.WeekEnd = we > 0 ? we : 17;
                entry.StartTime = PadTime(t1);
                entry.EndTime = PadTime(t2);
                entry.Location = ExtractLocation(after);
                entry.Teacher = ExtractTeacher(after, entry.Location);
                list.Add(entry);

                nextName = tailName.Trim();
            }
            return list;
        }

        static string PadTime(string t)
        {
            if (string.IsNullOrEmpty(t)) return "";
            var m = Regex.Match(t, @"^(\d{1,2}):(\d{2})$");
            if (!m.Success) return t;
            return m.Groups[1].Value.PadLeft(2, '0') + ":" + m.Groups[2].Value;
        }

        static string ExtractLocation(string after)
        {
            if (string.IsNullOrEmpty(after)) return "";
            List<string> found = new List<string>();
            var ms = Regex.Matches(after, @"([A-Z]\d+-\d+(?:[^\s]*)?)|(体育场[^\s]*)|([^\s]*教室)");
            foreach (Match m in ms) if (!found.Contains(m.Value.Trim())) found.Add(m.Value.Trim());
            if (found.Count > 0) return string.Join(" / ", found);
            var m2 = Regex.Match(after, @"校区\s*([^\s]+)");
            if (m2.Success) return m2.Groups[1].Value.Trim();
            return "";
        }

        static string ExtractTeacher(string after, string loc)
        {
            if (string.IsNullOrEmpty(after)) return "（未注明）";
            string clean = after;
            if (!string.IsNullOrEmpty(loc))
            {
                foreach (string part in loc.Split(new char[] { '/' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    string p = part.Trim();
                    if (p.Length > 0) clean = clean.Replace(p, " ");
                }
            }
            List<string> names = new List<string>();
            var ms = Regex.Matches(clean, @"([A-Za-z][A-Za-z.,]*(\s[A-Za-z.,]+)*|[\u4e00-\u9fa5]+)\s*\([A-Za-z0-9]+\)");
            foreach (Match m in ms)
            {
                string n = m.Groups[1].Value.Trim();
                if (n.Length > 0 && !names.Contains(n)) names.Add(n);
            }
            if (names.Count == 0) return "（未注明）";
            return string.Join("/", names);
        }
    }

    // ---------------- ICS 生成 ----------------
    public static class IcsWriter
    {
        static string Esc(string s)
        {
            if (s == null) return "";
            s = s.Replace("\\", "\\\\");
            s = s.Replace(";", "\\;");
            s = s.Replace(",", "\\,");
            s = s.Replace("\r", "");
            s = s.Replace("\n", "\\n");
            return s;
        }

        public static string Generate(List<CourseEntry> entries, DateTime firstMonday, int remindMinutes, string calName, bool withTz)
        {
            var sb = new StringBuilder();
            sb.AppendLine("BEGIN:VCALENDAR");
            sb.AppendLine("VERSION:2.0");
            sb.AppendLine("PRODID:-//KejianToIcs//Course Calendar//CN");
            sb.AppendLine("CALSCALE:GREGORIAN");
            sb.AppendLine("METHOD:PUBLISH");
            sb.AppendLine("X-WR-CALNAME:" + Esc(string.IsNullOrEmpty(calName) ? "课表" : calName));
            if (withTz)
            {
                sb.AppendLine("X-WR-TIMEZONE:Asia/Shanghai");
                sb.AppendLine("BEGIN:VTIMEZONE");
                sb.AppendLine("TZID:Asia/Shanghai");
                sb.AppendLine("BEGIN:STANDARD");
                sb.AppendLine("DTSTART:19700101T000000");
                sb.AppendLine("TZOFFSETFROM:+0800");
                sb.AppendLine("TZOFFSETTO:+0800");
                sb.AppendLine("TZNAME:CST");
                sb.AppendLine("END:STANDARD");
                sb.AppendLine("END:VTIMEZONE");
            }
            string dtstamp = DateTime.UtcNow.ToString("yyyyMMddTHHmmss") + "Z";

            for (int i = 0; i < entries.Count; i++)
            {
                CourseEntry e = entries[i];
                if (string.IsNullOrEmpty(e.Name)) continue;

                DateTime day = firstMonday.AddDays((e.WeekStart - 1) * 7 + (e.Weekday - 1));
                string ymd = day.ToString("yyyyMMdd");
                string tt1 = e.StartTime.Replace(":", "") + "00";
                string tt2 = e.EndTime.Replace(":", "") + "00";

                string dstart = withTz
                    ? "DTSTART;TZID=Asia/Shanghai:" + ymd + "T" + tt1
                    : "DTSTART:" + ymd + "T" + tt1;
                string dend = withTz
                    ? "DTEND;TZID=Asia/Shanghai:" + ymd + "T" + tt2
                    : "DTEND:" + ymd + "T" + tt2;

                string desc = "教学班代码：" + e.Code + "\n教师：" + e.Teacher
                            + "\n周次：" + e.WeekStart + "~" + e.WeekEnd + "周";

                sb.AppendLine("BEGIN:VEVENT");
                sb.AppendLine("UID:" + e.Code + "-" + (i + 1) + "@kejiantoics.local");
                sb.AppendLine("DTSTAMP:" + dtstamp);
                sb.AppendLine(dstart);
                sb.AppendLine(dend);
                sb.AppendLine("RRULE:FREQ=WEEKLY;COUNT=" + (e.WeekEnd - e.WeekStart + 1));
                sb.AppendLine("SUMMARY:" + Esc(e.Name));
                sb.AppendLine("LOCATION:" + Esc(e.Location));
                sb.AppendLine("DESCRIPTION:" + Esc(desc));
                if (remindMinutes > 0)
                {
                    sb.AppendLine("BEGIN:VALARM");
                    sb.AppendLine("TRIGGER:-PT" + remindMinutes + "M");
                    sb.AppendLine("ACTION:DISPLAY");
                    sb.AppendLine("DESCRIPTION:课程提醒");
                    sb.AppendLine("END:VALARM");
                }
                sb.AppendLine("END:VEVENT");
            }
            sb.AppendLine("END:VCALENDAR");
            return sb.ToString();
        }

        public static void Save(string path, string content)
        {
            File.WriteAllText(path, content, new UTF8Encoding(true));
        }
    }

    // ---------------- 命令行自检模式 ----------------
    public static class Cli
    {
        public static int Run(string[] args)
        {
            string docx = "", output = "", report = "", calName = "课表", tz = "Asia/Shanghai";
            DateTime firstMonday = new DateTime(2026, 8, 31);
            int min = 20;
            for (int i = 0; i < args.Length; i++)
            {
                string a = args[i];
                string next = (i + 1 < args.Length) ? args[i + 1] : "";
                if (a == "--docx" && next != "") { docx = next; i++; }
                else if (a == "--out" && next != "") { output = next; i++; }
                else if (a == "--report" && next != "") { report = next; i++; }
                else if (a == "--name" && next != "") { calName = next; i++; }
                else if (a == "--tz" && next != "") { tz = next; i++; }
                else if (a == "--start" && next != "") { DateTime.TryParse(next, out firstMonday); i++; }
                else if (a == "--min" && next != "") { int.TryParse(next, out min); i++; }
            }
            if (docx == "" || output == "")
            {
                File.WriteAllText(report == "" ? "cli_usage.txt" : report,
                    "用法: 课表转日历.exe --docx 课表.docx --out 输出.ics [--start 2026-08-31] [--tz Asia/Shanghai|floating] [--min 20] [--name 课表] [--report 报告.txt]");
                return 2;
            }

            var entries = DocxParser.Parse(docx);
            bool withTz = tz != "floating";
            string ics = IcsWriter.Generate(entries, firstMonday, min, calName, withTz);
            IcsWriter.Save(output, ics);

            var sb = new StringBuilder();
            sb.AppendLine("条目数: " + entries.Count);
            for (int i = 0; i < entries.Count; i++)
            {
                var e = entries[i];
                DateTime day = firstMonday.AddDays((e.WeekStart - 1) * 7 + (e.Weekday - 1));
                sb.AppendLine(string.Format("{0,2}. {1,-14} 周{2} {3:yyyy-MM-dd}({4}) {5}-{6} 周次{7}-{8} | {9} | {10} | {11}",
                    i + 1, e.Name, e.Weekday, day, day.DayOfWeek, e.StartTime, e.EndTime,
                    e.WeekStart, e.WeekEnd, e.Location, e.Teacher, e.Code));
            }
            File.WriteAllText(report == "" ? Path.ChangeExtension(output, ".txt") : report, sb.ToString(), new UTF8Encoding(true));
            return 0;
        }
    }

    // ---------------- 主窗体（Windows 原生风格） ----------------
    public class MainForm : Form
    {
        readonly Font UI_FONT = new Font("Segoe UI", 9F);
        readonly Color BG = Color.FromArgb(243, 243, 243);
        readonly Color GRID_LINE = Color.FromArgb(231, 231, 231);
        readonly Color ALT_ROW = Color.FromArgb(249, 249, 249);
        readonly Color HOVER_ROW = Color.FromArgb(245, 245, 245);

        int hoverRow = -1;

        ToolStrip ts = new ToolStrip();
        ToolStripButton btnOpen = new ToolStripButton("打开 .docx");
        ToolStripButton btnParse = new ToolStripButton("解析课表");
        ToolStripButton btnAdd = new ToolStripButton("添加行");
        ToolStripButton btnDel = new ToolStripButton("删除选中行");
        ToolStripButton btnImportIcs = new ToolStripButton("导入 .ics（作业）");
        ToolStripButton btnHelp = new ToolStripButton("教程");
        ToolStripButton btnExport = new ToolStripButton("导出 .ics");

        StatusStrip ss = new StatusStrip();
        ToolStripStatusLabel stFile = new ToolStripStatusLabel();
        ToolStripStatusLabel stSpring = new ToolStripStatusLabel();
        ToolStripStatusLabel stStatus = new ToolStripStatusLabel();

        DataGridView dgv = new DataGridView();
        GroupBox grpSettings = new GroupBox();
        DateTimePicker dtpStart = new DateTimePicker();
        ComboBox cmbTz = new ComboBox();
        NumericUpDown nudMin = new NumericUpDown();
        TextBox txtCalName = new TextBox();

        string lastDocx = "";

        public MainForm()
        {
            Text = "课表转日历 — docx → Outlook .ics";
            Width = 1120;
            Height = 730;
            MinimumSize = new Size(1024, 640);
            StartPosition = FormStartPosition.CenterScreen;
            Font = UI_FONT;
            BackColor = BG;
            KeyPreview = true;

            BuildToolStrip();
            BuildSettings();
            BuildStatusStrip();
            InitGrid();

            Controls.Add(dgv);          // Fill
            Controls.Add(grpSettings);  // Bottom
            Controls.Add(ss);           // Bottom
            Controls.Add(ts);           // Top
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            string flag = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                                       "KejianToIcs", "tutorial.shown");
            if (!File.Exists(flag))
            {
                try
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(flag));
                    File.WriteAllText(flag, "1", new UTF8Encoding(false));
                }
                catch { }
                ShowTutorial();
            }
        }

        void ShowTutorial()
        {
            using (var f = new TutorialForm()) { f.ShowDialog(this); }
        }

        void BuildToolStrip()
        {
            ts.Dock = DockStyle.Top;
            ts.GripStyle = ToolStripGripStyle.Hidden;
            ts.RenderMode = ToolStripRenderMode.Professional;
            ts.Renderer = new Win11Renderer();
            ts.Font = UI_FONT;
            ts.Padding = new Padding(8, 6, 8, 6);
            ts.BackColor = Color.White;

            foreach (var b in new ToolStripButton[] { btnOpen, btnParse, btnAdd, btnDel, btnImportIcs, btnHelp, btnExport })
            {
                b.DisplayStyle = ToolStripItemDisplayStyle.Text;
                b.ImageScaling = ToolStripItemImageScaling.None;
                b.Padding = new Padding(6, 0, 6, 0);
            }
            btnExport.Alignment = ToolStripItemAlignment.Right;

            ts.Items.Add(btnOpen);
            ts.Items.Add(btnParse);
            ts.Items.Add(new ToolStripSeparator());
            ts.Items.Add(btnAdd);
            ts.Items.Add(btnDel);
            ts.Items.Add(new ToolStripSeparator());
            ts.Items.Add(btnImportIcs);
            ts.Items.Add(btnHelp);
            ts.Items.Add(new ToolStripSeparator());
            ts.Items.Add(btnExport);

            btnOpen.Click += (s, e) => OpenFile();
            btnParse.Click += (s, e) => ParseFile();
            btnAdd.Click += (s, e) => dgv.Rows.Add("", "周一", "08:00", "09:35", "1", "17", "", "", "");
            btnDel.Click += (s, e) => DeleteSelectedRows();
            btnImportIcs.Click += (s, e) => ImportIcs();
            btnHelp.Click += (s, e) => ShowTutorial();
            btnExport.Click += (s, e) => ExportIcs();
        }

        void BuildSettings()
        {
            grpSettings.Text = "导出设置";
            grpSettings.Dock = DockStyle.Bottom;
            grpSettings.Height = 64;
            grpSettings.Font = UI_FONT;
            grpSettings.Padding = new Padding(10, 4, 10, 4);

            var table = new TableLayoutPanel();
            table.Dock = DockStyle.Fill;
            table.ColumnCount = 8;
            table.RowCount = 1;
            table.Padding = new Padding(0);
            for (int i = 0; i < 8; i++) table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            table.Controls.Add(MakeLabel("第1周周一", 5), 0, 0);
            dtpStart.Format = DateTimePickerFormat.Custom;
            dtpStart.CustomFormat = "yyyy-MM-dd (dddd)";
            dtpStart.Value = new DateTime(2026, 8, 31);
            dtpStart.Width = 150;
            dtpStart.Margin = new Padding(6, 0, 18, 0);
            dtpStart.Font = UI_FONT;
            table.Controls.Add(dtpStart, 1, 0);

            table.Controls.Add(MakeLabel("时区", 5), 2, 0);
            cmbTz.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbTz.Items.AddRange(new object[] { "Asia/Shanghai (UTC+8)", "不指定时区（浮动）" });
            cmbTz.SelectedIndex = 0;
            cmbTz.Width = 175;
            cmbTz.Margin = new Padding(6, 0, 18, 0);
            cmbTz.Font = UI_FONT;
            table.Controls.Add(cmbTz, 3, 0);

            table.Controls.Add(MakeLabel("提前提醒(分)", 5), 4, 0);
            nudMin.Minimum = 0;
            nudMin.Maximum = 240;
            nudMin.Value = 20;
            nudMin.Width = 65;
            nudMin.Margin = new Padding(6, 0, 18, 0);
            nudMin.Font = UI_FONT;
            table.Controls.Add(nudMin, 5, 0);

            table.Controls.Add(MakeLabel("日历名", 5), 6, 0);
            txtCalName.Text = "课表";
            txtCalName.Width = 150;
            txtCalName.Margin = new Padding(6, 0, 0, 0);
            txtCalName.Font = UI_FONT;
            table.Controls.Add(txtCalName, 7, 0);

            grpSettings.Controls.Add(table);
        }

        Label MakeLabel(string text, int topPad)
        {
            return new Label { Text = text, AutoSize = true, Font = UI_FONT, Anchor = AnchorStyles.Left, Margin = new Padding(0, topPad, 0, 0) };
        }

        void BuildStatusStrip()
        {
            ss.Dock = DockStyle.Bottom;
            ss.SizingGrip = true;
            ss.Font = UI_FONT;
            ss.RenderMode = ToolStripRenderMode.System;

            stFile.Text = "就绪";
            stFile.AutoSize = false;
            stFile.Width = 460;
            stFile.TextAlign = ContentAlignment.MiddleLeft;
            stFile.Spring = true;

            stSpring.Spring = true;

            stStatus.Text = "请打开一个 .docx 课表文件。";
            stStatus.AutoSize = false;
            stStatus.Width = 340;
            stStatus.TextAlign = ContentAlignment.MiddleRight;

            ss.Items.Add(stFile);
            ss.Items.Add(stSpring);
            ss.Items.Add(stStatus);
        }

        void InitGrid()
        {
            dgv.Dock = DockStyle.Fill;
            dgv.BackgroundColor = Color.White;
            dgv.BorderStyle = BorderStyle.None;
            dgv.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            dgv.GridColor = GRID_LINE;
            dgv.EnableHeadersVisualStyles = true;
            dgv.AllowUserToAddRows = true;
            dgv.AllowUserToDeleteRows = true;
            dgv.RowHeadersVisible = false;
            dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgv.MultiSelect = true;
            dgv.Margin = new Padding(8);
            dgv.Font = UI_FONT;

            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9F, FontStyle.Regular);
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(250, 250, 250);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(80, 80, 80);
            dgv.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.FromArgb(250, 250, 250);
            dgv.ColumnHeadersDefaultCellStyle.SelectionForeColor = Color.FromArgb(80, 80, 80);
            dgv.ColumnHeadersHeight = 34;
            dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;

            dgv.DefaultCellStyle.Font = UI_FONT;
            dgv.DefaultCellStyle.BackColor = Color.White;
            dgv.DefaultCellStyle.ForeColor = Color.FromArgb(30, 30, 30);
            dgv.DefaultCellStyle.SelectionBackColor = SystemColors.Highlight;
            dgv.DefaultCellStyle.SelectionForeColor = Color.White;
            dgv.DefaultCellStyle.Padding = new Padding(5, 0, 5, 0);
            dgv.AlternatingRowsDefaultCellStyle.BackColor = ALT_ROW;
            dgv.RowTemplate.Height = 28;

            dgv.AutoGenerateColumns = false;
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            string[] weeks = { "周一", "周二", "周三", "周四", "周五", "周六", "周日" };
            var colWeek = new DataGridViewComboBoxColumn
            {
                HeaderText = "星期",
                Name = "Weekday",
                DataPropertyName = "Weekday",
                FlatStyle = FlatStyle.Flat,
                DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton,
                FillWeight = 64,
                MinimumWidth = 64
            };
            colWeek.Items.AddRange(weeks);
            colWeek.DefaultCellStyle.NullValue = "周一";

            AddTextCol("课程名", "Name", 160, 120);
            dgv.Columns.Add(colWeek);
            AddTextCol("开始", "StartTime", 62, 58);
            AddTextCol("结束", "EndTime", 62, 58);
            AddTextCol("起始周", "WeekStart", 58, 54);
            AddTextCol("结束周", "WeekEnd", 58, 54);
            AddTextCol("地点", "Location", 130, 100);
            AddTextCol("教师", "Teacher", 120, 100);
            AddTextCol("教学班代码", "Code", 135, 110);

            dgv.CellMouseMove += (s, e) =>
            {
                if (e.RowIndex >= 0) SetHoverRow(e.RowIndex);
            };
            dgv.CellMouseLeave += (s, e) => SetHoverRow(-1);
            dgv.RowsRemoved += (s, e) => { if (hoverRow >= dgv.Rows.Count) hoverRow = -1; };
        }

        void SetHoverRow(int idx)
        {
            if (hoverRow == idx) return;
            if (hoverRow >= 0 && hoverRow < dgv.Rows.Count && !dgv.Rows[hoverRow].Selected)
            {
                dgv.Rows[hoverRow].DefaultCellStyle.BackColor = (hoverRow % 2 == 1) ? ALT_ROW : Color.White;
            }
            hoverRow = idx;
            if (idx >= 0 && idx < dgv.Rows.Count && !dgv.Rows[idx].Selected)
            {
                dgv.Rows[idx].DefaultCellStyle.BackColor = HOVER_ROW;
            }
        }

        void AddTextCol(string header, string name, int weight, int minWidth)
        {
            dgv.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = header,
                Name = name,
                DataPropertyName = name,
                FillWeight = weight,
                MinimumWidth = minWidth
            });
        }

        void OpenFile()
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Filter = "Word 文档 (*.docx)|*.docx|所有文件 (*.*)|*.*";
                dlg.Title = "选择课表文档";
                if (dlg.ShowDialog(this) != DialogResult.OK) return;
                lastDocx = dlg.FileName;
                stFile.Text = lastDocx;
                stStatus.Text = "已选择文件，点“解析课表”开始。";
            }
        }

        void ParseFile()
        {
            if (lastDocx == "" || !File.Exists(lastDocx))
            {
                MessageBox.Show(this, "请先选择一个 .docx 文件。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            try
            {
                Cursor = Cursors.WaitCursor;
                var entries = DocxParser.Parse(lastDocx);
                dgv.Rows.Clear();
                foreach (var e in entries)
                {
                    dgv.Rows.Add(
                        e.Name,
                        WeekdayName(e.Weekday),
                        e.StartTime,
                        e.EndTime,
                        e.WeekStart.ToString(),
                        e.WeekEnd.ToString(),
                        e.Location,
                        e.Teacher,
                        e.Code);
                }
                stStatus.Text = "解析完成：" + entries.Count + " 个上课时段，请预览并修正后再导出。";
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "解析失败：\r\n" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        string WeekdayName(int w)
        {
            string[] weeks = { "周一", "周二", "周三", "周四", "周五", "周六", "周日" };
            if (w < 1 || w > 7) return "周一";
            return weeks[w - 1];
        }

        int WeekdayVal(string s)
        {
            string[] weeks = { "周一", "周二", "周三", "周四", "周五", "周六", "周日" };
            for (int i = 0; i < weeks.Length; i++) if (weeks[i] == s) return i + 1;
            int n; if (int.TryParse(s, out n) && n >= 1 && n <= 7) return n;
            return 1;
        }

        void DeleteSelectedRows()
        {
            var rows = dgv.SelectedRows.Cast<DataGridViewRow>().Where(r => !r.IsNewRow).ToList();
            if (rows.Count == 0) return;
            foreach (var r in rows) dgv.Rows.Remove(r);
            stStatus.Text = "已删除 " + rows.Count + " 行。";
        }

        void ExportIcs()
        {
            var entries = new List<CourseEntry>();
            foreach (DataGridViewRow row in dgv.Rows)
            {
                if (row.IsNewRow) continue;
                string name = Cell(row, "Name").Trim();
                if (name == "") continue;
                int ws, we;
                int.TryParse(Cell(row, "WeekStart"), out ws);
                int.TryParse(Cell(row, "WeekEnd"), out we);
                if (ws <= 0) ws = 1;
                if (we <= 0) we = 17;
                if (we < ws) { int tmp = ws; ws = we; we = tmp; }

                entries.Add(new CourseEntry
                {
                    Name = name,
                    Weekday = WeekdayVal(Cell(row, "Weekday")),
                    StartTime = Cell(row, "StartTime").Trim(),
                    EndTime = Cell(row, "EndTime").Trim(),
                    WeekStart = ws,
                    WeekEnd = we,
                    Location = Cell(row, "Location").Trim(),
                    Teacher = Cell(row, "Teacher").Trim(),
                    Code = Cell(row, "Code").Trim()
                });
            }
            if (entries.Count == 0)
            {
                MessageBox.Show(this, "没有可导出的课程（表格为空）。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            DateTime firstMonday = dtpStart.Value.Date;
            bool withTz = cmbTz.SelectedIndex == 0;
            int min = (int)nudMin.Value;
            string calName = txtCalName.Text.Trim();
            string ics = IcsWriter.Generate(entries, firstMonday, min, calName, withTz);

            using (var dlg = new SaveFileDialog())
            {
                dlg.Filter = "iCalendar 文件 (*.ics)|*.ics|所有文件 (*.*)|*.*";
                dlg.Title = "导出 Outlook 日历";
                dlg.FileName = (calName == "" ? "课表" : calName) + ".ics";
                if (dlg.ShowDialog(this) != DialogResult.OK) return;
                IcsWriter.Save(dlg.FileName, ics);
                stStatus.Text = "已导出 " + entries.Count + " 个时段到：" + dlg.FileName;
            }
        }

        string Cell(DataGridViewRow row, string colName)
        {
            var v = row.Cells[colName].Value;
            return v == null ? "" : v.ToString();
        }

        void ImportIcs()
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Filter = "iCalendar 文件 (*.ics)|*.ics|所有文件 (*.*)|*.*";
                if (dlg.ShowDialog(this) != DialogResult.OK) return;
                try
                {
                    int added = 0, skipped = 0;
                    foreach (var ev in IcsReader.Import(dlg.FileName, dtpStart.Value.Date))
                    {
                        if (ev.Ws < 1) { skipped++; continue; }
                        string wd = WeekdayName(ev.Weekday);
                        dgv.Rows.Add(ev.Name, wd, ev.St, ev.Et, ev.Ws.ToString(), ev.Ws.ToString(), "", "", "Canvas");
                        added++;
                    }
                    stStatus.Text = "已导入 " + added + " 个作业事件" +
                        (skipped > 0 ? "，跳过 " + skipped + " 个早于第1周周一的事件（请检查开学日期）" : "") +
                        "，请核对后再导出。";
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "导入失败：" + ex.Message, "导入 .ics", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }
    }

    // ---------------- iCalendar 事件导入 ----------------
    public class IcsEventRow
    {
        public string Name;
        public int Weekday;
        public int Ws;
        public int We;
        public string St;
        public string Et;
    }

    public static class IcsReader
    {
        public static List<IcsEventRow> Import(string path, DateTime firstMon)
        {
            var list = new List<IcsEventRow>();
            string text = File.ReadAllText(path, Encoding.UTF8);
            var re = new Regex(@"BEGIN:VEVENT(.*?)END:VEVENT", RegexOptions.Singleline);
            foreach (Match m in re.Matches(text))
            {
                string block = m.Groups[1].Value;
                string summary = Field(block, "SUMMARY");
                DateTime? start = ParseDt(Field(block, "DTSTART"));
                DateTime? end = ParseDt(Field(block, "DTEND"));
                if (start == null) continue;
                if (end == null) end = start.Value.AddMinutes(30);
                var ev = Map(start.Value, end.Value, firstMon, summary);
                if (ev != null) list.Add(ev);
            }
            return list;
        }

        static string Field(string block, string name)
        {
            var m = Regex.Match(block, @"^(?:" + name + @")(?:;[^\r\n]*)?:(.*)$", RegexOptions.Multiline);
            return m.Success ? m.Groups[1].Value.Trim().TrimEnd('\r') : "";
        }

        static DateTime? ParseDt(string s)
        {
            if (string.IsNullOrEmpty(s)) return null;
            var m = Regex.Match(s, @"^(\d{4})(\d{2})(\d{2})T?(\d{2})?(\d{2})?(\d{2})?(Z)?$");
            if (!m.Success) return null;
            int y = int.Parse(m.Groups[1].Value);
            int mo = int.Parse(m.Groups[2].Value);
            int d = int.Parse(m.Groups[3].Value);
            int h = m.Groups[4].Success ? int.Parse(m.Groups[4].Value) : 0;
            int mi = m.Groups[5].Success ? int.Parse(m.Groups[5].Value) : 0;
            int sec = m.Groups[6].Success ? int.Parse(m.Groups[6].Value) : 0;
            if (m.Groups[7].Success)
            {
                return new DateTime(y, mo, d, h, mi, sec, DateTimeKind.Utc).ToLocalTime();
            }
            return new DateTime(y, mo, d, h, mi, sec);
        }

        static IcsEventRow Map(DateTime start, DateTime end, DateTime firstMon, string summary)
        {
            int diff = (int)(start.Date - firstMon.Date).TotalDays;
            int ws = diff / 7 + 1;
            int wd = ((diff % 7) + 7) % 7 + 1;
            return new IcsEventRow
            {
                Name = string.IsNullOrEmpty(summary) ? "（未命名作业）" : summary,
                Weekday = wd,
                Ws = ws,
                We = ws,
                St = start.ToString("HH:mm"),
                Et = end.ToString("HH:mm")
            };
        }
    }

    // ---------------- 教程窗口 ----------------
    public class TutorialForm : Form
    {
        public TutorialForm()
        {
            Text = "使用教程";
            Width = 700;
            Height = 620;
            MinimumSize = new Size(620, 520);
            StartPosition = FormStartPosition.CenterParent;
            ShowInTaskbar = false;
            Font = new Font("Segoe UI", 9F);
            BackColor = Color.FromArgb(250, 250, 250);

            var rtb = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                BackColor = Color.White,
                Font = new Font("Segoe UI", 10F),
                ScrollBars = RichTextBoxScrollBars.Vertical
            };

            var btn = new Button
            {
                Text = "我知道了",
                Width = 100,
                Height = 32,
                DialogResult = DialogResult.OK,
                FlatStyle = FlatStyle.System,
                Font = new Font("Segoe UI", 9F)
            };
            var bottom = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 56,
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(0, 10, 12, 10)
            };
            bottom.Controls.Add(btn);

            Controls.Add(bottom);
            Controls.Add(rtb);

            BuildContent(rtb);
        }

        void A(RichTextBox rtb, string text, float size, FontStyle st, Color c)
        {
            rtb.SelectionFont = new Font("Segoe UI", size, st);
            rtb.SelectionColor = c;
            rtb.SelectionIndent = 0;
            rtb.SelectionHangingIndent = 0;
            rtb.AppendText(text + "\n");
        }

        void BuildContent(RichTextBox rtb)
        {
            Color dark = Color.FromArgb(30, 30, 30);
            Color dim = Color.FromArgb(68, 68, 68);

            A(rtb, "课表转日历 · 使用教程", 16, FontStyle.Bold, dark);
            A(rtb, "把教务系统导出的 Word 课表（.docx）转换成 Outlook 日历（.ics）。", 10, FontStyle.Regular, dim);
            A(rtb, "整体流程：导出课表(.docx) → 打开 → 解析 → 校对 → 设置 → 导出 .ics → 导入 Outlook。", 10, FontStyle.Regular, dim);
            A(rtb, "", 10, FontStyle.Regular, dim);

            A(rtb, "0. 先从 i西湖 导出课表（得到 .docx）", 11, FontStyle.Bold, dark);
            A(rtb, "打开 i西湖：工作台 → 教务系统 → 我的课表 → 打印 → 导出 → 导出至一个 Word 文件。", 10, FontStyle.Regular, dim);
            A(rtb, "把导出的 .docx 保存并传到电脑上（微信文件传输助手 / 网盘 / 数据线均可）。", 10, FontStyle.Regular, dim);
            A(rtb, "", 10, FontStyle.Regular, dim);

            A(rtb, "1. 打开课表", 11, FontStyle.Bold, dark);
            A(rtb, "点击工具栏「打开 .docx」，选择刚从 i西湖 导出的 Word 课表文件。", 10, FontStyle.Regular, dim);
            A(rtb, "", 10, FontStyle.Regular, dim);

            A(rtb, "2. 解析课表", 11, FontStyle.Bold, dark);
            A(rtb, "点击「解析课表」。程序会自动定位「星期一…星期日」各列，并解析出课程名、星期、上下课时间、起止周次、上课地点、任课教师和教学班代码。", 10, FontStyle.Regular, dim);
            A(rtb, "", 10, FontStyle.Regular, dim);

            A(rtb, "3. 校对与修改", 11, FontStyle.Bold, dark);
            A(rtb, "在表格中逐行检查识别结果。发现错误可直接在单元格里修改；工具栏的「添加行」「删除选中行」用于增删条目，星期列是下拉选择框。", 10, FontStyle.Regular, dim);
            A(rtb, "提示：同一门课每周上多次、或分不同周次（如 8~13 周与 14~16 周两段），应各占一行；地点有多个教室时用「 / 」分隔。", 10, FontStyle.Regular, dim);
            A(rtb, "", 10, FontStyle.Regular, dim);

            A(rtb, "4. 设置导出参数", 11, FontStyle.Bold, dark);
            A(rtb, "「第1周周一」：填学期第一周星期一的日期，程序据此计算每门课首次上课的真实日期。", 10, FontStyle.Regular, dim);
            A(rtb, "「时区」：默认 Asia/Shanghai (UTC+8)；「提前提醒」：课程开始前多少分钟弹提醒（0 表示不提醒）；「日历名」：导入 Outlook 后的日历名称。", 10, FontStyle.Regular, dim);
            A(rtb, "", 10, FontStyle.Regular, dim);

            A(rtb, "5. 导出 .ics", 11, FontStyle.Bold, dark);
            A(rtb, "点击「导出 .ics」选择保存位置。生成的文件带 UTF-8 编码与周重复规则，可直接导入 Outlook。", 10, FontStyle.Regular, dim);
            A(rtb, "", 10, FontStyle.Regular, dim);

            A(rtb, "6. 导入 Outlook（建议导入为独立日历）", 11, FontStyle.Bold, dark);
            A(rtb, "网页版：日历 → 添加日历 → 从文件上传。想要独立日历，可先「添加日历 → 创建空白日历」命名，再上传该 .ics 并选择导入到新建日历。", 10, FontStyle.Regular, dim);
            A(rtb, "手机版：日历 → 左上角头像/菜单 → 设置(齿轮) → 添加日历 → 从文件上传，导入时选择「新建日历」，即成为独立日历。", 10, FontStyle.Regular, dim);
            A(rtb, "桌面版：文件 → 打开和导出 → 导入/导出 → 导入 iCalendar (.ics) 文件，并选择「作为新日历导入」。", 10, FontStyle.Regular, dim);
            A(rtb, "", 10, FontStyle.Regular, dim);

            A(rtb, "7. 从 Canvas 导入作业（可选）", 11, FontStyle.Bold, dark);
            A(rtb, "打开 i西湖：工作台 → Canvas 学习系统 → 日历 → 日历源(Calendar Feed) → 复制 ICS 地址。", 10, FontStyle.Regular, dim);
            A(rtb, "在电脑浏览器打开该地址下载 .ics 文件；回到本工具点「导入 .ics（作业）」选择该文件。", 10, FontStyle.Regular, dim);
            A(rtb, "作业会按截止时间追加到表格，与课表一起导出。请先设好「第1周周一」；早于该日期的作业会被跳过，导入后请核对截止时间。", 10, FontStyle.Regular, dim);
            A(rtb, "", 10, FontStyle.Regular, dim);

            A(rtb, "常见问题", 12, FontStyle.Bold, dark);
            A(rtb, "· 解析结果不对或漏课程？", 10, FontStyle.Bold, dark);
            A(rtb, "  个别课表的排版或合并单元格较特殊，直接在表格中修改相应行即可，无需重新解析。", 10, FontStyle.Regular, dim);
            A(rtb, "· 某门课一周上两次或分周次？", 10, FontStyle.Bold, dark);
            A(rtb, "  拆成两行分别填写，程序会为每行生成独立的周重复事件。", 10, FontStyle.Regular, dim);
            A(rtb, "· 导入后时间错乱？", 10, FontStyle.Bold, dark);
            A(rtb, "  确认「时区」选了 Asia/Shanghai，且「第1周周一」日期填写正确。", 10, FontStyle.Regular, dim);
            A(rtb, "· 如何删除已导入的旧日历？", 10, FontStyle.Bold, dark);
            A(rtb, "  网页版：日历左侧“我的日历”中把鼠标悬停在该日历上 → 点右侧「⋯」→ 移除；若当时并入了默认日历，则进入日历逐条删除旧事件即可。", 10, FontStyle.Regular, dim);
            A(rtb, "  手机版：日历 → 左上角头像/菜单 → 找到该日历 → 点旁边的 ⓘ → 删除日历。", 10, FontStyle.Regular, dim);
            A(rtb, "  桌面版：日历视图左侧右键该日历 → 删除。", 10, FontStyle.Regular, dim);
        }
    }

    // ---------------- Windows 11 风格渲染 ----------------
    sealed class Win11ColorTable : ProfessionalColorTable
    {
        public override Color MenuItemBorder { get { return Color.FromArgb(229, 229, 229); } }
        public override Color ToolStripContentPanelGradientBegin { get { return Color.White; } }
        public override Color ToolStripContentPanelGradientEnd { get { return Color.White; } }
        public override Color ImageMarginGradientBegin { get { return Color.White; } }
        public override Color ImageMarginGradientMiddle { get { return Color.White; } }
        public override Color ImageMarginGradientEnd { get { return Color.White; } }
    }

    sealed class Win11Renderer : ToolStripProfessionalRenderer
    {
        public Win11Renderer() : base(new Win11ColorTable()) { RoundedEdges = false; }

        protected override void OnRenderToolStripBackground(ToolStripRenderEventArgs e)
        {
            using (var b = new SolidBrush(Color.White))
                e.Graphics.FillRectangle(b, e.AffectedBounds);
            using (var p = new Pen(Color.FromArgb(229, 229, 229)))
                e.Graphics.DrawLine(p, 0, e.ToolStrip.Height - 1, e.ToolStrip.Width, e.ToolStrip.Height - 1);
        }

        protected override void OnRenderButtonBackground(ToolStripItemRenderEventArgs e)
        {
            var btn = e.Item as ToolStripButton;
            if (btn == null) { base.OnRenderButtonBackground(e); return; }

            var g = e.Graphics;
            var r = new Rectangle(0, 0, e.Item.Width, e.Item.Height);
            r.Inflate(-3, -4);
            if (r.Width < 4 || r.Height < 4) return;

            bool selected = e.Item.Selected;
            bool pressed = selected && (Control.MouseButtons & MouseButtons.Left) == MouseButtons.Left;

            Color fill;
            if (pressed) fill = Color.FromArgb(227, 227, 227);
            else if (selected) fill = Color.FromArgb(243, 243, 243);
            else return;

            using (var path = RoundedRect(r, 6))
            using (var br = new SolidBrush(fill))
                g.FillPath(br, path);
        }

        protected override void OnRenderSeparator(ToolStripSeparatorRenderEventArgs e)
        {
            if (!e.Vertical) { base.OnRenderSeparator(e); return; }
            int x = e.Item.Width / 2;
            using (var p = new Pen(Color.FromArgb(229, 229, 229)))
                e.Graphics.DrawLine(p, x, 8, x, e.Item.Height - 8);
        }

        internal static GraphicsPath RoundedRect(Rectangle r, int radius)
        {
            var path = new GraphicsPath();
            if (r.Width < 1 || r.Height < 1) return path;
            int d = radius * 2;
            path.AddArc(r.X, r.Y, d, d, 180, 90);
            path.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            path.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            path.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }
    }

    static class Program
    {
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        static extern bool SetProcessDPIAware();

        [STAThread]
        static int Main(string[] args)
        {
            try { if (Environment.OSVersion.Version.Major >= 6) SetProcessDPIAware(); }
            catch { }

            if (args.Length > 0)
            {
                try { return Cli.Run(args); }
                catch (Exception ex)
                {
                    try { File.WriteAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "cli_error.log"), ex.ToString(), new UTF8Encoding(true)); }
                    catch { }
                    return 1;
                }
            }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
            return 0;
        }
    }
}
