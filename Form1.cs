using System;
using System.Windows.Forms;

using System.Threading; // Sleep
using System.Net; // IPAddress
using System.Net.NetworkInformation; // Ping
using System.IO; // FileInfo
using System.Collections; // ArrayList
using System.Collections.Generic; // Dictionary

// Excel Writing
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace kkomi
{
    public partial class Form1 : Form
    {
        private const int INTERVAL = 60 * 60 * 1000;
        private int logIndex = 1;
        private Thread thread = null;
        private bool running = false;
        private object lockObj = new object();
        private uint count = 0;
        private System.Windows.Forms.Timer runningTimer;

        private struct CollectData
        {
            public String title;
            public DateTime dateTime;
            public String startIP;
            public String endIP;
            public uint cost;
            public uint runningCount;
            public uint totalCount;
            public uint earn;
        };

        private Dictionary<String, ArrayList> collectData = null;

        public Form1()
        {
            InitializeComponent();
        }

        private void GatheringThread()
        {
            for (int i = 0; i < lvTargets.Items.Count; i++)
            {
                String title = "";
                String startIP = "";
                String endIP = "";
                String cost = "";

                this.Invoke(new Action(() =>
                {
                    title = lvTargets.Items[i].SubItems[1].Text;
                    startIP = lvTargets.Items[i].SubItems[2].Text;
                    endIP = lvTargets.Items[i].SubItems[3].Text;
                    cost = lvTargets.Items[i].SubItems[5].Text;
                }));

                count = 0;
                GetRunningCount(title, startIP, endIP, cost);
            }
        }

        void timer_Tick(object sender, EventArgs e)
        {
            thread = new Thread(new ThreadStart(GatheringThread));
            thread.Start();
        }

        private void BtnStart_Click(object sender, EventArgs e)
        {
            btnStart.Enabled = false;

            collectData = new Dictionary<string, ArrayList>();
            foreach (ListViewItem item in lvTargets.Items)
            {
                collectData[item.SubItems[1].Text] = new ArrayList();
            }
            lvLog.Items.Clear();

            running = true;
            lbStatus.Text = "작 동 중";

            // Run first time
            timer_Tick(null, null);

            runningTimer = new System.Windows.Forms.Timer();
            runningTimer.Interval = INTERVAL;
            runningTimer.Tick += new EventHandler(timer_Tick);
            runningTimer.Start();

            btnStop.Enabled = true;
        }

        private void BtnStop_Click(object sender, EventArgs e)
        {
            btnStop.Enabled = false;

            running = false;
            thread.Join();

            lbStatus.Text = "대 기 중";

            btnStart.Enabled = true;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Add Default Values
            AddTarget("스타덤", "123.141.80.1", "123.141.80.80", "1000");
            AddTarget("엔플러스", "211.245.234.1", "211.245.234.105", "500");
            // WriteExcelData();
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            AddTarget(tbTitle.Text, tbStartIP.Text, tbEndIP.Text, tbCost.Text);
        }

        private void AddTarget(String title, String startIP, String endIP, String cost)
        {
            ListViewItem item = new ListViewItem((lvTargets.Items.Count + 1).ToString());
            item.SubItems.Add(title);
            item.SubItems.Add(startIP);
            item.SubItems.Add(endIP);
            item.SubItems.Add(GetPCCount(startIP, endIP).ToString());
            item.SubItems.Add(cost);
            lvTargets.Items.Add(item);
        }

        private uint ConvertIPStringToUInt(String ipAddrStr)
        {
            IPAddress ipAddr = IPAddress.Parse(ipAddrStr);
            return (uint)((ipAddr.GetAddressBytes()[0] << 24) |
                   (ipAddr.GetAddressBytes()[1] << 16) |
                   (ipAddr.GetAddressBytes()[2] << 8) |
                   (ipAddr.GetAddressBytes()[3] << 0));
        }

        private string ConvertIPUintToString(uint ipAddrUInt)
        {
            return String.Format("{0}.{1}.{2}.{3}",
                (ipAddrUInt >> 24) & 0xff,
                (ipAddrUInt >> 16) & 0xff,
                (ipAddrUInt >> 8) & 0xff,
                (ipAddrUInt >> 0) & 0xff);
        }

        private uint GetPCCount(String start, String end)
        {
            IPAddress startIPAddress;
            IPAddress endIPAddress;
            if (IPAddress.TryParse(start, out startIPAddress) && IPAddress.TryParse(end, out endIPAddress))
            {
                return ConvertIPStringToUInt(end) - ConvertIPStringToUInt(start) + 1;
            }
            else
            {
                return 0;
            }
        }

        private void Gather(String tilte, String ipStr)
        {
            Ping pingSender = new Ping();
            PingReply reply = pingSender.Send(ipStr);

            if (reply.Status == IPStatus.Success)
            {
                lock (lockObj)
                {
                    count++;
                }
            }
        }

        private void GetRunningCount(String title, String start, String end, String cost)
        {
            IPAddress startIPAddress;
            IPAddress endIPAddress;

            if (IPAddress.TryParse(start, out startIPAddress) && IPAddress.TryParse(end, out endIPAddress))
            {
                for (uint i = ConvertIPStringToUInt(start); running && (i <= ConvertIPStringToUInt(end)); i++)
                {
                    String ipStr = ConvertIPUintToString(i);

                    Thread th = new Thread(() =>
                    {
                        Gather(title, ipStr);
                    });
                    th.Start();
                }

                // After 10 seconds, update the result
                Thread.Sleep(10 * 1000);

                this.Invoke(new Action(() =>
                {
                    Result(title, start, end, cost, count);
                }));
            }
        }

        private uint GetCollectCount(String title)
        {
            if (collectData[title].Count >= 24)
            {
                return 24;
            }
            else
            {
                return (uint)collectData[title].Count;
            }
        }

        private uint Get24EarnSum(String title)
        {
            uint sum = 0;
            for (int i = collectData[title].Count - 1; (i >= collectData[title].Count - 24) && (i >= 0); i--)
            {
                sum += ((CollectData)collectData[title][i]).earn;
            }

            return sum;
        }

        private void Result(String title, String start, String end, String cost, uint count)
        {
            String date = String.Format("{0,4:D4}-{1,2:D2}-{2,2:D2}",
                DateTime.Now.Year,
                DateTime.Now.Month,
                DateTime.Now.Day);

            String time = DateTime.Now.Hour.ToString();

            // Add to CollectData
            CollectData cd = new CollectData();
            cd.title = title;
            cd.dateTime = DateTime.Now;
            cd.startIP = start;
            cd.endIP = end;
            cd.cost = uint.Parse(cost);
            cd.runningCount = count;
            cd.totalCount = GetPCCount(start, end);
            cd.earn = count * cd.cost;
            collectData[title].Add(cd);

            // Add to ListVIew
            ListViewItem item = new ListViewItem(logIndex.ToString());
            item.SubItems.Add(date);
            item.SubItems.Add(time);
            item.SubItems.Add(title);
            item.SubItems.Add(count.ToString());
            item.SubItems.Add(String.Format("{0:N2}%", ((float)count / (float)cd.totalCount) * 100.0f));
            item.SubItems.Add(cd.earn.ToString());
            item.SubItems.Add(Get24EarnSum(title).ToString());

            lvLog.Columns[7].Text = "" + GetCollectCount(title) + "시간 누적 매출";

            lvLog.Items.Add(item);

            logIndex++;
        }

        private void WriteExcelData()
        {
            Excel.Application excelApp = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            try
            {
                excelApp = new Excel.Application();

                String desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                String excelPath = String.Format("{0}\\{1,4:D4}-{2,2:D2}-{3,2:D2}.xlsx",
                    desktopPath, DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);

                FileInfo fi = new FileInfo(excelPath);
                if (!fi.Exists)
                {
                    wb = excelApp.Workbooks.Add();
                }
                else
                {
                    wb = excelApp.Workbooks.Open(excelPath);
                }

                ws = wb.Worksheets.get_Item(1) as Excel.Worksheet;

                ws.Cells[2, 3] = 1;
                ws.Cells[2, 4] = 2;
                ws.Cells[2, 5] = "=C2+D2";

                Console.WriteLine(excelPath + ", " + wb.Path);
                excelApp.DisplayAlerts = false;
                wb.SaveAs(excelPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                wb.Close();
                excelApp.Quit();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                ReleaseExcelObject(ws);
                ReleaseExcelObject(wb);
                ReleaseExcelObject(excelApp);
            }
        }

        private void ReleaseExcelObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    Marshal.ReleaseComObject(obj);
                    obj = null;
                }
            }
            catch (Exception ex)
            {
                obj = null;
                throw ex;
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
