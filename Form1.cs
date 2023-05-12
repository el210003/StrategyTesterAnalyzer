using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using HtmlAgilityPack;
using System.Xml;
using System.IO;
using System.Threading;
using System.Globalization;

namespace StrategyTesterAnalyzer
{
    public partial class Form1 : Form
    {
        private Dictionary<int, StrategyTesterAnalyzer.TradeOrders> tradeDict = new Dictionary<int, TradeOrders>();
        private string eaName;
        private string ccy;
        private string tf;
        private string timeString = "";
        private string outputPath = @".\Output";
        private string generatedSet = "";
        private decimal maxDD = 0;
        private decimal relativeDD = 0;
        private decimal absDD = 0;
        private decimal expectedPayoff = 0;
        private decimal profitFactor = 0;


        public Form1()
        {
            InitializeComponent();
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            timer1.Enabled = true;
            timeString = "";
            Thread t = new Thread(ProcessFileThread);
            t.Start();
            
        }


        private void ProcessFileThread()
        { 
            string sourcePath = @".\Reports";
            string customSource = textBox1.Text;
            if(!String.IsNullOrEmpty(customSource))
            {
                sourcePath = customSource;
            }

            if(System.IO.Directory.Exists(outputPath))
            {
                System.IO.DirectoryInfo di = new DirectoryInfo(outputPath);

                foreach (FileInfo file in di.GetFiles())
                {
                    try
                    {
                        file.Delete();
                    }
                    catch(IOException e)
                    {
                        MessageBox.Show("Please close the existing cvs file\nbefore running the analyzer");
                        return;
                    }
                }
            }
            else
            {
                System.IO.Directory.CreateDirectory(outputPath);
            }

            if(System.IO.Directory.Exists(sourcePath))
            {
                timeString = "Processing ...";
                string[] files = System.IO.Directory.GetFiles(sourcePath);
                int fileNumber = files.Length;
                int count = 0;
                int reports = 0;
                foreach(string s in files)
                {
                    if(IsSupportedFormat(System.IO.Path.GetExtension(s)))
                    {
                        decimal status = decimal.Round(((decimal)count /(decimal)fileNumber)*100);
                        timeString = "Processing... (" + status + "%)";
                        string fileName = System.IO.Path.GetFileName(s);
                        addTradeRecords(s);
                        WriteTradeRecordToFile();
                        reports++;
                    }
                    count++;
                    
                }

                string phase = " reports processed";
                if (reports < 2)
                    phase = " report processed";

                timeString = "DONE! " + reports.ToString() + phase;

            }
            else
            {
                MessageBox.Show("Reports folder does not exists");
            }
        }

        private bool IsSupportedFormat(string ext)
        {
            bool rc = false;
            switch (ext.ToLower())
            {
                case ".htm":
                case ".pdf":
                    rc = true;
                    break;
            }

            return rc;
        }

        private void WriteTradeRecordToFile()
        {
            //string sourcePath = @".\Download";
            string targetPath = @".\";
            //string advXmlFileName = eaName + ccy + @".xml";
            string csvFileName = eaName +"_"+ ccy + @".csv";

            if (checkBox1.Checked)
                csvFileName = eaName + @".csv";

            //XmlTextWriter writer = new XmlTextWriter(System.IO.Path.Combine(targetPath, advXmlFileName), null);
            //writer.WriteStartDocument();
            //writer.WriteStartElement("StrategyTesterReport");
            //writer.WriteAttributeString("EA", eaName);
            //writer.WriteAttributeString("CCY", ccy);
            //writer.WriteAttributeString("TimeFrame", tf);
            //writer.WriteStartElement("TradeOrders");
            //foreach(TradeOrders o in tradeDict.Values)
            //{
            //    writer.WriteStartElement("Order");
            //    writer.WriteElementString("ID", o.OrderID.ToString());
            //    writer.WriteElementString("OpenTime", o.StartTime.ToString());
            //    writer.WriteElementString("OpenHour", o.OrderHours.ToString());
            //    writer.WriteElementString("CloseTime", o.EndTime.ToString());
            //    writer.WriteElementString("Type", o.OrderType.ToString("G"));
            //    writer.WriteElementString("LotSize", o.LotSize.ToString());
            //    writer.WriteElementString("OpenPrice", o.openPrice.ToString());
            //    writer.WriteElementString("ClosePrice", o.closePrice.ToString());
            //    writer.WriteElementString("Profit", o.Profit.ToString());
            //    writer.WriteElementString("Profitable", o.Profitable);
            //    writer.WriteElementString("HoldingHours", o.HoldingHours.ToString());
            //    writer.WriteElementString("Balance", o.Balance.ToString());
            //    writer.WriteEndElement();
            //}

            //writer.WriteEndElement();
            //writer.WriteEndDocument();
            //writer.Close();

            string csvFile = System.IO.Path.Combine(outputPath, csvFileName);
            bool csvExist = false;
            if (File.Exists(csvFile))
                csvExist = true;

            using (FileStream fs = File.Open(csvFile, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
            {
                using (StreamWriter log = new StreamWriter(fs))
                {
                    if(!csvExist)
                    {
                        log.WriteLine(string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20},{21}",
                            "EA NAME",
                            "CCY",
                            "SET",
                            "TimeFrame",
                            "OrderID",
                            "StartTime",
                            "WeekNumber",
                            "OrderHours",
                            "EndTime",
                            "OrderType",
                            "LotSize",
                            "OpenPrice",
                            "ClosePrice",
                            "Profit",
                            "Profitable",
                            "HoldingHours",
                            "Balance",
                            "ProfitFactor",
                            "ExpectedPayoff",
                            "AbsDD",
                            "MaxDD",
                            "RelativeDD"));
                    }

                    #region 1st line
                    log.WriteLine(string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20},{21}",
                        eaName,
                        ccy,
                        generatedSet,
                        tf,
                        "",
                        "",
                        "",
                        "",
                        "",
                        "",
                        "",
                        "",
                        "",
                        "",
                        "",
                        "",
                        "",
                        profitFactor,
                        expectedPayoff,
                        absDD,
                        maxDD,
                        relativeDD));
                    #endregion

                    foreach (TradeOrders o in tradeDict.Values)
                    {
                        #region WeekNumber
                        CultureInfo culture = CultureInfo.CurrentCulture;
                        CalendarWeekRule weekRule = culture.DateTimeFormat.CalendarWeekRule;
                        DayOfWeek firstDayOfWeek = culture.DateTimeFormat.FirstDayOfWeek;
                        int weekNumber = culture.Calendar.GetWeekOfYear(o.StartTime, weekRule, firstDayOfWeek);
                        #endregion

                        log.WriteLine(string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20},{21}",
                            eaName,
                            ccy,                            
                            generatedSet,
                            tf,
                            o.OrderID.ToString(),
                            o.StartTime.ToString(),
                            weekNumber,
                            o.OrderHours.ToString(),
                            o.EndTime.ToString(),
                            o.OrderType.ToString("G"),
                            o.LotSize.ToString(),
                            o.openPrice.ToString(),
                            o.closePrice.ToString(),
                            o.Profit.ToString(),
                            o.Profitable,
                            o.HoldingHours.ToString(),
                            o.Balance.ToString(),
                            "",
                            "",
                            "",
                            "",
                            ""));
                        /*
                        log.WriteLine(string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20}",
                            eaName,
                            ccy,
                            generatedSet,
                            tf,
                            o.OrderID.ToString(),
                            o.StartTime.ToString(),
                            o.OrderHours.ToString(),
                            o.EndTime.ToString(),
                            o.OrderType.ToString("G"),
                            o.LotSize.ToString(),
                            o.openPrice.ToString(),
                            o.closePrice.ToString(),
                            o.Profit.ToString(),
                            o.Profitable,
                            o.HoldingHours.ToString(),
                            o.Balance.ToString(),
                            profitFactor,
                            expectedPayoff,
                            absDD,
                            maxDD,
                            relativeDD));*/
                    }
                    log.Close();
                    log.Dispose();
                }
            }
            tradeDict.Clear();
        }

        private void addTradeRecords(string file)
        { 
            var doc = new HtmlAgilityPack.HtmlDocument();
            doc.Load(file);

            var htmlEAName = doc.DocumentNode.SelectSingleNode("//html/head/title");
            eaName = htmlEAName.InnerText.Substring(17);

            var htmlHeader = doc.DocumentNode.SelectSingleNode("//html/body/div/table[1]/tr");
            ccy = htmlHeader.InnerText.Substring(6, 6);

            htmlHeader = doc.DocumentNode.SelectSingleNode("//html/body/div/table[1]/tr[2]");
            string tff = htmlHeader.InnerText;
            int tfIndex = tff.IndexOf("(");
            int tfIndexEnd = tff.IndexOf(")");
            tf = tff.Substring(tfIndex + 1, 3);
            tf = tf.Replace(")", "");

            #region Get the Set name 
            if (chkGroup.Checked)
            {
                generatedSet = "";
                try
                {
                    string fn = System.IO.Path.GetFileName(file);
                    char[] delimiters = { '_' };
                    string[] words = fn.Split(delimiters, StringSplitOptions.RemoveEmptyEntries);

                    bool ccyOK = false;
                    bool tfOK = false;
                    int j = 0;

                    for (int i = 0; i < words.Length - 1; i++)
                    {
                        if (words[i] == ccy)
                            ccyOK = true;

                        if (ccyOK)
                        {
                            if (words[i] == tf)
                            {
                                tfOK = true;
                                j = i+1;
                            }
                        }
                    }

                    if (ccyOK && tfOK)
                    {
                        for (int i = j; i < words.Length - 1; i++)
                        {
                            generatedSet += words[i] + "_";
                        }
                        generatedSet = generatedSet.Substring(0, generatedSet.Length - 1);
                    }

                    //int setIndex = fn.IndexOf("(");
                    //int setIndexEnd = fn.IndexOf(")");
                    //generatedSet = fn.Substring(setIndex + 1, setIndexEnd - setIndex - 1);
                }
                catch (Exception e)
                { }
            }
            #endregion

            htmlHeader = doc.DocumentNode.SelectSingleNode("//html/body/div/table[1]/tr[11]/td[2]");
            decimal.TryParse(htmlHeader.InnerText, out profitFactor);

            htmlHeader = doc.DocumentNode.SelectSingleNode("//html/body/div/table[1]/tr[11]/td[4]");
            decimal.TryParse(htmlHeader.InnerText, out expectedPayoff);

            htmlHeader = doc.DocumentNode.SelectSingleNode("//html/body/div/table[1]/tr[12]/td[2]");
            decimal.TryParse(htmlHeader.InnerText, out absDD);

            htmlHeader = doc.DocumentNode.SelectSingleNode("//html/body/div/table[1]/tr[12]/td[4]");
            string dd = htmlHeader.InnerText;
            string[] arrDD = dd.Split(' ');

            decimal.TryParse(arrDD[0], out maxDD);
            decimal.TryParse(arrDD[1].Substring(1,arrDD[1].Length-3), out relativeDD);
            relativeDD = relativeDD / 100;

            var htmlNodes = doc.DocumentNode.SelectNodes("//html/body/div/table[2]/tr");
            foreach (var node in htmlNodes)
            {
                int nodeCount = node.ChildNodes.Count;
                TradeOrders trade;
                int orderID = 0;

                //string line = "";
                if (nodeCount == 10)
                {
                    if (node.ChildNodes[0].InnerText != "#")
                    {
                        /*
                        line = node.ChildNodes[0].InnerText + "," + node.ChildNodes[1].InnerText + "," + node.ChildNodes[2].InnerText + "," + node.ChildNodes[3].InnerText + "," + node.ChildNodes[4].InnerText + "," +
                        node.ChildNodes[5].InnerText + "," + node.ChildNodes[6].InnerText + "," + node.ChildNodes[7].InnerText + "," + node.ChildNodes[8].InnerText + "," + node.ChildNodes[9].InnerText;
                    */
                    
                        orderID = int.Parse(node.ChildNodes[3].InnerText);
                        trade = tradeDict[orderID];
                        trade.UpdateTradeOrder(node.ChildNodes[1].InnerText,
                                               decimal.Parse(node.ChildNodes[5].InnerText),
                                            decimal.Parse(node.ChildNodes[6].InnerText),
                                            decimal.Parse(node.ChildNodes[7].InnerText),
                                            decimal.Parse(node.ChildNodes[8].InnerText),
                                            decimal.Parse(node.ChildNodes[9].InnerText));
                    }
                }
                else if(nodeCount == 9)
                {
                    /*
                    line = node.ChildNodes[0].InnerText + "," + node.ChildNodes[1].InnerText + "," + node.ChildNodes[2].InnerText + "," + node.ChildNodes[3].InnerText + "," + node.ChildNodes[4].InnerText + "," +
                    node.ChildNodes[5].InnerText + "," + node.ChildNodes[6].InnerText + "," + node.ChildNodes[7].InnerText + "," + node.ChildNodes[8].InnerText + ", ";
                    */
                    orderID = int.Parse(node.ChildNodes[3].InnerText);
                    string orderType = node.ChildNodes[2].InnerText.ToLower();

                    if ((orderType != "modify") && (orderType != "delete") && (orderType != "buy limit") && (orderType != "sell limit"))
                    {
                        trade = new TradeOrders(int.Parse(node.ChildNodes[0].InnerText),
                                                node.ChildNodes[1].InnerText,
                                                orderType,
                                                orderID,
                                                decimal.Parse(node.ChildNodes[4].InnerText),
                                                decimal.Parse(node.ChildNodes[5].InnerText),
                                                decimal.Parse(node.ChildNodes[6].InnerText),
                                                decimal.Parse(node.ChildNodes[7].InnerText)); 

                        tradeDict.Add(orderID, trade);
                    }
                }                
                ///Console.WriteLine(line);
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            label1.Text = timeString;
            if (timeString.Contains("DONE!"))
                timer1.Enabled = false;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();
            folderBrowserDialog1.Description = "Select a BT report folder";
            folderBrowserDialog1.RootFolder = Environment.SpecialFolder.MyComputer;

            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                string selectedFolder = folderBrowserDialog1.SelectedPath;
                // Use the selected folder path as required
                textBox1.Text = selectedFolder;
            }

        }

        private void chkGroup_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
