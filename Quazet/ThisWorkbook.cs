using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

using Microsoft.Azure.ServiceBus;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;

using Microsoft.WindowsAzure.Storage.Table;
using Microsoft.WindowsAzure.Storage;
using System.Media;
using Newtonsoft.Json;
using Microsoft.Office.Interop.Excel;

namespace Quazet
{
    public class Position
    {
        public string Symbol { get; set; }
        public int Quantity { get; set; }
        public OpenRange Open { get; set; }
        public List<Alert> Alerts = new List<Alert>();
        public int row;
    }

    public class OpenRange
    {
        public Double Low;
        public Double High;
        public int Quantity;
    }

    public class Alert
    {
        public string Symbol;
        public string Trade;
        public string Status;
        public double Limit;
        public double Stop;
        public bool flag;
    }

    public class OpeningRangesTable : TableEntity
    {
        public string Symbol { get; set; }
        public int Quantity { get; set; }
        public DateTime Date { get; set; }
        public double Low { get; set; }
        public double High { set; get; }
        public string Status { set; get; }
        public string Trade { set; get; }
        public double Limit { set; get; }
        public double Stop { set; get; }

        public OpeningRangesTable(string pKey, string rKey)
        {
            PartitionKey = pKey;
            RowKey = rKey;
        }

        public OpeningRangesTable() { }
    }
    public partial class ThisWorkbook
    {
        private List<Position> Positions = new List<Position>();
#if DEBUG
        public static string topic = "trade-alerts2";
#else
        public static string topic = "trade-alerts";
#endif
        private void ThisWorkbook_Startup(object sender, System.EventArgs e)
        {
            var serviceBusConnectionString = ConfigurationManager.AppSettings["serviceBus"];
            var queueClient = new QueueClient(serviceBusConnectionString, topic); //, ReceiveMode.PeekLock);
            Debug.WriteLine(topic);

            var player = new System.Media.SoundPlayer();
            try
            {
                player.SoundLocation = System.IO.Path.GetDirectoryName(Application.ActiveWorkbook.FullName)+"\\Resources\\AllABoard2.wav";
                player.Play();
            }
            catch (Exception ex)
            {
                Debug.WriteLine("No Sounds");
                Debug.WriteLine(System.IO.Path.GetDirectoryName(Application.ActiveWorkbook.FullName));
                Debug.WriteLine(ex.Message);
                Console.Beep(); 
            }

            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Application.ActiveSheet);
            Excel.Range firstRow = activeWorksheet.get_Range("A1");
            firstRow.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
            Excel.Range newFirstRow = activeWorksheet.get_Range("A1");
            //'newFirstRow.Value2 = "This text was added by using code";

            var TradeAlertsConnectionString = ConfigurationManager.AppSettings["TradeAlertsStorage"];

            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(TradeAlertsConnectionString);
            CloudTableClient tableClient = storageAccount.CreateCloudTableClient();
            CloudTable table = tableClient.GetTableReference("OpeningRanges");

            TableQuery<OpeningRangesTable> query = new TableQuery<OpeningRangesTable>()
                .Where(
                     TableQuery.GenerateFilterConditionForDate("Date", QueryComparisons.GreaterThanOrEqual, DateTime.Today)
                )
                .Select(
                    new string[] { "Symbol", "Quantity", "Date", "Low", "High" }
                );

            //await table.ExecuteQuerySegmentedAsync<OpenRange>(query, null);
            //var list = table.ExecuteQuery(query);
            List<OpeningRangesTable> list = table.CreateQuery<OpeningRangesTable>()
                .Where(x => x.Date > DateTime.Today).ToList();
                //.Where(x => x.Date > DateTime.Parse("04/03/2019")).ToList();

            int pos = 0;
            list = list.OrderBy(x => (x.Symbol, x.Date)).ToList();
            foreach (OpeningRangesTable O in list)
            {
                //Positions[O.Symbol]
                if (O.High > 0)
                    Positions.Add(new Position { Symbol = O.Symbol, Quantity = O.Quantity, Open = new OpenRange { Low = O.Low, High = O.High, Quantity = O.Quantity} });
                else
                {
                    pos = Positions.FindIndex(x => x.Symbol == O.Symbol);
                    try
                    {
                        Positions[pos].Alerts.Add(new Alert { Status = O.Status, Trade = O.Trade, Limit = O.Limit, Stop = O.Stop, flag = false });
                    }
                    catch { }
                }
            }

            activeWorksheet.Cells[1, 1] = "Ticker Symbol";
            activeWorksheet.Cells[1, 2] = "Opening Low";
            activeWorksheet.Cells[1, 3] = "Opening High";
            activeWorksheet.Cells[1, 4] = "Opening Spread $";
            activeWorksheet.Cells[1, 5] = "Opening Spread %";
            activeWorksheet.Cells[1, 6] = "Opening Mean";
            activeWorksheet.Cells[1, 7] = "Trade Quantity";
            //activeWorksheet.Cells[2, 7].EntireColumn.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            activeWorksheet.Cells[1, 1].EntireRow.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            activeWorksheet.Range[activeWorksheet.Cells[1, 7], activeWorksheet.Cells[99, 7]].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            activeWorksheet.Cells[1, 1].EntireRow.WrapText = true;
            activeWorksheet.Cells[1, 1].EntireRow.Style.Font.Bold = true;
            activeWorksheet.Cells[1, 1].EntireRow.Interior.Color = System.Drawing.Color.GhostWhite;

            int row = 2;
            foreach (Position P in Positions)
            {
                //Console.Write($"{P.Symbol.PadRight(5)} {P.Open.Low.ToString("0.00").PadRight(6)} - {P.Open.High.ToString("0.00").PadRight(6)} ");
                //Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Application.ActiveSheet);
                activeWorksheet.Cells[row,1].Value = P.Symbol;
                activeWorksheet.Cells[row, 2].Value = Math.Round(P.Open.Low, 2);
                activeWorksheet.Cells[row,3].Value = Math.Round(P.Open.High,2);
                activeWorksheet.Cells[row, 4].Value = Math.Round(P.Open.High - P.Open.Low,2);
                activeWorksheet.Cells[row, 5].Value = Math.Round(((P.Open.High - P.Open.Low) / P.Open.High)*100,2);
                activeWorksheet.Cells[row, 6].Value = Math.Round((P.Open.High + P.Open.Low)/2,2);
                activeWorksheet.Cells[row, 7].Value = P.Open.Quantity;
                P.row = row++;

                foreach (Alert A in P.Alerts)
                    PushAlert( P.Symbol, A);
            }

            queueClient.RegisterMessageHandler(async (msg, exception) =>
            {
                var body = Encoding.UTF8.GetString(msg.Body);
                Debug.WriteLine(body);
                try
                {
                    if (body.IndexOf("A-Up") > 0)
                        player.SoundLocation = System.IO.Path.GetDirectoryName(Application.ActiveWorkbook.FullName) + "\\Resources\\BuyBuyBuy.wav";
                    else
                        if (body.IndexOf("A-D") > 0)
                        player.SoundLocation = System.IO.Path.GetDirectoryName(Application.ActiveWorkbook.FullName) + "\\Resources\\SellSellSell.wav";
                    else
                        player.SoundLocation = System.IO.Path.GetDirectoryName(Application.ActiveWorkbook.FullName) + "\\Resources\\BabyCry.wav";
                    player.Play();
                }
                catch
                {
                    Debug.WriteLine("No Sounds");
                    Console.Beep();
                }
                Alert A = JsonConvert.DeserializeObject<Alert>(body);
                PushAlert(A.Symbol, A);

                await Task.CompletedTask;
            },
            async exception =>
            {
                await Task.CompletedTask;
                Debug.WriteLine(exception);
                // log exception
            }
            );
            Debug.WriteLine("Finished Queue.");

            void PushAlert(string symbol, Alert a) {
                int Apos = Positions.FindIndex(x => x.Symbol == symbol);
                int Arow = Positions[Apos].row;
                //activeWorksheet.Cells.Font.Color = System.Drawing.Color.Black;
                activeWorksheet.Range[activeWorksheet.Cells[2, 1], activeWorksheet.Cells[50, 20]].Interior.ColorIndex = 0; // System.Drawing.Color.Transparent;
                for (int Acol = 4; Acol < 99; Acol += 4)
                {
                    if (activeWorksheet.Cells[Arow, Acol].Value == null)
                    {
                        activeWorksheet.Cells[Arow, Acol+0].value = a.Status;
                        activeWorksheet.Cells[Arow, Acol+1].value = a.Trade;
                        activeWorksheet.Cells[Arow, Acol + 2].value = Math.Round(a.Limit, 2);
                        activeWorksheet.Cells[Arow, Acol+3].value = Math.Round(a.Stop,2);
                        activeWorksheet.Range[activeWorksheet.Cells[Arow, Acol],activeWorksheet.Cells[Arow, Acol + 3]].Interior.Color = System.Drawing.Color.YellowGreen;
                        Microsoft.Office.Tools.Excel.Controls.Button button =
                            new Microsoft.Office.Tools.Excel.Controls.Button();
                        activeWorksheet.Cells[1, Acol + 0].value = "Alert Status";
                        activeWorksheet.Cells[1, Acol + 1].value = "Trade Type";
                        activeWorksheet.Cells[1, Acol + 2].value = "Limit Order";
                        activeWorksheet.Cells[1, Acol + 3].value = "Stop  Loss";
                        Acol = 100;
                    }
                }
            }
        }

        private void ThisWorkbook_Shutdown(object sender, System.EventArgs e)
        {
        }

#region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisWorkbook_Startup);
            this.Shutdown += new System.EventHandler(ThisWorkbook_Shutdown);
        }

#endregion

    }
}
