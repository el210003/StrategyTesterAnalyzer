using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;

namespace StrategyTesterAnalyzer
{
    public enum OrderTypes
    {
        CLOSE = 0,
        BUY,
        SELL,
        BUY_LIMIT,
        SELL_LIMIT
    }

    public class TradeOrders
    {
        int Index;
        public DateTime StartTime;
        public DateTime EndTime;
        public double HoldingHours;
        public int OrderHours;
        public OrderTypes OrderType;
        public int OrderID;
        public decimal LotSize;
        public decimal openPrice;
        public decimal closePrice;
        public decimal SLValue;
        public decimal TPValue;
        public decimal Profit;
        public decimal Balance;
        public string Profitable;
        public TradeOrders(int index, string orderTime, string orderType, int orderID, decimal lotSize, decimal price, decimal SL, decimal TP)
        {
            this.Index = index;
            DateTime dd = DateTime.ParseExact(orderTime, "yyyy.MM.dd HH:mm", CultureInfo.InvariantCulture);
            if(orderType.ToUpper() == "CLOSE" || orderType.ToUpper() == "T/P" || orderType.ToUpper() == "S/L")
            {
                this.EndTime = dd;
                this.SLValue = SL;
                this.TPValue = TP;
                this.closePrice = price;
            }
            else
            {
                this.StartTime = dd;
                this.OrderHours = dd.Hour;
                this.OrderID = orderID;
                this.LotSize = lotSize;
                this.openPrice = price;

                switch(orderType.ToUpper())
                {
                    case "BUY":
                        this.OrderType = OrderTypes.BUY;
                        break;
                    case "SELL":
                        this.OrderType = OrderTypes.SELL;
                        break;
                    case "BUY LIMIT":
                        this.OrderType = OrderTypes.BUY_LIMIT;
                        break;
                    case "SELL LIMIT":
                        this.OrderType = OrderTypes.SELL_LIMIT;
                        break;
                    default:
                        this.OrderType = OrderTypes.CLOSE;
                        break;

                }

            }
        }


        public void UpdateTradeOrder(string orderTime, decimal price, decimal SL, decimal TP, decimal profit, decimal balance)
        {
            DateTime dd = DateTime.ParseExact(orderTime, "yyyy.MM.dd HH:mm", CultureInfo.InvariantCulture);
            this.EndTime = dd;
            this.closePrice = price;
            this.TPValue = TP;
            this.SLValue = SL;
            this.Profit = profit;
            this.Balance = balance;
            this.Profitable = (profit > 0) ? "TP":"SL";
            double hoildingHrs = dd.Subtract(this.StartTime).TotalHours;

            int decimalPlaces = 1;
            double powerOfTen = Math.Pow(10, decimalPlaces);
            this.HoldingHours = Math.Ceiling(hoildingHrs * powerOfTen) / powerOfTen;

        }
    }
}
