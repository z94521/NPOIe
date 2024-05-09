namespace NPOIe
{
    public class Supply 
    {
        public int TradingDay { get; set; }//交易日
        public string Time { get; set; }//时间
        public string Breed { get; set; }//品种
        public string Pact { get; set; }//合约
        public string Business { get; set; }//买卖

        public string OpenClose { get; set; }//开平
        public string Closingprice { get; set; }//成交价

        public int Tradingvolume { get; set; }//成交量

        public string? ProfitandLoss { get; set; }//盈亏
        public double HandlingCharge { get; set; }//手续费
        public double Netprofit { get; set; }//净利润

    }

}
