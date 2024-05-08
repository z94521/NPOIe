using NPOIe;
using System.Data;
Path1 pathwrite = new Path1();
Path1 pathread = new Path1();
pathread.Readfilename("20240508");
//pathwrite.Writefilename("9183");
Console.WriteLine(pathread.ReadPath);
DataTable dt = Reead.ExcelToDatatable(pathread.ReadPath, "Sheet1", true);
//@"C:\Users\Administrator\Desktop\20240426.xls"
//将excel表格数据存入list集合中
//EachdayTX定义的类，字段值对应excel表中的每一列
//List<EachdayTX> eachdayTX = new List<EachdayTX>();
//foreach (DataRow dr in dt.Rows)
//{
//    EachdayTX model = new Supply
//    {
//        Sta = dr[0].ToString() + "zh",//excel表中第一列的值，依次类推
//                                      // Date = dr[1].ToString(),
//                                      //TXnum = Convert.ToInt32(dr[2])
//    };
//    eachdayTX.Add(model);
//}
//    List<Supply> data = new List<Supply>();
//    data.Add(new Supply
//    {

//        Value2 = "",
//        Value3 = "111",
//    });
//    //List<Supply> data = new List<Supply>();
//    //假设data 已经存入了数据，根据自己需要添加数据
//    bool a = Write.ListToExcel(eachdayTX);//调用写入excel的方法，写入数据
int FF = 0;
List<Supply> supplies = new List<Supply>();
try
{
    foreach (DataRow data in dt.Rows)
    {
        Supply supply = new Supply
        {
            TradingDay = Convert.ToInt32(data[0]),//交易日
            Time = (string)data[1],//时间
            Breed = (string)data[2],//品种
            Pact = (string)data[3],//合约
            Business = (string)data[4],//买卖
            OpenClose = (string)data[5],//开平
            Closingprice =(string)data[6], //成交价
            Tradingvolume =Convert.ToInt32(data[7]),//成交量
            ProfitandLoss = data[9].ToString(),//盈亏
            HandlingCharge = Convert.ToDouble(data[11]),//手续费
        };
        supplies.Add(supply);
    };
}
catch (Exception e)
{

    Console.WriteLine(e.Message);
}

bool b = Write.ListToExcel(supplies,pathread.ReadPath);
Console.WriteLine("Ok");
//StreamReader streamReader = new StreamReader("");
//string line;
//while ((line = streamReader.ReadLine()) != null)
//{
//    string[] strings = line.Split(",");
//}