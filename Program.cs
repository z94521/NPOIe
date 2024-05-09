using NPOIe;
using System.Data;
string finame;
Path1 pathwrite = new Path1();
Path1 pathread = new Path1();
do
{
    Console.ForegroundColor = ConsoleColor.Yellow;
    Console.WriteLine("----------------------------------------");
    Console.WriteLine("请输入表格名称");
    finame =Console.ReadLine();
    if (finame.ToUpper() == "Y")
    {
        break;
    }
    pathread.Readfilename(finame);
    pathwrite.Writefilename(finame);
    //Console.WriteLine(pathread.ReadPath);
    DataTable dt = Reead.ExcelToDatatable(pathread.ReadPath, "Sheet1", true);
    //将excel表格数据存入list集合中
    //EachdayTX定义的类，字段值对应excel表中的每一列
    List<Supply> supplies = new List<Supply>();
    try
    {
        foreach (DataRow data in dt.Rows)
        {
            Supply supply = new Supply
            {
                TradingDay = Convert.ToInt32(data[0]),//交易日 //excel表中第一列的值，依次类推
                Time = (string)data[1],//时间
                Breed = (string)data[2],//品种
                Pact = (string)data[3],//合约
                Business = (string)data[4],//买卖
                OpenClose = (string)data[5],//开平
                Closingprice = (string)data[6], //成交价
                Tradingvolume = Convert.ToInt32(data[7]),//成交量
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
    bool b = Write.ListToExcel(supplies, pathwrite.WritePath);//调用写入excel的方法，写入数据 
    Console.ForegroundColor = ConsoleColor.Cyan;
    Console.WriteLine("Ok");
    Console.WriteLine("退出输入Y");

} while (true);
Console.WriteLine("程序结束！");
Console.ReadLine();