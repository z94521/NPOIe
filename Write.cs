using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOIe;


//此处是将list集合写入excel表，Supply也是自己定义的类，每一个字段对应需要写入excel表的每一列的数据
//一次最多能写65535行数据，超过需将list集合拆分，分多次写入
#region 写入excel
class Write
{
    public static bool ListToExcel(List<Supply> list,string writpath)
    {
        
        bool result = false;
        IWorkbook workbook = new HSSFWorkbook();
        Color color = new Color(workbook);
        ISheet sheet = workbook.CreateSheet("Sheet1");//创建一个名称为Sheet0的表;
        IRow row = sheet.CreateRow(0);//（第一行写标题)

        row.CreateCell(0).SetCellValue("交易日");//第一列标题，以此类推
        row.CreateCell(1).SetCellValue("时间");
        row.CreateCell(2).SetCellValue("品种");
        row.CreateCell(3).SetCellValue("合约");
        row.CreateCell(4).SetCellValue("买卖");
        row.CreateCell(5).SetCellValue("开平");
        row.CreateCell(6).SetCellValue("成交价");
        row.CreateCell(7).SetCellValue("成交量");
        row.CreateCell(8).SetCellValue("平仓盈亏(逐笔)");
        row.CreateCell(9).SetCellValue("手续费");
        row.CreateCell(10).SetCellValue("净利润");
        
        int count = list.Count;

        int max = 65535;//最大行数限制
        if (count < max)
        {
            //每一行依次写入
            for (int i = 0; i < list.Count; i++)
            {
                row = sheet.CreateRow(i + 1);//i+1:从第二行开始写入(第一行可同理写标题)，i从第一行写入
                row.CreateCell(0).SetCellValue(list[i].TradingDay);//第一列的值
                row.CreateCell(1).SetCellValue(list[i].Time);//第二列的值
                row.CreateCell(2).SetCellValue(list[i].Breed);
                row.CreateCell(3).SetCellValue(list[i].Pact);
                row.CreateCell(4).SetCellValue(list[i].Business);
                row.CreateCell(5).SetCellValue(list[i].OpenClose);
                row.CreateCell(6).SetCellValue(list[i].Closingprice);
                row.CreateCell(7).SetCellValue(list[i].Tradingvolume);
                if (list[i].ProfitandLoss.StartsWith('-'))
                {
                    ICell cell1 = row.CreateCell(8);
                    cell1.SetCellValue(list[i].ProfitandLoss);
                    cell1.CellStyle = color.ColorGreen();
                }
                else
                {
                    ICell cell1 = row.CreateCell(8);
                    cell1.SetCellValue(list[i].ProfitandLoss);
                    cell1.CellStyle = color.ColorRed();
                }

                row.CreateCell(9).SetCellValue(list[i].HandlingCharge);

            }
            IRow qrow = sheet.CreateRow(count);
            qrow = sheet.CreateRow(count);

            qrow.CreateCell(0).SetCellValue(list[--count].TradingDay);
            qrow.CreateCell(1).SetCellValue(list[count].Time);
            qrow.CreateCell(2).SetCellValue(list[count].Breed);
            qrow.CreateCell(3).SetCellValue(list[count].Pact);
            qrow.CreateCell(4).SetCellValue(list[count].Business);
            qrow.CreateCell(5).SetCellValue(list[count].OpenClose);
            qrow.CreateCell(6).SetCellValue(list[count].Closingprice);
            qrow.CreateCell(7).SetCellValue(list[count].Tradingvolume);
            //qrow.CreateCell(8).SetCellValue(list[count].ProfitandLoss);
            if (list[count].ProfitandLoss.StartsWith('-'))
            {
                ICell cell3 = qrow.CreateCell(8);;
                cell3.SetCellValue(Convert.ToDouble(list[count].ProfitandLoss.ToString()));
                cell3.CellStyle = color.GreenYellow();
            }
            else
            {
                ICell cell3 = qrow.CreateCell(8);
              
                cell3.SetCellValue(Convert.ToDouble(list[count].ProfitandLoss.ToString()));
               
                cell3.CellStyle = color.RedYellow();
            }

            qrow.CreateCell(9).SetCellValue(list[count].HandlingCharge.ToString());
                //净利润的计算
            double b = Convert.ToDouble(qrow.GetCell(8).ToString());//盈亏
            double c = Convert.ToDouble(qrow.GetCell(9).ToString());//手续费
            //Console.WriteLine($"{b}  and {c}");
            list[count].Netprofit = b - c;
            if (list[count].Netprofit > 0)
            {
                ICell cell5 = qrow.CreateCell(10);
                cell5.SetCellValue((double)list[count].Netprofit);
               cell5.CellStyle = color.RedYellow();
                

            }
            else 
            {
                ICell cell5 = qrow.CreateCell(10);
                cell5.SetCellValue((double)list[count].Netprofit);
                cell5.CellStyle = color.GreenYellow();
                
            }
           
           //成交量和手续费的字体颜色及单元格背景 
            ICell cell = qrow.CreateCell(7);
            cell.SetCellValue(list[count].Tradingvolume);
            ICell cell2 = qrow.CreateCell(9);
            cell2.SetCellValue((list[count].HandlingCharge));
            cell.CellStyle = color.RedYellow();
            cell2.CellStyle = color.RedYellow();         
            //Console.WriteLine("================华丽分割线================");
            //文件写入的位置@"C:\Users\Administrator\Desktop\Max.xls"
            using (FileStream fs = File.OpenWrite(writpath))
            {
                workbook.Write(fs);//向打开的这个xls文件中写入数据  
                result = true;
            }
        }
        else
        {
            Console.WriteLine("超过行数限制！");
            result = false;
        }

        return result;

    }

}
#endregion
