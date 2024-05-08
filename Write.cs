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
                    var cellstyle = workbook.CreateCellStyle();
                    IFont font1 = workbook.CreateFont();
                    cell1.SetCellValue(list[i].ProfitandLoss);
                    font1.Color = IndexedColors.Green.Index;
                    cellstyle.SetFont(font1);
                    cell1.CellStyle = cellstyle;
                }
                else
                {
                    ICell cell1 = row.CreateCell(8);
                    var cellstyle = workbook.CreateCellStyle();
                    IFont font1 = workbook.CreateFont();
                    cell1.SetCellValue(list[i].ProfitandLoss);
                    font1.Color = IndexedColors.Red.Index;
                    cellstyle.SetFont(font1);
                    cell1.CellStyle = cellstyle;
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
                ICell cell3 = qrow.CreateCell(8);
                var cellstyle3 = workbook.CreateCellStyle();
                IFont font3 = workbook.CreateFont();
                cell3.SetCellValue(list[count].ProfitandLoss);
                font3.Color = IndexedColors.Green.Index;
                cellstyle3.FillForegroundColor = IndexedColors.Yellow.Index;
                cellstyle3.FillPattern =FillPattern.SolidForeground;
                cellstyle3.SetFont(font3);
                cell3.CellStyle = cellstyle3;
            }
            else
            {
                ICell cell3 = qrow.CreateCell(8);
                var cellstyle3 = workbook.CreateCellStyle();
                IFont font3 = workbook.CreateFont();
                cell3.SetCellValue(list[count].ProfitandLoss);
                font3.Color = IndexedColors.Red.Index;
                cellstyle3.FillForegroundColor = IndexedColors.Yellow.Index;
                cellstyle3.FillPattern = FillPattern.SolidForeground;
                cellstyle3.SetFont(font3);
                cell3.CellStyle = cellstyle3;
            }
            qrow.CreateCell(9).SetCellValue(list[count].HandlingCharge);
            qrow.CreateCell(10).SetCellValue();
            // ICell cell1 = qrow.CreateCell(0);//创建单元格
            //ICell cell2 = row.CreateCell(count);
            //cell1.SetCellValue("Null");//单元格内容

            //var cellstyle1 = workbook.CreateCellStyle();//单元格样式

            //cellstyle1.FillForegroundColor = IndexedColors.BrightGreen.Index;//颜色
            //cellstyle1.FillPattern = FillPattern.SolidForeground;//填充颜色的方式
            ////cell1.CellStyle = cellstyle1;//把样式赋值给单元格
            //qrow.CreateCell(0).CellStyle = cellstyle1;
            //// qrow.RowStyle = cellstyle1;
            //// cell2.SetCellValue("");

            var cellstyle2 = workbook.CreateCellStyle();

            IFont font = workbook.CreateFont();//创建字体
            font.Color = IndexedColors.Red.Index;//字体颜色
            ICell cell = qrow.CreateCell(7);
            cell.SetCellValue(list[count].Tradingvolume);
            ICell cell2 = qrow.CreateCell(9);
            cell2.SetCellValue((list[count].HandlingCharge));
            cellstyle2.SetFont(font);//把字体赋给样式
            cellstyle2.FillForegroundColor = IndexedColors.Yellow.Index;//单元格颜色
            cellstyle2.FillPattern = FillPattern.SolidForeground;	// 填充方式
            cell.CellStyle = cellstyle2;
            cell2.CellStyle = cellstyle2;         
            Console.WriteLine("华丽分割线");
            //文件写入的位置@"C:\Users\Administrator\Desktop\Max.xls"
            using (FileStream fs = File.Create(writpath))
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
