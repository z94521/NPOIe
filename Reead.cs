using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Data;


#region 读取Excel数据
/// <summary>
/// 将excel中的数据导入到DataTable中
/// </summary>
/// <param name="fileName">文件路径</param>
/// <param name="sheetName">excel工作薄sheet的名称</param>
/// <param name="isFirstRowColumn">第一行是否是DataTable的列名，true是</param>
/// <returns>返回的DataTable</returns>
/// 
public class Reead
{
    public static DataTable ExcelToDatatable(string fileName, string sheetName, bool isFirstRowColumn)
    {
        ISheet sheet = null;
        DataTable data = new DataTable();
        int startRow = 0;
        FileStream fs;
        IWorkbook workbook = null;
        int cellCount = 0;//列数
        int rowCount = 0;//行数
        try
        {
            fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            if (fileName.IndexOf(".xlsx") > 0) // 2007版本
            {
                workbook = new XSSFWorkbook(fs);
            }
            else if (fileName.IndexOf(".xls") > 0) // 2003版本
            {
                workbook = new HSSFWorkbook(fs);
            }
            if (sheetName != null)
            {
                sheet = workbook.GetSheet(sheetName);//根据给定的sheet名称获取数据
            }
            else
            {
                //也可以根据sheet编号来获取数据
                sheet = workbook.GetSheetAt(0);//获取第几个sheet表（此处表示如果没有给定sheet名称，默认是第一个sheet表）  
            }
            if (sheet != null)
            {
                IRow firstRow = sheet.GetRow(0);
                cellCount = firstRow.LastCellNum; //第一行最后一个cell的编号 即总的列数
                
                if (isFirstRowColumn)//如果第一行是标题行
                {
                    for (int i = firstRow.FirstCellNum; i < cellCount; ++i)//第一行列数循环
                    {
                        DataColumn column = new DataColumn(firstRow.GetCell(i).StringCellValue);//获取标题
                        data.Columns.Add(column);//添加列
                    }
                    startRow = sheet.FirstRowNum + 1;//1（即第二行，第一行0从开始）
                }
                else
                {
                    startRow = sheet.FirstRowNum;//0
                }
                //最后一行的标号
                rowCount = sheet.LastRowNum;
                for (int i = startRow; i <= rowCount; ++i)//循环遍历所有行
                {
                    IRow row = sheet.GetRow(i);//第几行
                    if (row == null)
                    {
                        continue; //没有数据的行默认是null;
                    }
                    //将excel表每一行的数据添加到datatable的行中
                    DataRow dataRow = data.NewRow();
                    for (int j = row.FirstCellNum; j < cellCount; ++j)
                    {
                        if (row.GetCell(j) != null) //同理，没有数据的单元格都默认是null
                        {
                            dataRow[j] = row.GetCell(j).ToString();
                        }
                    }
                    data.Rows.Add(dataRow);
                }
            }
            return data;
        }
        catch (Exception ex)
        {
            Console.WriteLine("Exception: " + ex.Message);
            return null;
        }
    }
}
#endregion
