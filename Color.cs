using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIe
{
    internal class Color
    {
        public Color(IWorkbook workbook) { 
            this.workbook = workbook;
        }
        IWorkbook workbook;
        public ICellStyle ColorRed()
        {
            var cellstyle = workbook.CreateCellStyle();
            IFont font = workbook.CreateFont();          
            font.Color = IndexedColors.Red.Index;
            cellstyle.SetFont(font);
            return cellstyle;
        }//字体颜色
        public ICellStyle ColorGreen()
        {
            var cellstyle = workbook.CreateCellStyle();
            IFont font = workbook.CreateFont();
            font.Color = IndexedColors.Green.Index;
            cellstyle.SetFont(font);
            return cellstyle;
        }
        public ICellStyle RedYellow()
        {
            var cellstyle = workbook.CreateCellStyle();
            cellstyle.FillForegroundColor = IndexedColors.Yellow.Index;//单元格颜色
            cellstyle.FillPattern = FillPattern.SolidForeground;//单元格填充方式
            IFont font = workbook.CreateFont();//创建字体
            font.Color = IndexedColors.Red.Index;//字体颜色
            cellstyle.SetFont(font);//将字体赋给样式
            return cellstyle;
        }
        public ICellStyle GreenYellow()
        {
            var cellstyle = workbook.CreateCellStyle();
            cellstyle.FillForegroundColor = IndexedColors.Yellow.Index;//单元格颜色
            cellstyle.FillPattern = FillPattern.SolidForeground;//单元格填充方式
            IFont font = workbook.CreateFont();//创建字体
            font.Color = IndexedColors.Green.Index;//字体颜色
            cellstyle.SetFont(font);//将字体赋给样式
            return cellstyle;
        }

    }
}
