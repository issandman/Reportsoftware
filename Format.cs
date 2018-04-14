using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Cells;
using System.Drawing;

namespace WindowsFormsApplication2
{
    class Format
    {
        public Style Titlestyle(Workbook workbook)
        {
            //为标题设置样式     
            Style styleTitle = workbook.Styles[workbook.Styles.Add()];//新增样式
            styleTitle.HorizontalAlignment = TextAlignmentType.Center;//文字居中
            styleTitle.Font.Name = "宋体";//文字字体
            styleTitle.Font.Size = 11;//文字大小
            styleTitle.VerticalAlignment = TextAlignmentType.Center;//垂直居中
            styleTitle.IsTextWrapped = true;//单元格内容自动换行
            styleTitle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;//设置上边框
            styleTitle.Borders[BorderType.TopBorder].Color = Color.Black;//颜色
            styleTitle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            styleTitle.Borders[BorderType.BottomBorder].Color = Color.Black;
            styleTitle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            styleTitle.Borders[BorderType.LeftBorder].Color = Color.Black;
            styleTitle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
            styleTitle.Borders[BorderType.RightBorder].Color = Color.Black;
            return styleTitle;
        }
    }
}
