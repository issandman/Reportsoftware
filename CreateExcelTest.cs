using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Excel = Aspose.Cells;
/*
    1443955916@qq.com
*/
namespace WindowsFormsApplication2
{
    //单例模式
    class CreateExcelTest
    {
        Workbook workBook_excel;
        Worksheet workSheet_excel;
        private static CreateExcelTest createExcel;
        static string today = new GetTime().getDateToday();
        //string excelFilePath = @"C:\Users\14439\Desktop\yingpanhao\报表\"
        //                        + string.Format("Excel_{0}.xlsx", DateTime.Now.ToString("yyyy-MM-dd"));//getTime();
        string excelFilePath = @"C:\Users\han\Desktop\报表\" + string.Format("Excel_{0}.xlsx", DateTime.Now.ToString("yyyy-MM-dd"));//getTime();
        string filepath = @"..\..\modelFile\报表头.xlsx";
        
        private CreateExcelTest()
        {
            //导入破解证书
            try
            {
                Excel.License el = new Excel.License();
                el.SetLicense("Aid/License.lic");
            }
            catch (Exception)
            {
                //...
            }

            //
            workBook_excel = File.Exists(excelFilePath) ? new Workbook(excelFilePath) : new Workbook();
            workSheet_excel = workBook_excel.Worksheets[0];
            workBook_excel.Save(excelFilePath, SaveFormat.Xlsx);

        }

        public static CreateExcelTest GetCreateExcelTest()
        {
            if (createExcel == null)
                return createExcel = new CreateExcelTest();
            return createExcel;
        }
        public Workbook GetWorkBookExcel() {

            return workBook_excel;
        }
        public string ExcelFilePath()
        {
            return excelFilePath;
        }
    }
}
