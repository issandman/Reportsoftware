using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace WindowsFormsApplication2
{
    class BaseDateTest
    {
        private static string constr = "server=.;database=UPRESSURE;uid=sa;pwd=sakjdx";
        //模板excel文件
        private Workbook workBook;
        //模板工作sheet
        private Worksheet workSheet;
        //报表excel 
        private Workbook workBook_excel;
        private Worksheet workSheet_excel;
        //总报表excel 单例模式
        string excelFilePath = CreateExcelTest.GetCreateExcelTest().ExcelFilePath();
        SqlConnection sqlConnection = new SqlConnection(constr);
        SqlDataAdapter sqlDataAdapter;
        

        public BaseDateTest() { }
        public void Start(GetTime gt) {

            OpenExcel();
            string sqlquary = string.Format(@"SELECT * FROM 基本数据表 WHERE 日期='{0}'", gt.getDateToday()); // 替换成 new GetTime().getDateToday()
            DataTable datatable = GetDataTable(sqlquary);
            ToExcel(datatable);
            workBook_excel.Save(excelFilePath, SaveFormat.Xlsx);

        }

        private void ToExcel(DataTable datatable)
        {

            DataRow dataRow = datatable.Rows[0];
            Cell cellItem1 = workSheet_excel.Cells["D1"];
            Cell cellItem2 = workSheet_excel.Cells["D2"];
            Cell cellItem3 = workSheet_excel.Cells["D3"];
            cellItem1.PutValue(dataRow["辅运顺槽总进尺"].ToString());
            cellItem2.PutValue(dataRow["胶运顺槽总进尺"].ToString());
            cellItem3.PutValue(dataRow["总进尺平均"]);

            Cell cellItem4 = workSheet_excel.Cells["I1"];
            Cell cellItem5 = workSheet_excel.Cells["I2"];
            Cell cellItem6 = workSheet_excel.Cells["I3"];
            cellItem4.PutValue(dataRow["辅运当日进尺"].ToString());
            cellItem5.PutValue(dataRow["胶运当日进尺"].ToString());
            cellItem6.PutValue(dataRow["当日平均"]);

            Cell cellItem7 = workSheet_excel.Cells["C5"];
            Cell cellItem8 = workSheet_excel.Cells["D5"];
            Cell cellItem9 = workSheet_excel.Cells["F5"];
            Cell cellItem10 = workSheet_excel.Cells["H5"];
            Cell cellItem11 = workSheet_excel.Cells["J4"];
            Cell cellItem12 = workSheet_excel.Cells["B6"];
            cellItem7.PutValue(dataRow["初采时间"].ToString());
            cellItem8.PutValue(dataRow["实测倾斜长度"].ToString());
            cellItem9.PutValue(dataRow["平均采高"]);
            cellItem10.PutValue(dataRow["剩余推进长度"].ToString());
            cellItem11.PutValue(dataRow["工作面涌水量"].ToString());
            cellItem12.PutValue(dataRow["时空关系"]);

        }

        private DataTable GetDataTable(string sqlquary)
        {
            DataTable datatable = new DataTable();
            sqlDataAdapter = new SqlDataAdapter(sqlquary, sqlConnection);
            sqlDataAdapter.Fill(datatable);

            return datatable;
        }

        private void OpenExcel()
        {

            workBook_excel = CreateExcelTest.GetCreateExcelTest().GetWorkBookExcel();

            string filepath = @"..\..\modelFile\基本数据模板.xlsx";
            workBook = new Workbook(filepath);
            workSheet = workBook.Worksheets[0];

            if (workBook_excel.Worksheets["Sheet2"] != null)
            {

                workSheet_excel = workBook_excel.Worksheets["Sheet2"];

            }
            else {

                workBook_excel.Worksheets.Add("Sheet2");
                workSheet_excel = workBook_excel.Worksheets["Sheet2"];

            }

            workSheet_excel.Copy(workSheet);
            //workBook_excel.Save(excelFilePath, SaveFormat.Xlsx);
        }
    }
}
