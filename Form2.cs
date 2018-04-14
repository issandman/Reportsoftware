using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Aspose.Cells;
using System.IO;

namespace WindowsFormsApplication2
{
    public partial class Form2 : Form
    {
        BaseDateTest bdt = new BaseDateTest();
        GetTime gt = new GetTime();
        string working_face = "2101";    //工作面
        //int work_face = int.Parse(working_face);
        private double[,] data = new double[21, 16];
        public List<ComboBox> list_ComboBox = new List<ComboBox>();   //不同深度煤粉量链表
        double[] data_test = new double[10];

        int count = 1;   //钻孔计数器
        int count1 = 0;   //单孔深度计数器
        int tag = 0;    //标志是否出现卡钻等异常情况

        string year = DateTime.Now.Year.ToString();     //日期
        string month = DateTime.Now.Month.ToString();
        string day = DateTime.Now.Day.ToString();
        string year1 = DateTime.Now.AddDays(-1).Year.ToString();     //前一天日期
        string month1 = DateTime.Now.AddDays(-1).Month.ToString();
        string day1 = DateTime.Now.AddDays(-1).Day.ToString();

        double auxiliary;   //辅运顺槽进尺
        double rubber;      //胶运顺槽进尺

        public Form2()
        {
            InitializeComponent();
            loadList();
        }

        private void loadList()
        {
            list_ComboBox.Add(comboBox2);
            list_ComboBox.Add(comboBox3);
            list_ComboBox.Add(comboBox4);
            list_ComboBox.Add(comboBox5);
            list_ComboBox.Add(comboBox6);
            list_ComboBox.Add(comboBox7);
            list_ComboBox.Add(comboBox8);
            list_ComboBox.Add(comboBox9);
            list_ComboBox.Add(comboBox10);
            list_ComboBox.Add(comboBox11);
            list_ComboBox.Add(comboBox12);
            list_ComboBox.Add(comboBox13);
            list_ComboBox.Add(comboBox14);
            list_ComboBox.Add(comboBox15);
            list_ComboBox.Add(comboBox16);
        }

        private void loaddata()     //存储钻屑法数据
        {
            int a = 0;  //标识符（保证数据有问题时可修改不被删除）
            double max = 0;  //单孔最大值
            double max_depth = 0;   //最大值对应深度
            double add = 0;   //累加

            //将煤粉量读入数组
            for (int i = 0; i < list_ComboBox.Count(); i++)
            {
                string temp = list_ComboBox[i].Text.ToString();
                if (String.IsNullOrWhiteSpace(temp))
                {
                    continue;
                }
                try
                {
                    if (temp == "吸钻" || temp == "卡钻" || temp == "煤炮" || temp == "卡钻吸钻")
                    {
                        switch (temp)
                        {
                            case "吸钻":
                                tag = -1; break;
                            case "卡钻":
                                tag = -2; break;
                            case "煤炮":
                                tag = -3; break;
                            case "卡钻吸钻":
                                tag = -4; break;

                        }
                        data[i, count - 1] = tag;
                    }
                    else
                    {
                        tag = 0;
                        data[i, count - 1] = double.Parse(temp);

                        add += double.Parse(temp);
                        if (max < double.Parse(temp))
                        {
                            max = double.Parse(temp);
                            max_depth = i + 1;
                        }
                        count1++; 
                    }
                }
                catch (Exception ex)
                {
                    string tips = "数据的五种格式为：“煤粉量数字”、“吸钻”、“卡钻”、“卡钻吸钻”“煤炮”";
                    tips = ex.Message + tips;
                    MessageBox.Show(tips);
                    count--;
                    a = 1;
                }
            }
            //将钻孔距工作面距离、最大值、对应深度、平均值读入数组，清空内容
            if (a == 0)
            {
                try
                {
                    data[15, count - 1] = double.Parse(comboBox17.Text.ToString());
                    data[16, count - 1] = max;
                    data[17, count - 1] = max_depth;
                    data[18, count - 1] = add / count1;
                    if(comboBox18.Text.ToString() == "辅运顺槽")
                        data[19, count - 1] = 0;
                    else if (comboBox18.Text.ToString() == "胶运顺槽")
                        data[19, count - 1] = 1;

                    comboBox17.Text = null;     //清空填空内容
                    foreach (ComboBox combobox in list_ComboBox)
                    {
                        combobox.Text = null;
                    }
                }
                catch
                {
                    MessageBox.Show("未设置该钻孔距工作面距离");
                    count--;
                }
            }
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            string date = dateTimePicker1.Value.ToShortDateString(); //取年月日
            year = dateTimePicker1.Value.Year.ToString();
            month = dateTimePicker1.Value.Month.ToString();
            day = dateTimePicker1.Value.Day.ToString();

            year1 = dateTimePicker1.Value.AddDays(-1).Year.ToString();
            month1 = dateTimePicker1.Value.AddDays(-1).Month.ToString();
            day1 = dateTimePicker1.Value.AddDays(-1).Day.ToString();
            //MessageBox.Show(date);
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            working_face = comboBox1.Text.ToString();
            //MessageBox.Show(working_face);
        }

        private void button9_Click(object sender, EventArgs e)  //钻屑下一个按钮
        {
            loaddata();
            count++;
            label38.Text = count.ToString() + "号钻孔";
        }

        private void button10_Click(object sender, EventArgs e)  //钻屑上一个按钮
        {
            count--;
            label38.Text = count.ToString() + "号钻孔";
        }

        private void button4_Click(object sender, EventArgs e)  //钻屑法保存预览按钮
        {
            Format format = new Format();
            string filename = string.Format("{0}工作面钻屑监测统计表{1}年{2}月.xlsx", working_face, year, month);
            string headline = string.Format("{0}月{1}日轨顺槽钻屑监测统计表", month, day);
            string headline1 = string.Format("{0}月{1}日运顺槽钻屑监测统计表", month, day);
            string path = @"F:\" + filename;

            string test = @"F:\钻屑监测统计表模板_test.xlsx";

            loaddata();
            if (data[0, count - 1] != 0)
                count++;

            //生成添加钻屑监测统计表
            if (!File.Exists(path))
            {
                data_test[0] = 1.6;
                data_test[1] = 3.2;
                string a = data_test[3].ToString();
                MessageBox.Show(a);
            }
            else
            {
                Workbook workbook1 = new Workbook(test);
                Worksheet DetailSheet = workbook1.Worksheets[0];
                Cells cells = DetailSheet.Cells;

                cells.Merge(0, 2, 2, count-1);
                cells[0, 2].PutValue(headline);//填写标题
                cells[0, 2].SetStyle(format.Titlestyle(workbook1));

                for (int i = 1; i < count; i++)//填写钻孔编号
                {
                    cells[2, i + 1].PutValue(i);
                    cells[2, i + 1].SetStyle(format.Titlestyle(workbook1));
                }

                for (int i = 0; i < 15; i++)
                {
                    for (int j = 0; j < list_ComboBox.Count() + 4; j++)
                    {
                        if (data[j, i] == 0)
                            continue;
                        else
                        {
                            switch (data[j, i].ToString())
                            {
                                case "-1":
                                    cells[3 + j, 2 + i].PutValue("吸钻");
                                    cells[3 + j, 2 + i].SetStyle(format.Titlestyle(workbook1));
                                    break;
                                case "-2":
                                    cells[3 + j, 2 + i].PutValue("卡钻");
                                    cells[3 + j, 2 + i].SetStyle(format.Titlestyle(workbook1));
                                    break;
                                case "-3":
                                    cells[3 + j, 2 + i].PutValue("煤炮");
                                    cells[3 + j, 2 + i].SetStyle(format.Titlestyle(workbook1));
                                    break;
                                case "-4":
                                    cells[3 + j, 2 + i].PutValue("卡钻吸钻");
                                    cells[3 + j, 2 + i].SetStyle(format.Titlestyle(workbook1));
                                    break;
                                default:
                                    cells[3 + j, 2 + i].PutValue(data[j, i]);
                                    cells[3 + j, 2 + i].SetStyle(format.Titlestyle(workbook1));
                                    break;
                            }
                        }
                    }
                }
                try
                {
                    workbook1.Save(@"F:\Hello_test.xlsx");
                    MessageBox.Show("完成");
                }
                catch
                {
                    MessageBox.Show("该文件已打开，请关闭文件后再点击“完成预览”");
                }
            }
            string constr_test = "server=.;database=UPRESSURE;uid=sa;pwd=sakjdx";
            string constr = "server=192.168.1.111;database=UPRESSURE;uid=sa;pwd=sdkjdx";
            string sqlString_ins = "";

            using (SqlConnection sqlConnection = new SqlConnection(constr_test))
            {
                bool conok = SqlExtensions.QuickOpen(sqlConnection, 5000);  //连接到数据库
                if (conok)
                {
                    SqlCommand cmd_ins = new SqlCommand(sqlString_ins, sqlConnection);
                    cmd_ins.ExecuteNonQuery();
                    MessageBox.Show("完成");
                }
            }
        }


        private void button1_Click(object sender, EventArgs e)  //基本信息保存预览
        {
            string date = year + "-" + month + "-" + day;
            string date_y = year1 + "-" + month1 + "-" + day1;

            string constr = "server=192.168.1.111;database=UPRESSURE;uid=sa;pwd=sdkjdx";
            string constr_test = "server=.;database=UPRESSURE;uid=sa;pwd=sakjdx";
            //插入主键时间+工作面
            string sqlString_ins = string.Format(@"IF NOT EXISTS(SELECT 1 
                                                                 FROM [UPRESSURE].[dbo].[基本数据表] 
                                                                 WHERE  [日期] = '{0}' AND [工作面] = N'{1}')
                                                   BEGIN
                                                   INSERT INTO
                                                   [UPRESSURE].[dbo].[基本数据表]([日期], [工作面])
                                                   VALUES('{0}', N'{1}')
                                                   END", date, working_face);
            //寻找插入数据的上一条数据
            string sqlString_find = string.Format(@"SELECT TOP 1 *
                                                    FROM [UPRESSURE].[dbo].[基本数据表]
                                                    WHERE [工作面] LIKE N'{0}'
                                                    AND [日期] < '{1}' 
                                                    ORDER BY [日期] DESC", working_face, date);
            

            using (SqlConnection sqlConnection = new SqlConnection(constr_test))
            {
                bool conok = SqlExtensions.QuickOpen(sqlConnection, 5000);  //连接到数据库
                if (conok)
                {
                    SqlCommand cmd_ins = new SqlCommand(sqlString_ins, sqlConnection);
                    cmd_ins.ExecuteNonQuery();

                    DataTable datatable = new DataTable();
                    SqlCommand cmd_find = new SqlCommand(sqlString_find, sqlConnection);
                    using (SqlDataAdapter da = new SqlDataAdapter(cmd_find))
                    {
                        da.Fill(datatable);
                    }
                    //上一条记录进尺
                    double auxiliary_y = Convert.ToDouble(datatable.Rows[0][2].ToString());
                    double rubber_y = Convert.ToDouble(datatable.Rows[0][3].ToString());
                    //总进尺
                    auxiliary = (textBox1.Text == "") ? auxiliary_y : Convert.ToDouble(textBox1.Text); //辅
                    rubber =(textBox2.Text == "") ? rubber_y : Convert.ToDouble(textBox2.Text);   //胶
                    double transport_avg = Math.Round((auxiliary + rubber) / 2, 1);
                    //当日进尺
                    double auxiliary_td = Math.Round(auxiliary - auxiliary_y, 1);
                    double rubber_td = Math.Round(rubber - rubber_y, 1);
                    double transport_td_avg = Math.Round((auxiliary_td + rubber_td) / 2, 1);
                    //涌水量
                    double water = (textBox3.Text == "") ? 0.0 : Convert.ToDouble(textBox3.Text);

                    //DataRow dr = datatable.NewRow();
                    //object[] objs = { date, working_face, auxiliary, rubber, transport_avg, auxiliary_td, rubber_td, transport_td_avg, water, textBox4.Text, textBox5.Text, textBox6.Text, (2077 - transport_avg), textBox15.Text };
                    //dr.ItemArray = objs;
                    //datatable.Rows.Add(dr);
                    //datatable写入数据库
                    string sqlString_insdata = string.Format(@"UPDATE [UPRESSURE].[dbo].[基本数据表]
                                                               SET [辅运顺槽总进尺] = '{0}', [胶运顺槽总进尺] = '{1}',
                                                                   [总进尺平均] = '{2}', [辅运当日进尺] = '{3}',
                                                                   [胶运当日进尺] = '{4}', [当日平均] = '{5}',
                                                                   [工作面涌水量] = '{6}', [初采时间] = N'{7}',
                                                                   [实测倾斜长度] = '{8}', [平均采高] = '{9}',
                                                                   [剩余推进长度] = '{10}', [时空关系] = N'{11}'
                                                               WHERE [日期] = '{12}' AND [工作面] = N'{13}'",
                                                               auxiliary, rubber, transport_avg, auxiliary_td, rubber_td, transport_td_avg, water, textBox4.Text, textBox5.Text, textBox6.Text, (2077 - transport_avg), textBox15.Text, date, working_face);
                    SqlCommand cmd_insdata = new SqlCommand(sqlString_insdata, sqlConnection);
                    cmd_insdata.ExecuteNonQuery();
                    Properties.Settings.Default.Save();
                }
            }
            //charu
            gt.setToday(date);
            gt.setYesToday(date_y);
            bdt.Start(gt);
            MessageBox.Show("完成");
        }

        private void button3_Click(object sender, EventArgs e)  //工作面来压情况保存预览
        {
            string date = year + "-" + month + "-" + day;
            string constr_test = "server=.;database=UPRESSURE;uid=sa;pwd=sakjdx";

            //插入或更新来压情况
            string sqlString_ins = string.Format(@"IF NOT EXISTS(SELECT 1 
                                                                 FROM [UPRESSURE].[dbo].[工作面来压情况] 
                                                                 WHERE  [日期] = '{0}' AND [工作面] = N'{1}')
                                                   BEGIN
                                                   INSERT INTO
                                                   [UPRESSURE].[dbo].[工作面来压情况]([日期], [工作面],
                                                                                      [已来压次数],[上次位置],
                                                                                      [上次时间],[步距],
                                                                                      [本次来压情况],[持续距离],
                                                                                      [预计下次时间],[预计下次位置],
                                                                                      [预计下次步距],[下一危险区域名称],
                                                                                      [距离危险区域])
                                                   VALUES('{0}', N'{1}', '{2}', '{3}', N'{4}', '{5}', N'{6}', '{7}', N'{8}', '{9}', '{10}', N'{11}', '{12}')
                                                   END
                                                   ELSE
                                                   BEGIN
                                                   UPDATE [UPRESSURE].[dbo].[工作面来压情况]
                                                   SET [已来压次数] = '{2}', [上次位置] = '{3}',
                                                   [上次时间] = N'{4}', [步距] = '{5}',
                                                   [本次来压情况] = N'{6}', [持续距离] = '{7}',
                                                   [预计下次时间] = N'{8}', [预计下次位置] = '{9}',
                                                   [预计下次步距] = '{10}', [下一危险区域名称] = N'{11}',
                                                   [距离危险区域] = '{12}'
                                                   WHERE [日期] = '{0}' AND [工作面] = N'{1}'
                                                   END", 
                                                   date, working_face,
                                                   (textBox21.Text == "") ? 0 : Convert.ToInt32(textBox21.Text), //已来压次数
                                                   (textBox20.Text == "") ? 0.0 : Convert.ToDouble(textBox20.Text), //上次位置
                                                   (textBox19.Text == "") ? "无" : textBox19.Text, //上次时间
                                                   (textBox18.Text == "") ? 0.0 : Convert.ToDouble(textBox18.Text), //步距
                                                   (textBox17.Text == "") ? "无" : textBox17.Text, //本次来压情况
                                                   (textBox16.Text == "") ? 0.0 : Convert.ToDouble(textBox16.Text), //持续距离
                                                   (textBox26.Text == "") ? "无" : textBox26.Text, //预计下次时间
                                                   (textBox24.Text == "") ? 0.0 : Convert.ToDouble(textBox24.Text), //预计下次位置
                                                   (textBox22.Text == "") ? 0.0 : Convert.ToDouble(textBox22.Text), //预计下次步距
                                                   (textBox25.Text == "") ? "无" : textBox25.Text, //下一危险区域名称
                                                   (textBox23.Text == "") ? 0.0 : Convert.ToDouble(textBox23.Text)); //距离危险区域
            using (SqlConnection sqlConnection = new SqlConnection(constr_test))
            {
                bool conok = SqlExtensions.QuickOpen(sqlConnection, 5000);  //连接到数据库
                if (conok)
                {
                    SqlCommand cmd_ins = new SqlCommand(sqlString_ins, sqlConnection);
                    cmd_ins.ExecuteNonQuery();
                    MessageBox.Show("完成");
                }
            }
        }

        private void button6_Click(object sender, EventArgs e)  //应力在线保存预览
        {
            string constr_test = "server=.;database=UPRESSURE;uid=sa;pwd=sakjdx";
            string sqlString_ins = string.Format(@"IF NOT EXISTS(SELECT 1 
                                                                 FROM [UPRESSURE].[dbo].[地表沉降数据] 
                                                                 WHERE  [日期] = N'{0}')
                                                   BEGIN
                                                   INSERT INTO
                                                   [UPRESSURE].[dbo].[地表沉降数据]([日期], [最大沉降量],
                                                                                    [平均沉降量], [最大沉降位置])
                                                   VALUES(N'{0}', N'{1}', N'{2}', N'{3}')
                                                   END
                                                   ELSE
                                                   BEGIN
                                                   UPDATE [UPRESSURE].[dbo].[地表沉降数据]
                                                   SET [最大沉降量] = N'{1}', [平均沉降量] = N'{2}', [最大沉降位置] = N'{3}'
                                                   WHERE [日期] = N'{0}'
                                                   END",
                                                   (textBox7.Text == "") ? " " : textBox7.Text,
                                                   (textBox8.Text == "") ? null : textBox8.Text,
                                                   (textBox9.Text == "") ? null : textBox9.Text,
                                                   (textBox10.Text == "") ? null : textBox10.Text
                                                   );
            using (SqlConnection sqlConnection = new SqlConnection(constr_test))
            {
                bool conok = SqlExtensions.QuickOpen(sqlConnection, 5000);  //连接到数据库
                if (conok)
                {
                    SqlCommand cmd_ins = new SqlCommand(sqlString_ins, sqlConnection);
                    cmd_ins.ExecuteNonQuery();
                    Properties.Settings.Default.Save();
                    MessageBox.Show("完成");
                }
            }
        }
    }
}
