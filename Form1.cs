using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Aspose.Cells;
using System.IO;

namespace WindowsFormsApplication2
{
    public partial class Form1 : Form
    {
        private double[,] data = new double[16, 16];
        private List<TextBox> list_TextBox = new List<TextBox>();   //每个孔煤粉量填写数值链表
        double[] data_test = new double[10];

        int count = 1;   //钻孔计数器
        int count1 = 0;   //单孔深度计数器
        string year = DateTime.Now.Year.ToString();     //日期
        string month = DateTime.Now.Month.ToString();
        string day = DateTime.Now.Day.ToString();
        string working_face;    //工作面
        string begin;
        int begin_face;int end_face;

        //private Size beforeResizeSize = Size.Empty;


        public Form1()
        {
            InitializeComponent();
            loadList();
        }

        /*protected override void OnResizeBegin(EventArgs e)
        {
            base.OnResizeBegin(e);
            beforeResizeSize = this.Size;
        }
        protected override void OnResizeEnd(EventArgs e)
        {
            base.OnResizeEnd(e);
            //窗口resize之后的大小
            Size endResizeSize = this.Size;
            //获得变化比例
            float percentWidth = (float)endResizeSize.Width / beforeResizeSize.Width;
            float percentHeight = (float)endResizeSize.Height / beforeResizeSize.Height;
            foreach (Control control in this.Controls)
            {
                if (control is DataGridView)
                    continue;
                //按比例改变控件大小
                control.Width = (int)(control.Width * percentWidth);
                control.Height = (int)(control.Height * percentHeight);
                //为了不使控件之间覆盖 位置也要按比例变化
                control.Left = (int)(control.Left * percentWidth);
                control.Top = (int)(control.Top * percentHeight);
            }
        }*/

        private void loadList()
        {
            list_TextBox.Add(textBox1);
            list_TextBox.Add(textBox2);
            list_TextBox.Add(textBox3);
            list_TextBox.Add(textBox4);
            list_TextBox.Add(textBox5);
            list_TextBox.Add(textBox6);
            list_TextBox.Add(textBox7);
            list_TextBox.Add(textBox8);
            list_TextBox.Add(textBox9);
            list_TextBox.Add(textBox10);
            list_TextBox.Add(textBox11);
            list_TextBox.Add(textBox12);
            list_TextBox.Add(textBox13);
            list_TextBox.Add(textBox14);
        }

        private Style Titlestyle(Workbook workbook)
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

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            string date = dateTimePicker1.Value.ToShortDateString(); //取年月日
            string date1 = dateTimePicker1.Value.AddDays(-1).ToString();
            year = dateTimePicker1.Value.Year.ToString();
            month = dateTimePicker1.Value.Month.ToString();
            day = dateTimePicker1.Value.Day.ToString();
            MessageBox.Show(date);
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //comboBox1.Items.Insert(0, "A");
            working_face = comboBox1.Text.ToString();
            MessageBox.Show(working_face);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if(String.IsNullOrWhiteSpace(working_face))
                MessageBox.Show("未选择工作面");
            else
            {
                string filename = working_face + "工作面钻屑监测统计表" + year + "年" + month + "月.xlsx";
                string path = @"F:\" + filename;
                string headline = month + "月" + day + "日轨顺槽钻屑监测统计表";
                string test = @"F:\钻屑监测统计表模板_test.xlsx";
                if (File.Exists(path))
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

                    cells.Merge(0, 2, 2, count1);
                    cells[0, 2].PutValue(headline);//填写标题
                    cells[0, 2].SetStyle(Titlestyle(workbook1));

                    for(int i = 1; i <= count1; i++)//填写钻孔编号
                    {
                        cells[2, i + 1].PutValue(i);
                        cells[2, i + 1].SetStyle(Titlestyle(workbook1));
                    }

                    for(int i = 0; i < 15; i++)
                    {
                        if (data[0, i] == 0)
                            break;
                        else
                        {
                            for (int j = 0; j < list_TextBox.Count(); j++)
                            {
                                if (data[j, i] == 0)
                                    break;
                                else
                                {
                                    cells[3 + j, 2 + i].PutValue(data[j, i]);
                                    cells[3 + j, 2 + i].SetStyle(Titlestyle(workbook1));
                                }
                            }
                        }
                    }

                    workbook1.Save(@"F:\Hello_test.xlsx");
                    MessageBox.Show(path);
                }  
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int tag = 0;    //标志是否出现卡钻等异常情况
            for(int i = 0; i < list_TextBox.Count(); i++)
            {
                TextBox textbox = list_TextBox[i];
                string temp = textbox.Text.ToString();
                if (String.IsNullOrWhiteSpace(temp))
                {
                    count1 = i;
                    break;
                }
                try
                {
                    if (temp == "吸钻" || temp == "卡钻" || temp == "煤炮" || temp == "卡钻吸钻" || temp == "吸钻卡钻")
                    {
                        tag = -1;
                        MessageBox.Show(tag.ToString());
                    }
                    else
                    {
                        tag = 0;
                        data[i, count-1] = double.Parse(temp);
                        //MessageBox.Show(data[i, count-1].ToString());
                    }
                }
                catch (Exception ex)
                {
                    string tips = "数据的五种格式为：“煤粉量数字”、“吸钻”、“卡钻”、“卡钻吸钻”“煤炮”";
                    tips = ex.Message + tips;
                    MessageBox.Show(tips);
                }
            }
            foreach(TextBox textbox in list_TextBox)
            {
                textbox.Clear();
            }
            count++;
        }



        private void Form1_Load(object sender, EventArgs e)
        {
           
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            //OnResizeEnd(e);
        }
    }
}
