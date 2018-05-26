using CrystalDecisions.CrystalReports.Engine;
using KDTApiKit;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace OrderPrint
{ 
    public partial class Form1 : Form
    {
        public string timeKind;
        public string status;
        public int PrintKind=0;
        public List<PrinterInfo> QuanJwriter = new List<PrinterInfo>();
        
        public Form1()
        {
            InitializeComponent();
            comboBox1.SelectedIndex=0;
            comboBox2.SelectedIndex = 0;
            progressBar1.Hide();
           
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void dateTimeStart_ValueChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            PrintKind = 0;
            dataGridView1.Columns[1].Visible = true;
            database a=new database();
            if(comboBox1.SelectedItem.ToString()=="交易状态更新的时间")
            {
                timeKind = "update";

            }
            if(comboBox1.SelectedItem.ToString()=="交易创建的时间")
            {
                timeKind = "created";

            }

            //状态
            if(comboBox2.SelectedItem.ToString()=="没有创建支付交易")
            {
                status = "TRADE_NO_CREATE_PAY";

            }
            if (comboBox2.SelectedItem.ToString() == "等待买家付款")
            {
                status = "WAIT_BUYER_PAY";

            }
            if (comboBox2.SelectedItem.ToString() == "等待成团")
            {
                status = "WAIT_GROUP";

            }

            if (comboBox2.SelectedItem.ToString() == "等待卖家发货")
            {
                status = "WAIT_SELLER_SEND_GOODS";

            }

            if (comboBox2.SelectedItem.ToString() == "等待买家确认收货")
            {
                status = "WAIT_BUYER_CONFIRM_GOODS";

            }
            if (comboBox2.SelectedItem.ToString() == "买家已签收")
            {
                status = "TRADE_BUYER_SIGNED";

            }

            if (comboBox2.SelectedItem.ToString() == "付款以后用户退款成功，交易自动关闭")
            {
                status = "TRADE_CLOSED";

            }

            if (comboBox2.SelectedItem.ToString() == "没有创建支付交易或等待买家付款")
            {
                status = "ALL_WAIT_PAY";

            }

            if (comboBox2.SelectedItem.ToString() == "所有关闭订单")
            {
                status = "ALL_CLOSED";

            }

            string starTime= dateTimeStart.Value.ToString();
            string endTime= dateTimePicker3.Value.ToString();

            List<string> tid =a.getTid(starTime,endTime,timeKind,status,1);   //获取订单号

           

            

            
            dataGridView1.Rows.Clear();//清空表格


            if (tid[1] == "\u672a\u6307\u5b9a\u6b63\u786e\u7684\u8ba2\u5355\u65f6\u95f4\u8303\u56f4")
            { MessageBox.Show("未指定正确的订单范围"); }
            if (tid[0] != "")
            {
                int MaxPage = Convert.ToInt32(tid[0]) / 500 + 1;
                int MaxNum = Convert.ToInt32(tid[0]);   //总共订单个数

                progressBar1.Maximum = Convert.ToInt32(tid[0]);
                progressBar1.Show();
                for (int i = 0; i <= MaxNum - 1; i++)
                {
                    
                    progressBar1.Value = i+1;
                    
                    if(((i+1)%500)==0)
                    {
                        
                        tid = a.getTid(starTime, endTime, timeKind, status, i / 500 + 1);

                    }

                    if ((i + 1) <= tid.Count - 1)
                    {
                        IDictionary<String, String> info = a.inputOrderInfo(tid[i + 1]);

                        

                        dataGridView1.Rows.Add();

                        dataGridView1.Rows[i].Cells["tid"].Value = info["tid"];
                        dataGridView1.Rows[i].Cells["receiveName"].Value = unicode_js_1(info["receiver_name"]);
                        dataGridView1.Rows[i].Cells["TEL"].Value = info["receiver_mobile"];
                        dataGridView1.Rows[i].Cells["Address"].Value = unicode_js_1(info["receiver_state"] + info["receiver_city"] + info["receiver_district"] + info["receiver_address"]);
                        dataGridView1.Rows[i].Cells["xiangInfo"].Value = "订单详情";

                    }

                }
                progressBar1.Hide();
            }
            else if (tid[1] == "\u672a\u6307\u5b9a\u6b63\u786e\u7684\u8ba2\u5355\u65f6\u95f4\u8303\u56f4")
            { MessageBox.Show("未指定正确的订单范围"); }
            
            else
            {
                MessageBox.Show("没有相应订单");
            }

        }


        public static string unicode_js_1(string str)
        {
            string outStr = "";
            Regex reg = new Regex(@"(?i)\\u([0-9a-f]{4})");
            outStr = reg.Replace(str, delegate(Match m1)
            {
                return ((char)Convert.ToInt32(m1.Groups[1].Value, 16)).ToString();
            });
            return outStr;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

            if (dataGridView1.Columns[e.ColumnIndex].Name == "xiangInfo")
            {
                if (PrintKind == 0)
                {
                    xiangqing newXiang = new xiangqing(dataGridView1.Rows[e.RowIndex].Cells["tid"].Value.ToString());
                    newXiang.Show();
                }
                if(PrintKind==1)
                {
                    xiangqing newXiang = new xiangqing(dataGridView1.Rows[e.RowIndex].Cells["TEL"].Value.ToString(),QuanJwriter);
                    newXiang.Show();
                }

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {     
               Random rd = new Random();
               int r = rd.Next(256); int g = rd.Next(256); int b = rd.Next(256);
               for(int i=0;i<dataGridView1.RowCount;i++)
               {
                   if (dataGridView1.Rows[i].Cells["Select"].EditedFormattedValue.ToString() == "True")
                   {   

                       if(PrintKind==0)   // 打印1
                       { 
                           printShow(i, 0);
                       }

                       if (PrintKind == 1)  //打印2
                       {
                           printShowEXEL(i,0);
                       }
                       
                      
                        

                       dataGridView1.Rows[i].Cells["tid"].Style.ForeColor = Color.FromArgb(r,g,b);
                       dataGridView1.Rows[i].Cells["receiveName"].Style.ForeColor = Color.FromArgb(r, g, b);
                       dataGridView1.Rows[i].Cells["TEL"].Style.ForeColor = Color.FromArgb(r, g, b);
                       dataGridView1.Rows[i].Cells["Address"].Style.ForeColor = Color.FromArgb(r, g, b);
                   
                   }


               }
        }

        private void button3_Click(object sender, EventArgs e)
        {
              for(int i=0;i<dataGridView1.RowCount;i++)
              {
                  dataGridView1.Rows[i].Cells["Select"].Value="True";

              }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                dataGridView1.Rows[i].Cells["Select"].Value = "False";

            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
           
        }


        private void printShow(int row, int count)
        {


            
            moban newmoban = new moban();

            //收货人姓名
            TextObject RecevieName = (TextObject)newmoban.ReportDefinition.ReportObjects["RecevieName"];
            RecevieName.Text = dataGridView1.Rows[row].Cells["receiveName"].Value.ToString();

            //收货人电话
            TextObject ReceiveTel = (TextObject)newmoban.ReportDefinition.ReportObjects["ReceiveTel"];
            ReceiveTel.Text = dataGridView1.Rows[row].Cells["TEL"].Value.ToString();

            //提货点名称
            TextObject PickPointName = (TextObject)newmoban.ReportDefinition.ReportObjects["PickPointName"];
            PickPointName.Text = dataGridView1.Rows[row].Cells["Address"].Value.ToString();


            //订单号

            //TextObject OrderNo = (TextObject)newmoban.ReportDefinition.ReportObjects["OrderNo"];
            //OrderNo.Text = dataGridView1.Rows[row].Cells["tid"].Value.ToString() ;

            xiangqing goodsinfo = new xiangqing(dataGridView1.Rows[row].Cells["tid"].Value.ToString());



            //订单详情
            TextObject commidityInfo = (TextObject)newmoban.ReportDefinition.ReportObjects["commidityInfo"];
            for (; count < goodsinfo.dataGridView1.RowCount - 1; count++)
            {
                if ((count + 1) % 5 == 0)
                {
                    printShow(row, count + 1);
                    
                    commidityInfo.Text += (count + 1).ToString() + "," + goodsinfo.dataGridView1.Rows[count].Cells["GoodsName"].Value.ToString() + "\n"
                   + "数量：" + goodsinfo.dataGridView1.Rows[count].Cells["GoodsNum"].Value.ToString() + "份" +"\n";
                    break;

                }
                commidityInfo.Text += (count + 1).ToString() + "," + goodsinfo.dataGridView1.Rows[count].Cells["GoodsName"].Value.ToString() + "\n"
                    + "数量：" + goodsinfo.dataGridView1.Rows[count].Cells["GoodsNum"].Value.ToString() + "份" + "\n";



            }

            //小计

            database a = new database();

            TextObject subtotal = (TextObject)newmoban.ReportDefinition.ReportObjects["subtotal"];
            subtotal.Text = "订单总价￥:" + a.inputOrderInfo(dataGridView1.Rows[row].Cells["tid"].Value.ToString())["payment"] + "\n";
                           


            //下单时间

            TextObject PayTime = (TextObject)newmoban.ReportDefinition.ReportObjects["PayTime"];
            PayTime.Text = a.inputOrderInfo(dataGridView1.Rows[row].Cells["tid"].Value.ToString())["pay_time"];

            //表尾
            //收货人姓名2
            TextObject RecevieName2 = (TextObject)newmoban.ReportDefinition.ReportObjects["RecevieName2"];
            RecevieName2.Text = dataGridView1.Rows[row].Cells["receiveName"].Value.ToString();

            //收货人电话2
            TextObject ReceiveTel2 = (TextObject)newmoban.ReportDefinition.ReportObjects["ReceiveTel2"];
            ReceiveTel2.Text = dataGridView1.Rows[row].Cells["TEL"].Value.ToString();

            //提货点名称2
            TextObject PickPointName2 = (TextObject)newmoban.ReportDefinition.ReportObjects["PickPointName2"];
            PickPointName2.Text = dataGridView1.Rows[row].Cells["Address"].Value.ToString();

            TextObject OrderNo2 = (TextObject)newmoban.ReportDefinition.ReportObjects["OrderNo2"];
            OrderNo2.Text = dataGridView1.Rows[row].Cells["tid"].Value.ToString();

           // 表尾数字订单号
            TextObject Text1 = (TextObject)newmoban.ReportDefinition.ReportObjects["Text1"];
            Text1.Text = dataGridView1.Rows[row].Cells["tid"].Value.ToString();




            //crystalReportViewer1.ReportSource = newinfo;

            //crystalReportViewer1.Show();
            //newinfo.PrintOptions.PrinterName=

            //newmoban.PrintOptions.PrinterName = DefaultPrinter;   // 设置打印机名称



            System.Drawing.Printing.PrintDocument doc = new System.Drawing.Printing.PrintDocument();
            //List<string> Printer = new List<string>();



            //foreach (string fPrinterName in LocalPrinter.GetLocalPrinters())
            //{

            //    Printer.Add(fPrinterName);



            //}
            ////doc.PrinterSettings.PrinterName = DefaultPrinter;
            //doc.PrinterSettings.PrinterName = Printer[0].ToString();

            //LocalPrinter print=new LocalPrinter();
            //newmoban.PrintOptions.PrinterName=print.DefaultPrinter  ;
            int rawKind = 1;
            for (int i = 0; i <= doc.PrinterSettings.PaperSizes.Count - 1; i++)
            {
                if (doc.PrinterSettings.PaperSizes[i].PaperName == "xiannvguo")
                {
                    rawKind = doc.PrinterSettings.PaperSizes[i].RawKind;
                }
            }


            newmoban.PrintOptions.PrinterName = doc.PrinterSettings.PrinterName;

            newmoban.PrintOptions.PaperSize = (CrystalDecisions.Shared.PaperSize)rawKind;    // 设置打印纸张样式
            newmoban.PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.DefaultPaperOrientation;//默认纸张方向
            newmoban.PrintToPrinter(1, false, 1, 1);

            return;


        }



        private void printShowEXEL(int row, int count)
        {

            moban newmoban = new moban();

            //收货人姓名
            TextObject RecevieName = (TextObject)newmoban.ReportDefinition.ReportObjects["RecevieName"];
            RecevieName.Text = dataGridView1.Rows[row].Cells["receiveName"].Value.ToString();

            //收货人电话
            TextObject ReceiveTel = (TextObject)newmoban.ReportDefinition.ReportObjects["ReceiveTel"];
            ReceiveTel.Text = dataGridView1.Rows[row].Cells["TEL"].Value.ToString();

            //提货点名称
            TextObject PickPointName = (TextObject)newmoban.ReportDefinition.ReportObjects["PickPointName"];
            PickPointName.Text = dataGridView1.Rows[row].Cells["Address"].Value.ToString();


            ////订单号

            //TextObject OrderNo = (TextObject)newmoban.ReportDefinition.ReportObjects["OrderNo"];
            //OrderNo.Text = "*" + dataGridView1.Rows[row].Cells["TEL"].Value.ToString().Replace("\t","") + "*";

            ////QuanJwriter




            xiangqing goodsinfo = new xiangqing(dataGridView1.Rows[row].Cells["TEL"].Value.ToString(),QuanJwriter);
            //订单详情
            TextObject commidityInfo = (TextObject)newmoban.ReportDefinition.ReportObjects["commidity"];
            for (; count < goodsinfo.dataGridView1.RowCount - 1; count++)
            {
                if ((count + 1) % 5 == 0)
                {
                    printShowEXEL(row, count + 1);
                    commidityInfo.Text += (count + 1).ToString() + "," + goodsinfo.dataGridView1.Rows[count].Cells["GoodsName"].Value.ToString() + "\n"
                   + "数量：" + goodsinfo.dataGridView1.Rows[count].Cells["GoodsNum"].Value.ToString() + "份" + "\n";
                    break;

                }
                commidityInfo.Text += (count + 1).ToString() + "," + goodsinfo.dataGridView1.Rows[count].Cells["GoodsName"].Value.ToString() + "\n"
                    + "数量：" + goodsinfo.dataGridView1.Rows[count].Cells["GoodsNum"].Value.ToString() + "份" + "\n";



            }

            //小计

            //database a = new database();

            //TextObject subtotal = (TextObject)newmoban.ReportDefinition.ReportObjects["subtotal"];
            //subtotal.Text = "订单总价￥:" + a.inputOrderInfo(dataGridView1.Rows[row].Cells["tid"].Value.ToString())["payment"] + "\n";



            //下单时间

            //TextObject PayTime = (TextObject)newmoban.ReportDefinition.ReportObjects["PayTime"];
            //PayTime.Text = a.inputOrderInfo(dataGridView1.Rows[row].Cells["tid"].Value.ToString())["pay_time"];

            //表尾
            //收货人姓名2
            //TextObject RecevieName2 = (TextObject)newmoban.ReportDefinition.ReportObjects["RecevieName2"];
            //RecevieName2.Text = dataGridView1.Rows[row].Cells["receiveName"].Value.ToString();

            ////收货人电话2
            //TextObject ReceiveTel2 = (TextObject)newmoban.ReportDefinition.ReportObjects["ReceiveTel2"];
            //ReceiveTel2.Text = dataGridView1.Rows[row].Cells["TEL"].Value.ToString();

            ////提货点名称2
            //TextObject PickPointName2 = (TextObject)newmoban.ReportDefinition.ReportObjects["PickPointName2"];
            //PickPointName2.Text = dataGridView1.Rows[row].Cells["Address"].Value.ToString();

            //TextObject OrderNo2 = (TextObject)newmoban.ReportDefinition.ReportObjects["OrderNo2"];
            //OrderNo2.Text = dataGridView1.Rows[row].Cells["tid"].Value.ToString();

            // 表尾数字订单号
            //TextObject Text1 = (TextObject)newmoban.ReportDefinition.ReportObjects["Text1"];
            //Text1.Text = dataGridView1.Rows[row].Cells["tid"].Value.ToString();




            //crystalReportViewer1.ReportSource = newinfo;

            //crystalReportViewer1.Show();
            //newinfo.PrintOptions.PrinterName=

            //newmoban.PrintOptions.PrinterName = DefaultPrinter;   // 设置打印机名称



            System.Drawing.Printing.PrintDocument doc = new System.Drawing.Printing.PrintDocument();
            //List<string> Printer = new List<string>();



            //foreach (string fPrinterName in LocalPrinter.GetLocalPrinters())
            //{

            //    Printer.Add(fPrinterName);



            //}
            ////doc.PrinterSettings.PrinterName = DefaultPrinter;
            //doc.PrinterSettings.PrinterName = Printer[0].ToString();

            //LocalPrinter print=new LocalPrinter();
            //newmoban.PrintOptions.PrinterName=print.DefaultPrinter  ;
            int rawKind = 1;
            for (int i = 0; i <= doc.PrinterSettings.PaperSizes.Count - 1; i++)
            {
                if (doc.PrinterSettings.PaperSizes[i].PaperName == "订单打印")
                {
                    rawKind = doc.PrinterSettings.PaperSizes[i].RawKind;
                }
            }


            newmoban.PrintOptions.PrinterName = doc.PrinterSettings.PrinterName;

            newmoban.PrintOptions.PaperSize = (CrystalDecisions.Shared.PaperSize)rawKind;    // 设置打印纸张样式
            newmoban.PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.DefaultPaperOrientation;//默认纸张方向
            newmoban.PrintToPrinter(1, false, 1, 1);

            newmoban.Dispose();
            return;




        }

        private void button5_Click(object sender, EventArgs e)
        {   

            PrintKind=1;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";//设置打开文件类型
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if (openFileDialog1.FileName != "")
                {


                    EXEL newinfo = new EXEL();


                    DataTable a = newinfo.LoadDataFromExcel(openFileDialog1.FileName);
                    List<PrinterInfo> writer = new List<PrinterInfo>();

                    PrinterInfo one = null;
                    GoodsInfo son = null;



                    // 获取整理表格信息
                    for (int i = 1; i <= a.Rows.Count - 1; i++)
                    {
                        //PrinterInfo one = new PrinterInfo();
                        //GoodsInfo son = new GoodsInfo();
                        //one.Goods = new List<GoodsInfo>();
                        // one.Name=a.Rows[i][6].ToString();
                        // one.TEL=a.Rows[i][7].ToString();
                        // one.receiveAd = a.Rows[i][5].ToString();
                        // son.num = Convert.ToInt32(a.Rows[i][4]);
                        // son.title = a.Rows[i][2].ToString();

                        //one.Goods.Add(son);

                        if (a.Rows[i][6].ToString() != "")
                        {
                            if (a.Rows[i][6].ToString() != a.Rows[i - 1][6].ToString() && a.Rows[i][7].ToString() != a.Rows[i - 1][7].ToString())
                            {
                                one = new PrinterInfo();
                                son = new GoodsInfo();
                                one.Goods = new List<GoodsInfo>();
                                one.Name = a.Rows[i][6].ToString();
                                one.TEL = a.Rows[i][7].ToString();
                                one.receiveAd = a.Rows[i][5].ToString();

                                son.num = Convert.ToInt32(a.Rows[i][4]);
                                son.title = a.Rows[i][2].ToString();

                                one.Goods.Add(son);

                                writer.Add(one);
                            }
                            else
                            {
                                son = new GoodsInfo();
                                one.Name = a.Rows[i][6].ToString();
                                one.TEL = a.Rows[i][7].ToString();
                                one.receiveAd = a.Rows[i][5].ToString();

                                son.num = Convert.ToInt32(a.Rows[i][4]);
                                son.title = a.Rows[i][2].ToString();

                                one.Goods.Add(son);

                            }
                        }
                    }

                        dataGridView1.Columns[1].Visible = false;
                       //导入表格
                        

                        dataGridView1.Rows.Clear();   //清空表格
                        for (int k = 0; k <= writer.Count-1; k++)
                        {





                            dataGridView1.Rows.Add();

                            dataGridView1.Rows[k].Cells["tid"].Value = "";
                            dataGridView1.Rows[k].Cells["receiveName"].Value = writer[k].Name.ToString();
                            dataGridView1.Rows[k].Cells["TEL"].Value = writer[k].TEL.ToString();
                            dataGridView1.Rows[k].Cells["Address"].Value = writer[k].receiveAd.ToString();
                            dataGridView1.Rows[k].Cells["xiangInfo"].Value = "订单详情";

                        }

                         QuanJwriter = writer;// 返回全局变量

                    }
                }
            }
        }


        public class GoodsInfo
        {
            public string title;
            public int num;

        }
        public class PrinterInfo
        {

            public List<GoodsInfo> Goods; //商品信息
            public string receiveAd;  //收货地址
            public string Name;    //姓名
            public string TEL;  //电话

        }
        public class EXEL
        {
            public DataTable LoadDataFromExcel(string Path)
            {
                string strConn = "Provider=Microsoft.Ace.OleDb.12.0;" + "data source=" + Path + ";Extended Properties='Excel 12.0; HDR=NO; IMEX=1'"; //此連接可以操作.xls與.xlsx文件
                OleDbConnection conn = new OleDbConnection(strConn);
                conn.Open();
                string strExcel = "";
                OleDbDataAdapter myCommand = null;
                DataTable dt = null;
                strExcel = "select * from [Sheet$]";
                myCommand = new OleDbDataAdapter(strExcel, strConn);
                dt = new DataTable();
                myCommand.Fill(dt);
                return dt;
            }

        }


        }

   

