using KDTApiKit;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;


namespace OrderPrint
{
    public partial class xiangqing : Form
    {

        public List<PrinterInfo> QuanJwriter = new List<PrinterInfo>();
        public xiangqing(string tid)
        {
            InitializeComponent();
            
                database a = new database();
                for (int i = 0; i < a.inputOrderGoods(tid).Count; i++)
                {
                    dataGridView1.Rows.Add();
                    dataGridView1.Rows[i].Cells["GoodsName"].Value = unicode_js_1(a.inputOrderGoods(tid)[i].title);
                    dataGridView1.Rows[i].Cells["GoodsNum"].Value = a.inputOrderGoods(tid)[i].num;

                }
            
         
            
            
        }
        //重载
        public xiangqing(string TEL, List<PrinterInfo> Q)
        {
            InitializeComponent();

            QuanJwriter = Q;
            int num=0;

            for (int k = 0; k <= Q.Count - 1;k++ )
            {
                if(Q[k].TEL.ToString()==TEL)
                {
                    num = k;
                    break;
                }


            }

                for (int i = 0; i < Q[num].Goods.Count; i++)
                {
                    dataGridView1.Rows.Add();
                    dataGridView1.Rows[i].Cells["GoodsName"].Value = Q[num].Goods[i].title;
                    dataGridView1.Rows[i].Cells["GoodsNum"].Value = Q[num].Goods[i].num;

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

        }
    }
}
