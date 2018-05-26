using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using System.Data.SqlClient;
using System.IO;


namespace KDTApiKit
{
    class database
    {
        string ConnString = "Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=youzan;Data Source=(local)";
        
        /// <summary>
        /// 写入订单信息
        /// </summary>
        /// <param name="tid">订单号</param>
        public IDictionary<String,String> inputOrderInfo(string tid)
        {

            Dictionary<String, String> param = new Dictionary<String, String>();
            //param.Add("tid", "E20160630210242047043238");
            param.Add("tid", tid);
            param.Add("fields", "tid,type,pay_type,buyer_nick,receiver_state,receiver_city,receiver_district,receiver_address,receiver_name,receiver_mobile,buyer_message,pay_time,payment,post_fee");

            //param.Add("page_no", "10");
            KDTApiKit kit = new KDTApiKit("557cd20013f8d999a4", "b1f9b79f0cfbd7123f3498274d2ab5a4");
            //a=kit.get("kdt.trades.sold.get", param);
            //a = kit.get("kdt.trade.get", param);

            string nihao1 = kit.get("kdt.trade.get", param);

            string[] nihaome1 = nihao1.Split(',');
            ArrayList nihao121 = new ArrayList();
            for (int i = 0; i < nihaome1.Length; i++)
            {
                nihao121.Add(nihaome1[i]);
            }
            //string path = @"G:\danbi.text";
            //FileStream fs = new FileStream(path, FileMode.Create);
            //StreamWriter sw = new StreamWriter(fs);
            Dictionary<String, String> writer = new Dictionary<string, string>();
            foreach (string i in nihao121)
            {

                string replace = i.Replace('}', ' ');
                replace = replace.Replace('{', ' ');
                replace = replace.Replace(" \"response\": \"trade\": ", "");

                string[] sonString = replace.Split(':');
                sonString[0] = sonString[0].Replace(" ", "");
                sonString[0] = sonString[0].Replace("\"", "");
                sonString[1] = sonString[1].Replace(" ", "");
                sonString[1] = sonString[1].Replace("\"", "");

                writer.Add(sonString[0], sonString[1]);

                //sw.Write(sonString[0] + "\n" + sonString[1] + "\n");

                //sw.Write(replace + "\n");

            }

            return writer;




            ////数据库部分

            //string ConnQuery = " insert into OrderInfo values('" + writer["tid"] + "','" + writer["type"] + "','" + writer["pay_type"] + "','" + writer["buyer_nick"] + "','" + writer["receiver_state"] + "','" + writer["receiver_city"] + "','" + writer["receiver_district"] + "','" + writer["receiver_address"] + "','" + writer["receiver_name"] + "','" + writer["receiver_mobile"] + "','" + writer["buyer_message"] + "','" + writer["payment"] + "','" + writer["post_fee"] + "')";
            //SqlConnection connection = new SqlConnection(ConnString);
            //connection.Open();
            //SqlCommand lo_cmd = new SqlCommand(ConnQuery, connection);
            //lo_cmd.ExecuteNonQuery();




            //connection.Close();
            //connection.Dispose();




        }


        /// <summary>
        /// 获取订单号集合
        /// </summary>
        /// <returns></returns>
        public List<string> getTid(string startTime,string endTime,string timeKind ,string status,int page_no)
        {
            Dictionary<String, String> param = new Dictionary<String, String>();
            //param.Add("tid", "E20160630210242047043238");

            param.Add("fields", "tid");
            param.Add("page_size", "500");
            param.Add("page_no", page_no.ToString());
            
            param.Add("start_"+timeKind,startTime);
            param.Add("end_" + timeKind, endTime);
            param.Add("status", status);

            KDTApiKit kit = new KDTApiKit("557cd20013f8d999a4", "b1f9b79f0cfbd7123f3498274d2ab5a4");
            string nihao1 = kit.get("kdt.trades.sold.get", param);




            string[] nihaome1 = nihao1.Split(',');
            ArrayList nihao121 = new ArrayList();
            for (int i = 0; i < nihaome1.Length; i++)
            {
                nihao121.Add(nihaome1[i]);
            }


            string newStr;
            List<string> returnStr = new List<string>();
            foreach (string i in nihao121)
            {

                newStr = i;
                newStr = newStr.Replace("{\"response\":{", "");
                newStr = newStr.Replace("\"trades\":[{", "");
                newStr = newStr.Replace("{", "");
                newStr = newStr.Replace("[", "");
                newStr = newStr.Replace("]", "");
                newStr = newStr.Replace("}", "");


                string[] sonString = newStr.Split(':');
                sonString[1] = sonString[1].Replace("\"", "");

                returnStr.Add(sonString[1]);



            }

            return returnStr;




        }

        /// <summary>
        /// 订单商品详情类型
        /// </summary>
       public struct TYPE
        {
            public string title;
            public decimal price;
            public int num;
            public decimal discount_fee;
            public decimal total_fee;
            public string state_str;

        }
        /// <summary>
        ///  写入订单商品详情
        /// </summary>
        /// <param name="tid"></param>
        public List<TYPE> inputOrderGoods(string tid)
        {



            Dictionary<String, String> param = new Dictionary<String, String>();
            //param.Add("tid", "E20160630210242047043238");

            param.Add("fields", "orders");
            param.Add("tid", tid);

            KDTApiKit kit = new KDTApiKit("557cd20013f8d999a4", "b1f9b79f0cfbd7123f3498274d2ab5a4");
            string nihao1 = kit.get("kdt.trade.get", param);


            string[] nihaome1 = nihao1.Split(',');
            ArrayList nihao121 = new ArrayList();
            for (int i = 0; i < nihaome1.Length; i++)
            {
                nihao121.Add(nihaome1[i]);
            }
            //string path = @"G:\danbi.text";
            //FileStream fs = new FileStream(path, FileMode.Create);
            //StreamWriter sw = new StreamWriter(fs);
            string newStr;

            Dictionary<string, string> OrdersDetail = new Dictionary<string, string>();
            List<TYPE> OrdersTable = new List<TYPE>(); // 订单表

            foreach (string i in nihao121)
            {
                int count = 0;

                newStr = i;
                newStr = newStr.Replace("{\"response\":{", "");
                newStr = newStr.Replace("\"trades\":[{", "");
                newStr = newStr.Replace("{", "");
                newStr = newStr.Replace("[", "");
                newStr = newStr.Replace("]", "");
                newStr = newStr.Replace("}", "");
                newStr = newStr.Replace("\"trade\":\"orders\":", "");

                string[] sonString = newStr.Split(':');
                sonString[0] = sonString[0].Replace("\"", "");
                sonString[1] = sonString[1].Replace("\"", "");

                TYPE one = new TYPE();
                OrdersDetail.Add(sonString[0], sonString[1]);

                //if (sonString[0] == "oid")
                //{
                //    count++;
                //}

                if (sonString[0] == "num")
                {

                    one.num = Convert.ToInt32(OrdersDetail["num"]);
                    one.price = Convert.ToDecimal(OrdersDetail["price"]);
                    one.state_str = OrdersDetail["state_str"];
                    one.title = OrdersDetail["title"];
                    one.total_fee = Convert.ToDecimal(OrdersDetail["total_fee"]);

                    OrdersTable.Add(one);   //  加入订单表



                    OrdersDetail.Clear();  //清空键值对
                    //count++;
                }





                //sw.Write(sonString[0] + ":"+sonString[1]+"\n");

            }

            return OrdersTable;


            //数据库
            //for (int k = 0; k <= OrdersTable.Count - 1; k++)
            //{
            //    string ConnQuery = " insert into OrderGoods values('" + tid + "','" + OrdersTable[k].title + "','" + OrdersTable[k].price + "','" + OrdersTable[k].num + "','" + OrdersTable[k].discount_fee + "','" + OrdersTable[k].total_fee + "','" + OrdersTable[k].state_str + "')";
            //    SqlConnection connection = new SqlConnection(ConnString);
            //    connection.Open();
            //    SqlCommand lo_cmd = new SqlCommand(ConnQuery, connection);
            //    lo_cmd.ExecuteNonQuery();




            //    connection.Close();
            //    connection.Dispose();
            //}
            ////清空缓冲区
            //sw.Flush();
            ////关闭流
            //sw.Close();
            //fs.Close();



        }


        /// <summary>
        /// 清空指定数据库
        /// </summary>
        /// <param name="tablename"></param>
        public void Clear(string tablename)
        {
            string ConnQuery = "truncate table " + tablename;
            SqlConnection connection = new SqlConnection(ConnString);
            connection.Open();
            SqlCommand lo_cmd = new SqlCommand(ConnQuery, connection);
            lo_cmd.ExecuteNonQuery();




            connection.Close();
            connection.Dispose();
        }



        public void test()
        {
            Dictionary<String, String> param = new Dictionary<String, String>();
            param.Add("start_created", "2016-7-1");
            param.Add("end_created", "2016-7-19");
            param.Add("status", "TRADE_BUYER_SIGNED");
            
            
            //param.Add("tid", tid);
            param.Add("fields", "tid");

            //param.Add("page_no", "10");
            KDTApiKit kit = new KDTApiKit("557cd20013f8d999a4", "b1f9b79f0cfbd7123f3498274d2ab5a4");
            //a=kit.get("kdt.trades.sold.get", param);
            //a = kit.get("kdt.trade.get", param);

            string nihao1 = kit.get("kdt.trades.sold.get", param);

            string[] nihaome1 = nihao1.Split(',');
            ArrayList nihao121 = new ArrayList();
            for (int i = 0; i < nihaome1.Length; i++)
            {
                nihao121.Add(nihaome1[i]);
            }
            string path = @"G:\danbi.text";
            FileStream fs = new FileStream(path, FileMode.Create);
            StreamWriter sw = new StreamWriter(fs);
            Dictionary<String, String> writer = new Dictionary<string, string>();
            foreach (string i in nihao121)
            {

                //string replace = i.Replace('}', ' ');
                //replace = replace.Replace('{', ' ');
                //replace = replace.Replace(" \"response\": \"trade\": ", "");

                //string[] sonString = replace.Split(':');
                //sonString[0] = sonString[0].Replace(" ", "");
                //sonString[0] = sonString[0].Replace("\"", "");
                //sonString[1] = sonString[1].Replace(" ", "");
                //sonString[1] = sonString[1].Replace("\"", "");

                

                //sw.Write(sonString[0] + ":" + sonString[1] + "\n");
                sw.Write(i + "\n");

                


            }
            //清空缓冲区
            sw.Flush();
            //关闭流
            sw.Close();
            fs.Close();

        }
    }
}
