using Microsoft.AspNetCore.Http;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Wendy
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
           
        }
        public List<string> list_size;
        public List<string> list_color;
        List<goods> resultgood = new List<goods>();
        public string shop = "";
        private void button1_Click(object sender, EventArgs e)
        {
                string goods_no = txtAsin1.Text;
            if (string.IsNullOrWhiteSpace(goods_no))
            {
                MessageBox.Show("请输入正确的asin!");
                this.Hide();
                return;
            }
            //通过一个asin获取到该商品目录下所有商品的asin
           // List<string>


            List<goods> listgood = new List<goods>();
            goods good = new goods();
            good.asin = goods_no;
            listgood = listgood.Concat(getGoodsValue(goods_no)).ToList<goods>();
        }

        //封装抽象方法：通过asin获取到该商品的价格、尺寸等属性
        public List<goods> getGoodsValue(string asin)
        {
            //初始化
            List<goods> listgood = new List<goods>();
            goods good = new goods();

            var url = new Uri("https://www.amazon.de/dp/" + asin + "/ref=redir_mobile_desktop?_encoding=UTF8");
            Task task = new Task(() => {
                HttpWebRequest httpWebRequest = (HttpWebRequest)HttpWebRequest.Create(url);
                //请求的ContentType必须这样设置
                httpWebRequest.ContentType = "text/html;charset=UTF-8";
                HttpWebResponse httpWebResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                string result = "";
                using (StreamReader sr = new StreamReader(httpWebResponse.GetResponseStream(), Encoding.UTF8))
                {
                    result = sr.ReadToEnd();
                }
                //价格
                string t = @"<span id=""priceblock_ourprice"" class=""a-size-medium a-color-price priceBlockBuyingPriceString"">(.*?)</span>";
                string t_color = result.Substring(result.IndexOf("selected_variations"), 100).Trim().Replace("\n", "").Replace("\t", "").Replace("\r", ""); ;
                Regex regex = new Regex(t, RegexOptions.Multiline | RegexOptions.Singleline);
                result = regex.Match(result).Value;
                if (result != "")
                {
                    good.price = result.Substring(result.IndexOf("priceBlockBuyingPriceString") + 29, 6);
                }
                else
                {
                    good.price = "*";
                }

                if (t_color != "")
                {
                    good.size = t_color.Substring(t_color.IndexOf("size_name") + 12, t_color.IndexOf("color_name") - t_color.IndexOf("size_name") - 15);
                    good.color = t_color.Substring(t_color.IndexOf("color_name") + 12, t_color.IndexOf("},") - t_color.IndexOf("color_name") - 12);
                }
                else
                {
                    good.size = "*";
                    good.color = "*";
                }
                httpWebResponse.Dispose();
            });
            task.Start();
            good.shop = comboBox1.Text;
            listgood.Add(good);
            return listgood;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void txtAsin1_TextChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = 0;
            shop = comboBox1.Text;
        }

        private void button2_Click(object sender, EventArgs e)
        {

            var url = new Uri("https://www.amazon.de/dp/B07L4431WF/ref=redir_mobile_desktop?_encoding=UTF8");
            Task task = new Task(() => {
                HttpWebRequest httpWebRequest = (HttpWebRequest)HttpWebRequest.Create(url);
                //请求的ContentType必须这样设置
                httpWebRequest.ContentType = "text/html;charset=UTF-8";
                HttpWebResponse httpWebResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                string result = "";
                using (StreamReader sr = new StreamReader(httpWebResponse.GetResponseStream(), Encoding.UTF8))
                {
                    result = sr.ReadToEnd().Trim().Replace("\n", "").Replace("\t", "").Replace("\r", "");
                }

                string sizeAndColor = result.Substring(result.IndexOf("\"variationValues\"") +21, result.IndexOf("asinVariationValues") - result.IndexOf("\"variationValues\"") - 24);

                string str_size = sizeAndColor.Substring(sizeAndColor.IndexOf("\"size_name\":[")+13, sizeAndColor.IndexOf("],\"color_name\":[") - sizeAndColor.IndexOf("\"size_name\":[") - 16).Replace("\"", "");
                string str_Color= sizeAndColor.Substring(sizeAndColor.IndexOf("color_name") + 13).Replace("]","").Replace("\"", "");
                //尺寸组
                String[] sArray  = str_size.Split(",");
                list_size= new List<string>(sArray);
                //颜色组
                sArray = str_Color.Split(",");
                list_color = new List<string>(sArray);
                //转json
                 sizeAndColor = result.Substring(result.IndexOf("asinVariationValues") + 23, result.IndexOf("dimensionValuesData") - result.IndexOf("asinVariationValues") - 25);
                //获取除了价格之外的属性
                resultgood= ConvertJsonString(sizeAndColor);
                httpWebResponse.Dispose();
            });
            task.Start();
            task.Wait();
            //获取价格
            //foreach (var item in resultgood)
            //{
            //    item.price = GetPrice(item.asin);
            //}
            DataTable dt= ToDataTable<goods>(resultgood);
           ExportExcel(dt);
            MessageBox.Show("导出成功，请到D盘根目录下查看excel");
        }

        //字符串转json
        private List<goods> ConvertJsonString(string str)
        {
            //Data jsonData = JsonConvert.DeserializeObject<Data>(str);


            //格式化json字符串
            JsonSerializer serializer = new JsonSerializer();
            TextReader tr = new StringReader(str);
            JsonTextReader jtr = new JsonTextReader(tr);
            object obj = serializer.Deserialize(jtr);
            if (obj != null)
            {
                // JObject jObject = JObject.Parse(builder.ToString());
               


                StringWriter textWriter = new StringWriter();
                JsonTextWriter jsonWriter = new JsonTextWriter(textWriter)
                {
                    Formatting = Formatting.Indented,
                    Indentation = 4,
                    IndentChar = ' '
                };
                serializer.Serialize(jsonWriter, obj);

                var o = JObject.Parse(textWriter.ToString());
               
                foreach (JToken child in o.Children())
                {
                    //var property1 = child as JProperty;  
                    //MessageBox.Show(property1.Name + ":" + property1.Value);  
                    foreach (JToken grandChild in child)
                    {
                        goods good = new goods();
                        good.shop = shop;
                        foreach (JToken grandGrandChild in grandChild)
                        {
                            var property = grandGrandChild as JProperty;
                            if (property != null)
                            {
                                switch (property.Name)
                                {
                                    case "size_name": good.size = list_size[Convert.ToInt16(property.Value)]; break;
                                    case "ASIN": good.asin = property.Value.ToString(); break;
                                    case "color_name": good.color = list_color[Convert.ToInt16(property.Value)]; break;
                                }
                            }
                        }
                        resultgood.Add(good);
                    }
                }


            }
            return resultgood;
        }

        private string GetPrice(string asin)
        {
            string price = "";
            var url = new Uri("https://www.amazon.de/dp/"+asin+"/ref=redir_mobile_desktop?_encoding=UTF8");
            Task task = new Task(() => {
                HttpWebRequest httpWebRequest = (HttpWebRequest)HttpWebRequest.Create(url);
                //请求的ContentType必须这样设置
                httpWebRequest.ContentType = "text/html;charset=UTF-8";
                HttpWebResponse httpWebResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                string result = "";
                using (StreamReader sr = new StreamReader(httpWebResponse.GetResponseStream(), Encoding.UTF8))
                {
                    result = sr.ReadToEnd().Trim().Replace("\n", "").Replace("\t", "").Replace("\r", "");
                }
                //价格
                string t = @"<span id=""priceblock_ourprice"" class=""a-size-medium a-color-price priceBlockBuyingPriceString"">(.*?)</span>";
                Regex regex = new Regex(t, RegexOptions.Multiline | RegexOptions.Singleline);
                result = regex.Match(result).Value;
                if (result != "")
                {
                    price= result.Substring(result.IndexOf("priceBlockBuyingPriceString") + 29, 6);
                }
                httpWebResponse.Dispose();
            });
            task.Start();
            task.Wait();
            return price;
        }

        //写入到excel
        public void ExportExcel(DataTable dt)
        {
            try
            {
                //创建一个工作簿
                IWorkbook workbook = new HSSFWorkbook();

                //创建一个 sheet 表
                ISheet sheet = workbook.CreateSheet(dt.TableName);

                //创建一行
                IRow rowH = sheet.CreateRow(0);

                //创建一个单元格
                ICell cell = null;

                //创建单元格样式
                ICellStyle cellStyle = workbook.CreateCellStyle();

                //创建格式
                IDataFormat dataFormat = workbook.CreateDataFormat();

                //设置为文本格式，也可以为 text，即 dataFormat.GetFormat("text");
                cellStyle.DataFormat = dataFormat.GetFormat("@");

                //设置列名
                foreach (DataColumn col in dt.Columns)
                {
                    //创建单元格并设置单元格内容
                    rowH.CreateCell(col.Ordinal).SetCellValue(col.Caption);

                    //设置单元格格式
                    rowH.Cells[col.Ordinal].CellStyle = cellStyle;
                }

                //写入数据
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    //跳过第一行，第一行为列名
                    IRow row = sheet.CreateRow(i + 1);

                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        cell = row.CreateCell(j);
                        cell.SetCellValue(dt.Rows[i][j].ToString());
                        cell.CellStyle = cellStyle;
                    }
                }

                //设置导出文件路径
                string path = "D:\\";

                //设置新建文件路径及名称
                string savePath = path + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss") + ".xls";

                //创建文件
                FileStream file = new FileStream(savePath, FileMode.CreateNew, FileAccess.Write);

                //创建一个 IO 流
                MemoryStream ms = new MemoryStream();

                //写入到流
                workbook.Write(ms);

                //转换为字节数组
                byte[] bytes = ms.ToArray();

                file.Write(bytes, 0, bytes.Length);
                file.Flush();

                //还可以调用下面的方法，把流输出到浏览器下载
               // OutputClient(bytes);

                //释放资源
                bytes = null;

                ms.Close();
                ms.Dispose();

                file.Close();
                file.Dispose();

                workbook.Close();
                sheet = null;
                workbook = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        //list转datatable
        private DataTable ToDataTable<T>(List<T> items)
        {
            var tb = new DataTable(typeof(T).Name);

            PropertyInfo[] props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);

            foreach (PropertyInfo prop in props)
            {
                Type t = GetCoreType(prop.PropertyType);
                tb.Columns.Add(prop.Name, t);
            }

            foreach (T item in items)
            {
                var values = new object[props.Length];

                for (int i = 0; i < props.Length; i++)
                {
                    values[i] = props[i].GetValue(item, null);
                }

                tb.Rows.Add(values);
            }

            return tb;
        }
        public static bool IsNullable(Type t)
        {
            return !t.IsValueType || (t.IsGenericType && t.GetGenericTypeDefinition() == typeof(Nullable<>));
        }

        public static Type GetCoreType(Type t)
        {
            if (t != null && IsNullable(t))
            {
                if (!t.IsValueType)
                {
                    return t;
                }
                else
                {
                    return Nullable.GetUnderlyingType(t);
                }
            }
            else
            {
                return t;
            }
        }

    }
}
