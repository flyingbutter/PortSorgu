using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Web;
using System.Runtime.InteropServices;
using Jurassic.Library;
using HtmlAgilityPack;
using Newtonsoft.Json;
using System.Collections;

namespace PortSorgu
{
    public partial class Form1 : Form
    {
        OleDbConnection Econ;
        DataTable Exceldt;
        string constr, Query;
        string filePath;
        string outputPath;
       // private BackgroundWorker _worker;
        string a;
        int s;
        int max;



        public Form1()
        {
            InitializeComponent();
          
        }

        private const int INTERNET_COOKIE_HTTPONLY = 0x00002000;
        private string html_response;
        private string CsrfKey;

        [DllImport("wininet.dll", SetLastError = true)]
        private static extern bool InternetGetCookieEx(
            string url,
            string cookieName,
            StringBuilder cookieData,
            ref int size,
            int flags,
            IntPtr pReserved);

        /// <summary>
        /// Returns cookie contents as a string
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        public static string GetCookie(string url)
        {
            int size = 512;
            StringBuilder sb = new StringBuilder(size);
            if (!InternetGetCookieEx(url, null, sb, ref size, INTERNET_COOKIE_HTTPONLY, IntPtr.Zero))
            {
                if (size < 0)
                {
                    return null;
                }
                sb = new StringBuilder(size);
                if (!InternetGetCookieEx(url, null, sb, ref size, INTERNET_COOKIE_HTTPONLY, IntPtr.Zero))
                {
                    return null;
                }
            }
            return sb.ToString();
        }
    

    private void ExcelConn(string FilePath)
        {

            constr = string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0 Xml;HDR=YES;""", FilePath);
            Econ = new OleDbConnection(constr);

        }
        public delegate void updatebar();

        private void UpdateProgress()
        {

            progressBar1.Value += 1;
            label1.Text = s+"/"+max;

        }
        private async void button2_Click(object sender, EventArgs e)
        {
            string directoryPath = Path.GetDirectoryName(filePath) + "\\temps";
            string fileName = Path.GetFileNameWithoutExtension(filePath);
            Directory.CreateDirectory(directoryPath);

            outputPath = directoryPath + "\\" + fileName + ".csv";

            Query = string.Format("Select * FROM [{0}]", comboBox1.SelectedValue);
            OleDbCommand Ecom = new OleDbCommand(Query, Econ);

            Econ.Open();

            DataSet ds = new DataSet();

            OleDbDataAdapter oda = new OleDbDataAdapter(Query, Econ);
            Econ.Close();
            oda.Fill(ds);

            Exceldt = ds.Tables[0];
            max = Exceldt.Rows.Count;
            string pstnNo="";
            string adslNo="";
            if (comboBox2.SelectedIndex>-1)
            {
                 pstnNo = comboBox2.SelectedItem.ToString();
            }

            if (comboBox3.SelectedIndex>-1)
            {
                adslNo = comboBox3.SelectedItem.ToString();
            }

            webBrowser1.Navigate("https://bayi.dsmart.com.tr/WEB/Pages/InternetTechFinder.aspx?menuId=-1&customerId=-1");
             a = webBrowser1.DocumentText;
            int startindex = 0;
            if (File.Exists(outputPath))
            {
                startindex = File.ReadLines(outputPath).Count();
              //  Console.WriteLine(startindex);

            }
            progressBar1.Maximum = Exceldt.Rows.Count;
            progressBar1.Value = startindex;
            try
            {
                await Task.Factory.StartNew(() => sorgula(pstnNo,adslNo, startindex));
            }
            catch (Exception r)
            {

                Console.WriteLine(r.StackTrace);
            }   

        }


        public void sorgula(string pstnNo,string adslNo,int startIndex)
        {
            

            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(a);

            HtmlNode bodyNode = doc.DocumentNode.SelectSingleNode("/html/body");

            var script = bodyNode.Descendants()
                                         .Where(n => n.Name == "script")
                                         .First().InnerText;

            // Return the data of spect and stringify it into a proper JSON object
            var engine = new Jurassic.ScriptEngine();
            var result = engine.Evaluate(script);

            CsrfKey = engine.GetGlobalValue("csrfKey").ToString();
            Console.WriteLine(CsrfKey);

            // var json = JSONObject.Stringify(engine, result);

            //Console.WriteLine(json);
            ArrayList cookies = new ArrayList();
            foreach (string cookie in GetCookie("https://bayi.dsmart.com.tr/Web/Pages/index.aspx").Split(';'))
            {
                cookies.Add(cookie);
            }
    
            for (s = startIndex; s < Exceldt.Rows.Count; s++)
            {
                progressBar1.Invoke(new updatebar(this.UpdateProgress));

                try
                {
                    DataRow row = Exceldt.Rows[s];
                    string pstn="";
                    string adsl="";

                    if (!String.IsNullOrEmpty(pstnNo))
                    {
                        pstn = row[pstnNo].ToString();

                    }

                    if (!String.IsNullOrEmpty(adslNo))
                    {
                        adsl = row[adslNo].ToString();

                    }

                    String Url = "https://bayi.dsmart.com.tr/Facade/DealerService.asmx/GetAdslInfrastructure";
                    HttpWebRequest request = (HttpWebRequest)WebRequest.Create(Url);
                    request.CookieContainer = new CookieContainer();
                    request.Referer = "https://bayi.dsmart.com.tr/WEB/Pages/InternetTechFinder.aspx?menuId=-1&customerId=-1";
                    request.Host = "bayi.dsmart.com.tr";
                    request.KeepAlive = true;
                    request.Method = "POST";
                    request.Headers.Add("CsrfKey", CsrfKey);
                    request.Accept = "*/*";
                    request.Headers.Add("X-Requested-With", "XMLHttpRequest");
                    request.Headers.Add("Origin", "https://bayi.dsmart.com.tr");
                    request.ContentType = "application/json; charset=UTF-8";

                    // 
                    foreach (string cookie in cookies)
                    {
                        Cookie mycook = new Cookie(cookie.Split('=')[0].Replace(" ", ""), cookie.Split('=')[1]) { Domain = "bayi.dsmart.com.tr" };
                        request.CookieContainer.Add(mycook);
                    }




                    var postData = "{\"bbkquery\":\"\",\"pstnquery\":\"" + pstn + "\",\"xdslquery\":\""+adsl+"\"}";

                    var data = Encoding.ASCII.GetBytes(postData);
                    request.ContentLength = data.Length;
                    using (var stream = request.GetRequestStream())
                    {
                        stream.Write(data, 0, data.Length);
                    }


                    HttpWebResponse response = (HttpWebResponse)request.GetResponse();

                    html_response = new StreamReader(response.GetResponseStream()).ReadToEnd();

                    RootObjectMain r = JsonConvert.DeserializeObject<RootObjectMain>(html_response);

                    RootObject t = JsonConvert.DeserializeObject<RootObject>(r.d.Substring(3));

                   // Console.WriteLine(t.CurrentIsp);
                    string isp = t.CurrentIsp;

                    using (StreamWriter w = File.AppendText(outputPath))
                    {
                        w.WriteLine(pstn + "," + isp);
                    }
                 //   Console.Write("  " + s);
                   
                }

                catch (Exception e)
                {
                    Console.WriteLine(e.StackTrace);
                }
            }
        }
        public class RootObjectMain
        {
            public string d { get; set; }
        }
        public class FoundTechList
        {
            public string MaxSpeed { get; set; }
            public bool NdslAvailable { get; set; }
            public string GreenBrown { get; set; }
            public int DistaneToSwitchboard { get; set; }
            public string Tech { get; set; }
            public bool PortAvailable { get; set; }
            public string Description { get; set; }
            public int Provider { get; set; }
            public string RawResult { get; set; }
        }

        public class InputQueryList
        {
            public int Provider { get; set; }
            public int InputType { get; set; }
            public string Value { get; set; }
        }

        public class RootObject
        {
            public List<FoundTechList> FoundTechList { get; set; }
            public string CurrentIsp { get; set; }
            public string SaleType { get; set; }
            public string SaleTypeExplanation { get; set; }
            public bool ForceNakedTransfer { get; set; }
            public bool ForceNakedAdsl { get; set; }
            public bool ForbidNakedAdsl { get; set; }
            public List<InputQueryList> InputQueryList { get; set; }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ExcelConn(filePath);


            Query = string.Format("Select * FROM [{0}]", comboBox1.SelectedValue);
            OleDbCommand Ecom = new OleDbCommand(Query, Econ);

            Econ.Open();

            DataSet ds = new DataSet();

            OleDbDataAdapter oda = new OleDbDataAdapter(Query, Econ);
            Econ.Close();
            oda.Fill(ds);
            Exceldt = ds.Tables[0];

            foreach (DataColumn column in Exceldt.Columns)
            {

                comboBox2.Items.Add(column.ColumnName);
                comboBox3.Items.Add(column.ColumnName);



            }
        }

       

        private void button1_Click(object sender, EventArgs e)
        {
            Stream myStream = null;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*.xlsx";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;
            
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if ((myStream = openFileDialog1.OpenFile()) != null)
                    {
                        using (myStream)
                        {
                            filePath = openFileDialog1.FileName;

                            ExcelConn(filePath);
                            Econ.Open();

                            var ExcelTables = Econ.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, new Object[] { null, null, null, "TABLE" });
                            comboBox1.DisplayMember = "TABLE_NAME";
                            comboBox1.ValueMember = "TABLE_NAME";
                            comboBox1.DataSource = ExcelTables;


                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }
            }
        }
    }
}
