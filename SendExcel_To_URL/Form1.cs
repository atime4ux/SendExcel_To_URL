using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Net;
using System.Windows.Forms;

namespace SendExcel_To_URL
{
    public partial class Form1 : Form
    {
        libCommon.clsUtil objUtil = new libCommon.clsUtil();

        private System.Data.DataSet DS;
        private StringBuilder strBuilder = new StringBuilder();

        private int totalCnt;
        private int failCnt;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            totalCnt = 0;
            failCnt = 0;
            label3.Text = "";
            radioButton1.Checked = true;
            textBox3.Text = "http://";
        }

        //파일 선택 다이얼로그
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog objFD = new OpenFileDialog();

            objFD.Filter = "액셀파일|*.xls;*.xlsx|모든 파일|*.*";

            if (objFD.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                textBox1.Text = objFD.FileName;
            }
        }

        private void readExcelData()
        {
            DS = libMyUtil.clsExcel.readExcel(textBox1.Text);
            if (libMyUtil.clsCmnDB.validateDS(DS))
            {
                totalCnt = DS.Tables[0].Rows.Count;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            
            string Result=  "";
            string SendData;
            int i;

            readExcelData();
            
            if (totalCnt > 0)
            {
                for (i = 0; i < DS.Tables[0].Rows.Count; i++)
                {
                    SendData = makeSendData(i);
                    if (radioButton1.Checked)
                    {
                        //Post전송
                        Result = SendPostData(SendData);
                    }
                    else if (radioButton2.Checked)
                    {
                        //쿼리스트링 전송
                        Result = SendQueryString(SendData);
                    }
                    if (Result.Equals("FAIL"))
                    {
                        failCnt++;
                    }
                    strBuilder.AppendLine(i + 1 + " : " + Result);
                }

                label3.Text = string.Format("(전체 : {0}건, 실패 : {1}건)", totalCnt, failCnt);
                textBox2.AppendText(strBuilder.ToString());
            }
            else
            {
                MessageBox.Show("자료가 없습니다.");
            }
        }

        //POST방식으로 전송
        private string SendPostData(string SendData)
        {
            HttpWebRequest httpWebRequest;
            HttpWebResponse httpWebResponse;
            Stream requestStream;
            StreamReader streamReader;
            byte[] Data;
            string Result;

            Data = UTF8Encoding.UTF8.GetBytes(SendData);

            try
            {
                httpWebRequest = (HttpWebRequest)WebRequest.Create(textBox3.Text);
                httpWebRequest.ContentType = "application/x-www-form-urlencoded; charset=UTF-8";
                httpWebRequest.Method = "POST";
                httpWebRequest.ContentLength = Data.Length;

                requestStream = httpWebRequest.GetRequestStream();
                requestStream.Write(Data, 0, Data.Length);
                requestStream.Close();

                httpWebResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                streamReader = new StreamReader(httpWebResponse.GetResponseStream());

                Result = streamReader.ReadToEnd();
                streamReader.Close();
                httpWebResponse.Close();
            }
            catch (Exception ex)
            { 
                objUtil.writeLog("FAIL DOWNLOAD STRING : " + ex.ToString());
                Result = "FAIL";
            }

            //결과값 처리
            if (Encoding.Default.GetByteCount(Result) > 20)
            {
                Result = Result.Substring(0, 10);
            }

            return Result;
        }

        //쿼리스트링으로 전송
        private string SendQueryString(string SendData)
        {
            WebClient WC = new WebClient();
            string Result;

            try
            {
                Result = WC.DownloadString(textBox3.Text + "?" + SendData);
            }
            catch (Exception ex)
            {
                objUtil.writeLog("FAIL DOWNLOAD STRING : " + ex.ToString());
                Result = "FAIL";
            }
            
            //결과값 처리
            if (Encoding.Default.GetByteCount(Result) > 20)
            {
                Result = Result.Substring(0, 10);
            }

            return Result;
        }

        private string makeSendData(int RowNumber)
        {
            StringBuilder strBuilder = new StringBuilder();
            string Col;
            string Val;
            int j;

            for (j = 0; j < DS.Tables[0].Columns.Count; j++)
            {
                Col = DS.Tables[0].Columns[j].ColumnName;
                Val = DS.Tables[0].Rows[RowNumber][j].ToString();
                strBuilder.Append(Col + "=" + libMyUtil.clsWeb.encURL(Val, "UTF-8"));
                if (j < DS.Tables[0].Columns.Count - 1)
                {
                    strBuilder.Append("&");
                }
            }

            return strBuilder.ToString();
        }

        
    }
}
