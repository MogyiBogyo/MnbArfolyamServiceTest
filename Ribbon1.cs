using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using mnbTask.MnbServiceReference;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.Xml;
using System.Data.OleDb;

namespace mnbTask
{

    
    public partial class Ribbon1
    {

        public void makeExcellfile()
        {
            Excel.Worksheet activeWS = Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Range actCell = Globals.ThisAddIn.Application.ActiveCell;


            ((Excel.Range)activeWS.Cells[1, 1]).Value2 = "Date";
            ((Excel.Range)activeWS.Cells[1, 2]).Value2 = "Currency";
            ((Excel.Range)activeWS.Cells[1, 3]).Value2 = "Unit";
            ((Excel.Range)activeWS.Cells[1, 4]).Value2 = "Value";
        }
           
        public void AccessConnection()
        {
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\a2c57g\source\repos\mnbArfolyamTask\MnbGet.accdb";

            String WindowsFelhasznNev = Environment.UserName;
            DateTime Idopont = DateTime.Now;
            String Indoklas = "";


            OleDbCommand cmd = new OleDbCommand("INSERT into timeStamps (WindowsFelhasznNev, Idopont) Values(@WindowsFelhasznNev, @Idopont)");
            cmd.Connection = conn;

            conn.Open();

            if (conn.State == System.Data.ConnectionState.Open)
            {
                cmd.Parameters.Add("@WindowsFelhasznNev", OleDbType.VarChar).Value = WindowsFelhasznNev;
                cmd.Parameters.Add("@Idopont", OleDbType.Date).Value = Idopont;

                //cmd.Parameters.Add("@Indoklas", OleDbType.LongVarChar).Value = null;

                try
                {
                    cmd.ExecuteNonQuery();
                    System.Windows.Forms.MessageBox.Show("Data Added");
                    conn.Close();
                }
                catch (OleDbException ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message);
                    conn.Close();
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Connection Failed");
            }
        }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void MnbGetter_Click(object sender, RibbonControlEventArgs e)
        {
            makeExcellfile();
            MNBArfolyamServiceSoapClient client = new MNBArfolyamServiceSoapClient();
            GetExchangeRatesRequestBody getExchangeRatesRequestBody = new GetExchangeRatesRequestBody();
            AccessConnection();


            int row = 2;
            int column = 1;
            int year = 2020;
            int month = 4;
            int day = 1;
            bool exit = false;
            bool first = true;

            while (!exit)
            {
                DateTime date = new DateTime(year, month, day);
                var firstDayOfMonth = new DateTime(date.Year, date.Month, 1);
                var lastDayOfMonth = new DateTime(date.Year, date.Month, DateTime.DaysInMonth(date.Year, date.Month));

                //if (first)
                //{
                //    getExchangeRatesRequestBody.startDate = "2020-04-01";
                //    getExchangeRatesRequestBody.endDate = "2020-04-01";
                //    first = false;
                //}
                //else
                //{
                //    getExchangeRatesRequestBody.startDate = firstDayOfMonth.ToString();
                //    getExchangeRatesRequestBody.endDate = lastDayOfMonth.ToString();
                //}

                getExchangeRatesRequestBody.startDate = "2020-01-12";
                getExchangeRatesRequestBody.endDate = "2020-01-15";
                GetCurrenciesRequestBody currbody = new GetCurrenciesRequestBody();
                var currencies = client.GetCurrencies(currbody);

                XmlDocument curreciesXml = new XmlDocument();
                curreciesXml.LoadXml(currencies.GetCurrenciesResult);

                List<string> currencieList = new List<string>();
                foreach (XmlNode item in curreciesXml.GetElementsByTagName("Curr"))
                {
                    currencieList.Add(item.InnerText);
                }
                getExchangeRatesRequestBody.currencyNames = string.Join(",", currencieList);
                var ExchangeswithDate = client.GetExchangeRates(getExchangeRatesRequestBody);

                //Xml feldolgozás
                Excel.Worksheet activeWS = Globals.ThisAddIn.Application.ActiveSheet;
                Excel.Range actCell = Globals.ThisAddIn.Application.ActiveCell;

                XmlDocument myXml = new XmlDocument();
                myXml.LoadXml(ExchangeswithDate.GetExchangeRatesResult);
                

                foreach (XmlElement item in myXml.GetElementsByTagName("Day"))
                {
                    XmlNodeList rates = item.GetElementsByTagName("Rate");
                   
                    
                    foreach (XmlElement rateElement in rates)
                    {

                        //string currs = rateElement.GetAttribute("curr");
                        string unit = rateElement.GetAttribute("unit");
                        string value = rateElement.InnerText;
                        ((Excel.Range)activeWS.Cells[row, column]).Value2 = item.GetAttribute("date");
                        
                        ((Excel.Range)activeWS.Cells[row, column + 1]).Value2 = rateElement.GetAttribute("curr");
                        ((Excel.Range)activeWS.Cells[row, column + 2]).Value2 = int.Parse(rateElement.GetAttribute("unit"));
                        ((Excel.Range)activeWS.Cells[row, column + 3]).Value2 = float.Parse(rateElement.InnerText);
                        //táblázat feltöltése
                        row++;

                    }
                }

                //System.Windows.Forms.MessageBox.Show(ExchangeswithDate.GetExchangeRatesResult);

                //Dátumcsökkentés
                //if (year <= 2015 && month == 1 && day == 1)
                //{
                //    exit = true;
                //}
                //else if (month == 1)
                //{
                //    year--;
                //    month = 12;

                //}
                //else
                //{
                //    month--;
                //}
                exit = true;
            }

            Excel.Worksheet activeWSforResize = Globals.ThisAddIn.Application.ActiveSheet;
            activeWSforResize.Columns.AutoFit();
            client.Close();

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {

        }
    }
}
