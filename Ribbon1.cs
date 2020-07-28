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
        public string filepath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase).Substring(6);
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
            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source="+ filepath + "/MnbGet.accdb";

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
                    //System.Windows.Forms.MessageBox.Show("Data Added");
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
            makeExcellfile(); //fejléc létrehozása
            MNBArfolyamServiceSoapClient client = new MNBArfolyamServiceSoapClient();
            GetExchangeRatesRequestBody getExchangeRatesRequestBody = new GetExchangeRatesRequestBody();
            AccessConnection(); //dbconnection létrehozása + log beszúrása


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

        //visszaadja a tábla adatait
        private void MakeLog_Click(object sender, RibbonControlEventArgs e)
        {
            OleDbConnection conn = new OleDbConnection();
            List<string> ListboxItems = new List<string>();
            OleDbDataReader reader = null;
            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source="+ filepath+ "/MnbGet.accdb";
            conn.Open();
            OleDbCommand cmd = new OleDbCommand("Select * from timeStamps", conn);

            string name, time, reason;
            int row=2, column=1;
            Excel.Workbook activeWB = Globals.ThisAddIn.Application.ActiveWorkbook;
            var xlSheets = activeWB.Sheets as Excel.Sheets;
            var LogSheet = (Excel.Worksheet)xlSheets.Add(xlSheets[1], Type.Missing, Type.Missing, Type.Missing);
            LogSheet.Name = "Logs";

            ((Excel.Range)LogSheet.Cells[1, 1]).Value2 = "Username";
            ((Excel.Range)LogSheet.Cells[1, 2]).Value2 = "Time";
            ((Excel.Range)LogSheet.Cells[1, 3]).Value2 = "Reason";



            reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                //ListboxItems.Add(reader[0].ToString() + "," + reader[1].ToString() + ","+ reader[2].ToString());
                //System.Windows.Forms.MessageBox.Show(reader[0].ToString() + "," + reader[1].ToString() + "," + reader[2].ToString());

                ((Excel.Range)LogSheet.Cells[row, column]).Value2 = reader[0];
                ((Excel.Range)LogSheet.Cells[row, column + 1]).Value2 = (DateTime)reader[1];
                ((Excel.Range)LogSheet.Cells[row, column + 2]).Value2 = reader[2];
                row++;
            }
            
            conn.Close();

        }

        private void logSave_Click(object sender, RibbonControlEventArgs e)
        {
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filepath + "/MnbGet.accdb";
            conn.Open();


            Excel.Worksheet currentSheet = Globals.ThisAddIn.Application.Worksheets["Logs"];
            int row=2, col=1;

            var WindowsFelhasznNev = ((Excel.Range)currentSheet.Cells[row, 1]).Value2;
            var Idopont = ((Excel.Range)currentSheet.Cells[row, 2]).Value2;
            var Reason = ((Excel.Range)currentSheet.Cells[row, 3]).Value2;

            bool exit = false;
            /*
            while (WindowsFelhasznNev != "" && !exit)
            {
                WindowsFelhasznNev = ((Excel.Range)currentSheet.Cells[row, 1]).Value2;
                Idopont = ((Excel.Range)currentSheet.Cells[row, 2]).Value2;
                Reason = ((Excel.Range)currentSheet.Cells[row, 3]).Value2;
                if (Reason != "" ){
                    try
                    {
                        OleDbCommand cmd = new OleDbCommand("Update timeStamps (Reason,WindowsFelhasznNev,Idopont) SET Reason = @Reason, WindowsFelhasznNev = @WindowsFelhasznNev, Idopont=@Idopont  WHERE WindowsFelhasznNev = @WindowsFelhasznNev AND Idopont = @Idopont ", conn);
                        cmd.Parameters.Add("@WindowsFelhasznNev", OleDbType.VarChar).Value = WindowsFelhasznNev;
                        cmd.Parameters.Add("@Idopont", OleDbType.Date).Value = Idopont;
                        cmd.Parameters.Add("@Reason", OleDbType.LongVarChar).Value = Reason;
                        cmd.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        System.Windows.Forms.MessageBox.Show(ex.Message);

                    }
                   
                }
                row++;
               
            }*/

            //ciklus 2. sortol, van-e valami az username-ben
            //while username !empty
            //reason !ures akkor update sql

        }
    }
}
