using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System.Json;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Threading;
using ThreadedTimer = System.Threading.Timer;
using System.Runtime.InteropServices;

namespace Allegato3
{
    public partial class Allegato3Ribbon
    {
        private static readonly HttpClient httpClient = new HttpClient();

        private async void Allegato3Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            try
            {
                string result = "";
                HttpResponseMessage response = await httpClient.GetAsync("http://localhost:50884/api/values");
                if (response.IsSuccessStatusCode)
                {
                    result = await response.Content.ReadAsStringAsync();
                }

                Globals.Ribbons.Allegato3Ribbon.button10.Label = "Connessione riuscita";

                //string jsonString = "";
                //object jsonObject


            }
            catch (Exception ex)
            {
                Globals.Ribbons.Allegato3Ribbon.button10.Label = "Impossibile connettersi: " + ex.Message.ToString();
            }
        }

        private void Button1_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet currentSheet = Globals.ThisAddIn.GetActiveWorkSheet();

            Application oXL;
            Workbook oWB;
            Worksheet oSheet;
            oXL = (Application)Marshal.GetActiveObject("Excel.Application");
            oXL.Visible = true;
            oWB = oXL.ActiveWorkbook;
            dynamic docProps = oWB.CustomDocumentProperties;

            List<Dictionary<int, List<Tuple<dynamic, dynamic, dynamic>>>> fogliExcel = new List<Dictionary<int, List<Tuple<dynamic, dynamic, dynamic>>>>();
            Dictionary<int, List<Dictionary<int, List<Tuple<dynamic, dynamic, dynamic>>>>> fileExcel = new Dictionary<int, List<Dictionary<int, List<Tuple<dynamic, dynamic, dynamic>>>>>();

            HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create("http://localhost:50884/api/values");
            httpWebRequest.ContentType = "application/json";
            httpWebRequest.Method = "POST";
            using (StreamWriter streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
            {
                for (int s = 0; s < oWB.Sheets.Count; s++)
                {
                    fogliExcel.Add(new Dictionary<int, List<Tuple<dynamic, dynamic, dynamic>>>());
                    currentSheet = ((Worksheet)oXL.ActiveWorkbook.Sheets[s + 1]);
                    //Parallel.For(0, currentSheet.UsedRange.Columns.Count, i =>
                    for (int i = 0; i < currentSheet.UsedRange.Columns.Count; i++)
                    {
                        List<Tuple<dynamic, dynamic, dynamic>> lista = new List<Tuple<dynamic, dynamic, dynamic>>();
                        foreach (dynamic elem in currentSheet.UsedRange.Columns[i + 1, Type.Missing].Rows)
                        {
                            lista.Add(Tuple.Create<dynamic, dynamic, dynamic>(elem.Formula, elem.Interior.Color, elem.Font.Bold));
                        }
                        fogliExcel[s].Add(i, lista);
                    }
                }
                fileExcel.Add(0, fogliExcel);



                var jsonString = JsonConvert.SerializeObject(fileExcel);
                streamWriter.Write(jsonString);
                streamWriter.Flush();
                streamWriter.Close();
            }

            WebResponse resp = httpWebRequest.GetResponse();

            if (resp.Equals(HttpStatusCode.BadRequest))
            {
                label1.Label = "Salvataggio fallito";
                label1.ShowLabel = true;
                timer1.Enabled = true;
            }
            else
            {
                label1.Label = "Salvataggio effettuato";
                label1.ShowLabel = true;
                timer1.Enabled = true;
            }

            //using (var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse())
            //{
            //    using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
            //    {
            //        var result = streamReader.ReadToEnd();
            //    }
            //}
        }

        private void Button5_Click(object sender, RibbonControlEventArgs e)
        {

            group2.Visible = true;
            group5.Visible = true;

            Worksheet currentSheet = Globals.ThisAddIn.GetActiveWorkSheet();

            Application currentApp = Globals.ThisAddIn.Application;
            currentApp.ScreenUpdating = false;
            currentSheet.Name = button5.Label;

            Application oXL;
            Workbook oWB;
            Worksheet oSheet;
            oXL = (Application)Marshal.GetActiveObject("Excel.Application");
            oXL.Visible = true;
            oWB = oXL.ActiveWorkbook;
            dynamic docProps = oWB.CustomDocumentProperties;

            string jsonString = JsonString.jsonString;
            JObject jsonObject = (JObject)JsonConvert.DeserializeObject(jsonString);

            foreach (var fileExcel in jsonObject)
            {
                int s = 0;
                foreach (var foglioExcel in fileExcel.Value)
                {
                    int j = 1;
                    foreach (var colonna in foglioExcel)
                    {
                        for (int i = 0; i < colonna.Values().Count(); i++)
                        {
                            currentSheet.Cells[i + 1, j].Value = colonna.ElementAtOrDefault(0).ElementAtOrDefault(i).ElementAtOrDefault(0).ElementAtOrDefault(0); // il secondo è l'elm e il terzo l'item
                            if (!colonna.ElementAtOrDefault(0).ElementAtOrDefault(i).ElementAtOrDefault(1).FirstOrDefault().ToString().Equals("16777215")) currentSheet.Cells[i + 1, j].Interior.Color = colonna.ElementAtOrDefault(0).ElementAtOrDefault(i).ElementAtOrDefault(1).FirstOrDefault();
                            if (colonna.ElementAtOrDefault(0).ElementAtOrDefault(i).ElementAtOrDefault(2).FirstOrDefault().ToString().Equals("True")) currentSheet.Cells[i + 1, j].Font.Bold = true;
                        }
                        j++;
                    }
                    s++;
                    currentSheet.Columns.AutoFit();
                    currentSheet.Rows.AutoFit();
                    if (s < fileExcel.Value.Count())
                    {
                        Worksheet newWorksheet;
                        newWorksheet = (Worksheet)currentApp.Worksheets.Add(After:currentSheet);
                        currentSheet = newWorksheet;
                    }
                }

            }

            currentApp.ScreenUpdating = true;
        }

        private void ActiveWorkbook_NewSheet(object Sh)
        {
            throw new NotImplementedException();
        }

        public void Timer1_Tick(object sender, EventArgs e)
        {
            label1.ShowLabel = false;
        }

        private void Button3_Click(object sender, RibbonControlEventArgs e)
        {

        }
    }
}
