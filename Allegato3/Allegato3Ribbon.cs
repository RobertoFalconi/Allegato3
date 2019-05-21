using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Runtime.InteropServices;

namespace Allegato3
{
    public partial class Allegato3Ribbon
    {
        private static readonly HttpClient httpClient = new HttpClient();

        private async void Allegato3Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            Globals.Ribbons.Allegato3Ribbon.button10.Label = "Connessione riuscita";
            //try
            //{
            //    string result = "";
            //    HttpResponseMessage response = await httpClient.GetAsync("http://localhost:50884/api/values");
            //    if (response.IsSuccessStatusCode)
            //    {
            //        result = await response.Content.ReadAsStringAsync();
            //    }

            //    Globals.Ribbons.Allegato3Ribbon.button10.Label = "Connessione riuscita";

            //}
            //catch (Exception ex)
            //{
            //    Globals.Ribbons.Allegato3Ribbon.button10.Label = "Impossibile connettersi: " + ex.Message.ToString();
            //}
        }

        private void Button1_Click(object sender, RibbonControlEventArgs e)
        {
            //    label1.Label = "Salvataggio effettuato";
            //    label1.ShowLabel = true;
            //    timer1.Enabled = true;
            //Worksheet currentSheet = Globals.ThisAddIn.GetActiveWorkSheet();

            //Application oXL;
            //Workbook oWB;
            //Worksheet oSheet;
            //oXL = (Application)Marshal.GetActiveObject("Excel.Application");
            //oXL.Visible = true;
            //oWB = oXL.ActiveWorkbook;
            //dynamic docProps = oWB.CustomDocumentProperties;

            ////string nomeFoglio = editBox1.Text; // TODO: non è detto che editBox1.Text sia compilato
            //string nomeFoglio = lblFoglioId.Label;

            //List<Dictionary<int, List<Tuple<dynamic, dynamic, dynamic>>>> fogliExcel = new List<Dictionary<int, List<Tuple<dynamic, dynamic, dynamic>>>>();
            //Dictionary<string, List<Dictionary<int, List<Tuple<dynamic, dynamic, dynamic>>>>> fileExcel = new Dictionary<string, List<Dictionary<int, List<Tuple<dynamic, dynamic, dynamic>>>>>();

            //HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create("http://localhost:50884/api/values");
            //httpWebRequest.ContentType = "application/json";
            //httpWebRequest.Method = "PUT";
            //using (StreamWriter streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
            //{
            //    for (int s = 0; s < oWB.Sheets.Count; s++)
            //    {
            //        fogliExcel.Add(new Dictionary<int, List<Tuple<dynamic, dynamic, dynamic>>>());
            //        currentSheet = ((Worksheet)oXL.ActiveWorkbook.Sheets[s + 1]);
            //        //Parallel.For(0, currentSheet.UsedRange.Columns.Count, i =>
            //        for (int i = 0; i < currentSheet.UsedRange.Columns.Count; i++)
            //        {
            //            List<Tuple<dynamic, dynamic, dynamic>> lista = new List<Tuple<dynamic, dynamic, dynamic>>();
            //            foreach (dynamic elem in currentSheet.UsedRange.Columns[i + 1, Type.Missing].Rows)
            //            {
            //                lista.Add(Tuple.Create<dynamic, dynamic, dynamic>(elem.Formula, elem.Interior.Color, elem.Font.Bold));
            //            }
            //            fogliExcel[s].Add(i, lista);
            //        }
            //    }
            //    fileExcel.Add(nomeFoglio, fogliExcel);



            //    var jsonString = JsonConvert.SerializeObject(fileExcel);
            //    streamWriter.Write(jsonString);
            //    streamWriter.Flush();
            //    streamWriter.Close();
            //}

            label1.Label = "Salvataggio effettuato";
            label1.ShowLabel = true;
            timer1.Enabled = true;

            //WebResponse resp = httpWebRequest.GetResponse();

            //if (resp.Equals(HttpStatusCode.BadRequest))
            //{
            //    label1.Label = "Salvataggio fallito";
            //    label1.ShowLabel = true;
            //    timer1.Enabled = true;
            //}
            //else
            //{
            //    label1.Label = "Salvataggio effettuato";
            //    label1.ShowLabel = true;
            //    timer1.Enabled = true;
            //}
        }

        private void Button5_Click(object sender, RibbonControlEventArgs e)
        {

            group2.Visible = true;
            group5.Visible = true;

            Worksheet currentSheet = Globals.ThisAddIn.GetActiveWorkSheet();

            Application currentApp = Globals.ThisAddIn.Application;
            currentApp.ScreenUpdating = false;

            Application oXL;
            Workbook oWB;
            Worksheet oSheet;
            oXL = (Application)Marshal.GetActiveObject("Excel.Application");
            oXL.Visible = true;
            oWB = oXL.ActiveWorkbook;
            dynamic docProps = oWB.CustomDocumentProperties;

            // TODO: implementare get jsonString
            string jsonString = JsonString.jsonString;
            JObject jsonObject = (JObject)JsonConvert.DeserializeObject(jsonString);

            lblFoglioId.Label = jsonObject.Properties().Select(p => p.Name).FirstOrDefault();

            var list1 = new List<string>(); // { "e", "s" };
            list1.Add("e");
            list1.Add("s");
            var flatList1 = string.Join(";", list1);
            var list2 = new List<string>(); // { "t", "t/c", "ext", "co", "ka", "c", "k" };
            list2.Add("t");
            list2.Add("t/c");
            list2.Add("ext");
            list2.Add("co");
            list2.Add("ka");
            list2.Add("c");
            list2.Add("k");
            var flatList2 = string.Join(";", list2);

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
                            var value = colonna.ElementAtOrDefault(0).ElementAtOrDefault(i).ElementAtOrDefault(0).ElementAtOrDefault(0); // il secondo è l'elem e il terzo l'item
                            currentSheet.Cells[i + 1, j].Value = value;
                            if (""+ value == "e" || ""+ value == "s")
                            {
                                var cell = currentSheet.Cells[i + 1, j];
                                cell.Validation.Delete();
                                cell.Validation.Add(
                                   XlDVType.xlValidateList,
                                   XlDVAlertStyle.xlValidAlertInformation,
                                   XlFormatConditionOperator.xlBetween,
                                   flatList1,
                                   Type.Missing);
                                cell.Validation.IgnoreBlank = true;
                                cell.Validation.InCellDropdown = true;
                            }
                            else if (""+ value == "t" || ""+ value == "t/c" || ""+ value == "ext" ||
                                ""+ value == "co" || ""+ value == "ka" || ""+ value == "c" || ""+ value == "k")
                            {
                                var cell = currentSheet.Cells[i + 1, j];
                                cell.Validation.Delete();
                                cell.Validation.Add(
                                   XlDVType.xlValidateList,
                                   XlDVAlertStyle.xlValidAlertInformation,
                                   XlFormatConditionOperator.xlBetween,
                                   flatList2,
                                   Type.Missing);
                                cell.Validation.IgnoreBlank = true;
                                cell.Validation.InCellDropdown = true;
                            }
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
                        newWorksheet = (Worksheet)currentApp.Worksheets.Add(After: currentSheet);
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
            timer1.Dispose();
        }

        private void Button3_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void EliminaTemplate(object sender, RibbonControlEventArgs e)
        {
            // TODO: implementare eliminazione di un template
        }

        private void Button5_Click_1(object sender, RibbonControlEventArgs e)
        {
            if (button2.Visible && editBox1.Visible)
            {
                button2.Visible = false;
                editBox1.Visible = false;
            }
            else
            {
                button2.Visible = true;
                editBox1.Visible = true;
            }
        }

        private void FolderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {

        }

        private void EditBox1_TextChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private void Button2_Click(object sender, RibbonControlEventArgs e)
        {
            group2.Visible = true;
            group5.Visible = true;

            if (button2.Visible && editBox1.Visible)
            {
                button2.Visible = false;
                editBox1.Visible = false;
            }
            else
            {
                button2.Visible = true;
                editBox1.Visible = true;
            }
            label1.Label = "Salvataggio effettuato";
            label1.ShowLabel = true;
            timer1.Enabled = true;
            //Worksheet currentSheet = Globals.ThisAddIn.GetActiveWorkSheet();

            //group2.Visible = true;
            //group5.Visible = true;

            //if (button2.Visible && editBox1.Visible)
            //{
            //    button2.Visible = false;
            //    editBox1.Visible = false;
            //}
            //else
            //{
            //    button2.Visible = true;
            //    editBox1.Visible = true;
            //}

            //string nomeFoglio = editBox1.Text;
            //lblFoglioId.Label = nomeFoglio;

            //Application oXL;
            //Workbook oWB;
            //Worksheet oSheet;
            //oXL = (Application)Marshal.GetActiveObject("Excel.Application");
            //oXL.Visible = true;
            //oWB = oXL.ActiveWorkbook;
            //dynamic docProps = oWB.CustomDocumentProperties;

            //List<Dictionary<int, List<Tuple<dynamic, dynamic, dynamic>>>> fogliExcel = new List<Dictionary<int, List<Tuple<dynamic, dynamic, dynamic>>>>();
            //Dictionary<string, List<Dictionary<int, List<Tuple<dynamic, dynamic, dynamic>>>>> fileExcel = new Dictionary<string, List<Dictionary<int, List<Tuple<dynamic, dynamic, dynamic>>>>>();

            //HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create("http://localhost:50884/api/values");
            //httpWebRequest.ContentType = "application/json";
            //httpWebRequest.Method = "POST";
            //using (StreamWriter streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
            //{
            //    for (int s = 0; s < oWB.Sheets.Count; s++)
            //    {
            //        fogliExcel.Add(new Dictionary<int, List<Tuple<dynamic, dynamic, dynamic>>>());
            //        currentSheet = ((Worksheet)oXL.ActiveWorkbook.Sheets[s + 1]);
            //        //Parallel.For(0, currentSheet.UsedRange.Columns.Count, i =>
            //        for (int i = 0; i < currentSheet.UsedRange.Columns.Count; i++)
            //        {
            //            List<Tuple<dynamic, dynamic, dynamic>> lista = new List<Tuple<dynamic, dynamic, dynamic>>();
            //            foreach (dynamic elem in currentSheet.UsedRange.Columns[i + 1, Type.Missing].Rows)
            //            {
            //                lista.Add(Tuple.Create<dynamic, dynamic, dynamic>(elem.Formula, elem.Interior.Color, elem.Font.Bold));
            //            }
            //            fogliExcel[s].Add(i, lista);
            //        }
            //    }
            //    fileExcel.Add(nomeFoglio, fogliExcel);



            //    var jsonString = JsonConvert.SerializeObject(fileExcel);
            //    streamWriter.Write(jsonString);
            //    streamWriter.Flush();
            //    streamWriter.Close();
            //}

            //WebResponse resp = httpWebRequest.GetResponse();

            //if (resp.Equals(HttpStatusCode.BadRequest))
            //{
            //    label1.Label = "Salvataggio fallito";
            //    label1.ShowLabel = true;
            //    timer1.Enabled = true;
            //}
            //else
            //{
            //    label1.Label = "Salvataggio effettuato";
            //    label1.ShowLabel = true;
            //    timer1.Enabled = true;
            //}

            //using (var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse())
            //{
            //    using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
            //    {
            //        var result = streamReader.ReadToEnd();
            //    }
            //}
        }
    }
}
