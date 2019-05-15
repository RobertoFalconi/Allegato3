using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Newtonsoft.Json;

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

            }
            catch (Exception ex)
            {
                Globals.Ribbons.Allegato3Ribbon.button10.Label = "Impossibile connettersi: " + ex.Message.ToString();
            }
        }

        private void Button1_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet currentSheet = Globals.ThisAddIn.GetActiveWorkSheet();

            Dictionary<string, List<string>> ExcelData = new Dictionary<string, List<string>>();

            #region chiavi
            List<string> chiavi = new List<string>();
            foreach (string chiave in currentSheet.Range["A1:AA1"].Value2) // TODO: il limite AA1 deve essere variabile
            {
                chiavi.Add(chiave);
            }
            List<string> lettere = new List<string>
                { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
            #endregion chiavi

            #region valori
            int nChiave = 0;
            foreach (string chiave in chiavi)
            {
                List<string> valori = new List<string>();
                int i = 2;
                string lettera;
                if ((nChiave / 26 - 1) >= 0)
                {
                    lettera = lettere[(nChiave / 26) - 1] + lettere[nChiave % 26];
                }
                else
                {
                    lettera = lettere[nChiave % 26];
                }
                foreach (string valore in currentSheet.Range[lettera + "2:" + lettera + "15"].Value2) // TODO: il limite 15 deve essere variabile
                {
                    string cella = lettera + i.ToString();
                    valori.Add(currentSheet.Range[cella].Value2);
                    i += 1;
                }
                ExcelData.Add(chiave, valori);
                nChiave += 1;
            }
            #endregion valori

            // TODO: salvare la cella in modifica

            HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create("http://localhost:50884/api/values");
            httpWebRequest.ContentType = "application/json";
            httpWebRequest.Method = "POST";

            using (StreamWriter streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
            {
                string json = JsonConvert.SerializeObject(ExcelData);

                streamWriter.Write(json);
                streamWriter.Flush();
                streamWriter.Close();
            }

            WebResponse resp = httpWebRequest.GetResponse();

            //HttpWebResponse resp = (HttpWebResponse)httpWebRequest.GetResponse();

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
            Worksheet currentSheet = Globals.ThisAddIn.GetActiveWorkSheet();

            string json = "{\"Chiave1\":[\"ValoreA1\",\"ValoreA2\",\"ValoreA3\",\"ValoreA4\",\"ValoreA5\",\"ValoreA6\",\"ValoreA7\",\"ValoreA8\",\"ValoreA9\",\"ValoreA10\",\"ValoreA11\",\"ValoreA12\",\"ValoreA13\",\"ValoreA14\"],\"Chiave2\":[\"ValoreB1\",\"ValoreB2\",\"ValoreB3\",\"ValoreB4\",\"ValoreB5\",\"ValoreB6\",\"ValoreB7\",\"ValoreB8\",\"ValoreB9\",\"ValoreB10\",\"ValoreB11\",\"ValoreB12\",\"ValoreB13\",\"ValoreB14\"],\"Chiave3\":[\"ValoreC1\",\"ValoreC2\",\"ValoreC3\",\"ValoreC4\",\"ValoreC5\",\"ValoreC6\",\"ValoreC7\",\"ValoreC8\",\"ValoreC9\",\"ValoreC10\",\"ValoreC11\",\"ValoreC12\",\"ValoreC13\",\"ValoreC14\"],\"Chiave4\":[\"ValoreD1\",\"ValoreD2\",\"ValoreD3\",\"ValoreD4\",\"ValoreD5\",\"ValoreD6\",\"ValoreD7\",\"ValoreD8\",\"ValoreD9\",\"ValoreD10\",\"ValoreD11\",\"ValoreD12\",\"ValoreD13\",\"ValoreD14\"],\"Chiave5\":[\"ValoreE1\",\"ValoreE2\",\"ValoreE3\",\"ValoreE4\",\"ValoreE5\",\"ValoreE6\",\"ValoreE7\",\"ValoreE8\",\"ValoreE9\",\"ValoreE10\",\"ValoreE11\",\"ValoreE12\",\"ValoreE13\",\"ValoreE14\"],\"Chiave6\":[\"ValoreF1\",\"ValoreF2\",\"ValoreF3\",\"ValoreF4\",\"ValoreF5\",\"ValoreF6\",\"ValoreF7\",\"ValoreF8\",\"ValoreF9\",\"ValoreF10\",\"ValoreF11\",\"ValoreF12\",\"ValoreF13\",\"ValoreF14\"],\"Chiave7\":[\"ValoreG1\",\"ValoreG2\",\"ValoreG3\",\"ValoreG4\",\"ValoreG5\",\"ValoreG6\",\"ValoreG7\",\"ValoreG8\",\"ValoreG9\",\"ValoreG10\",\"ValoreG11\",\"ValoreG12\",\"ValoreG13\",\"ValoreG14\"],\"Chiave8\":[\"ValoreH1\",\"ValoreH2\",\"ValoreH3\",\"ValoreH4\",\"ValoreH5\",\"ValoreH6\",\"ValoreH7\",\"ValoreH8\",\"ValoreH9\",\"ValoreH10\",\"ValoreH11\",\"ValoreH12\",\"ValoreH13\",\"ValoreH14\"],\"Chiave9\":[\"ValoreI1\",\"ValoreI2\",\"ValoreI3\",\"ValoreI4\",\"ValoreI5\",\"ValoreI6\",\"ValoreI7\",\"ValoreI8\",\"ValoreI9\",\"ValoreI10\",\"ValoreI11\",\"ValoreI12\",\"ValoreI13\",\"ValoreI14\"],\"Chiave10\":[\"ValoreJ1\",\"ValoreJ2\",\"ValoreJ3\",\"ValoreJ4\",\"ValoreJ5\",\"ValoreJ6\",\"ValoreJ7\",\"ValoreJ8\",\"ValoreJ9\",\"ValoreJ10\",\"ValoreJ11\",\"ValoreJ12\",\"ValoreJ13\",\"ValoreJ14\"],\"Chiave11\":[\"ValoreK1\",\"ValoreK2\",\"ValoreK3\",\"ValoreK4\",\"ValoreK5\",\"ValoreK6\",\"ValoreK7\",\"ValoreK8\",\"ValoreK9\",\"ValoreK10\",\"ValoreK11\",\"ValoreK12\",\"ValoreK13\",\"ValoreK14\"],\"Chiave12\":[\"ValoreL1\",\"ValoreL2\",\"ValoreL3\",\"ValoreL4\",\"ValoreL5\",\"ValoreL6\",\"ValoreL7\",\"ValoreL8\",\"ValoreL9\",\"ValoreL10\",\"ValoreL11\",\"ValoreL12\",\"ValoreL13\",\"ValoreL14\"],\"Chiave13\":[\"ValoreM1\",\"ValoreM2\",\"ValoreM3\",\"ValoreM4\",\"ValoreM5\",\"ValoreM6\",\"ValoreM7\",\"ValoreM8\",\"ValoreM9\",\"ValoreM10\",\"ValoreM11\",\"ValoreM12\",\"ValoreM13\",\"ValoreM14\"],\"Chiave14\":[\"ValoreN1\",\"ValoreN2\",\"ValoreN3\",\"ValoreN4\",\"ValoreN5\",\"ValoreN6\",\"ValoreN7\",\"ValoreN8\",\"ValoreN9\",\"ValoreN10\",\"ValoreN11\",\"ValoreN12\",\"ValoreN13\",\"ValoreN14\"],\"Chiave15\":[\"ValoreO1\",\"ValoreO2\",\"ValoreO3\",\"ValoreO4\",\"ValoreO5\",\"ValoreO6\",\"ValoreO7\",\"ValoreO8\",\"ValoreO9\",\"ValoreO10\",\"ValoreO11\",\"ValoreO12\",\"ValoreO13\",\"ValoreO14\"],\"Chiave16\":[\"ValoreP1\",\"ValoreP2\",\"ValoreP3\",\"ValoreP4\",\"ValoreP5\",\"ValoreP6\",\"ValoreP7\",\"ValoreP8\",\"ValoreP9\",\"ValoreP10\",\"ValoreP11\",\"ValoreP12\",\"ValoreP13\",\"ValoreP14\"],\"Chiave17\":[\"ValoreQ1\",\"ValoreQ2\",\"ValoreQ3\",\"ValoreQ4\",\"ValoreQ5\",\"ValoreQ6\",\"ValoreQ7\",\"ValoreQ8\",\"ValoreQ9\",\"ValoreQ10\",\"ValoreQ11\",\"ValoreQ12\",\"ValoreQ13\",\"ValoreQ14\"],\"Chiave18\":[\"ValoreR1\",\"ValoreR2\",\"ValoreR3\",\"ValoreR4\",\"ValoreR5\",\"ValoreR6\",\"ValoreR7\",\"ValoreR8\",\"ValoreR9\",\"ValoreR10\",\"ValoreR11\",\"ValoreR12\",\"ValoreR13\",\"ValoreR14\"],\"Chiave19\":[\"ValoreS1\",\"ValoreS2\",\"ValoreS3\",\"ValoreS4\",\"ValoreS5\",\"ValoreS6\",\"ValoreS7\",\"ValoreS8\",\"ValoreS9\",\"ValoreS10\",\"ValoreS11\",\"ValoreS12\",\"ValoreS13\",\"ValoreS14\"],\"Chiave20\":[\"ValoreT1\",\"ValoreT2\",\"ValoreT3\",\"ValoreT4\",\"ValoreT5\",\"ValoreT6\",\"ValoreT7\",\"ValoreT8\",\"ValoreT9\",\"ValoreT10\",\"ValoreT11\",\"ValoreT12\",\"ValoreT13\",\"ValoreT14\"],\"Chiave21\":[\"ValoreU1\",\"ValoreU2\",\"ValoreU3\",\"ValoreU4\",\"ValoreU5\",\"ValoreU6\",\"ValoreU7\",\"ValoreU8\",\"ValoreU9\",\"ValoreU10\",\"ValoreU11\",\"ValoreU12\",\"ValoreU13\",\"ValoreU14\"],\"Chiave22\":[\"ValoreV1\",\"ValoreV2\",\"ValoreV3\",\"ValoreV4\",\"ValoreV5\",\"ValoreV6\",\"ValoreV7\",\"ValoreV8\",\"ValoreV9\",\"ValoreV10\",\"ValoreV11\",\"ValoreV12\",\"ValoreV13\",\"ValoreV14\"],\"Chiave23\":[\"ValoreW1\",\"ValoreW2\",\"ValoreW3\",\"ValoreW4\",\"ValoreW5\",\"ValoreW6\",\"ValoreW7\",\"ValoreW8\",\"ValoreW9\",\"ValoreW10\",\"ValoreW11\",\"ValoreW12\",\"ValoreW13\",\"ValoreW14\"],\"Chiave24\":[\"ValoreX1\",\"ValoreX2\",\"ValoreX3\",\"ValoreX4\",\"ValoreX5\",\"ValoreX6\",\"ValoreX7\",\"ValoreX8\",\"ValoreX9\",\"ValoreX10\",\"ValoreX11\",\"ValoreX12\",\"ValoreX13\",\"ValoreX14\"],\"Chiave25\":[\"ValoreY1\",\"ValoreY2\",\"ValoreY3\",\"ValoreY4\",\"ValoreY5\",\"ValoreY6\",\"ValoreY7\",\"ValoreY8\",\"ValoreY9\",\"ValoreY10\",\"ValoreY11\",\"ValoreY12\",\"ValoreY13\",\"ValoreY14\"],\"Chiave26\":[\"ValoreZ1\",\"ValoreZ2\",\"ValoreZ3\",\"ValoreZ4\",\"ValoreZ5\",\"ValoreZ6\",\"ValoreZ7\",\"ValoreZ8\",\"ValoreZ9\",\"ValoreZ10\",\"ValoreZ11\",\"ValoreZ12\",\"ValoreZ13\",\"ValoreZ14\"],\"Chiave27\":[\"ValoreAA1\",\"ValoreAA2\",\"ValoreAA3\",\"ValoreAA4\",\"ValoreAA5\",\"ValoreAA6\",\"ValoreAA7\",\"ValoreAA8\",\"ValoreAA9\",\"ValoreAA10\",\"ValoreAA11\",\"ValoreAA12\",\"ValoreAA13\",\"ValoreAA14\"]}";
            Dictionary<string, List<string>> jObj = JsonConvert.DeserializeObject<Dictionary<string, List<string>>>(json);
            List<string> lettere = new List<string>
                { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
            int i = 0;
            string lettera;
            foreach (string chiave in jObj.Keys)
            {
                int j = 2;
                if (i < 26)
                {
                    lettera = lettere[i].ToString();
                    currentSheet.Range[lettera + "1"].Value2 = chiave;
                }
                else
                {
                    lettera = lettere[i / 26 - 1].ToString() + lettere[i % 26].ToString();
                    currentSheet.Range[lettera + "1"].Value2 = chiave;
                }
                foreach (string valore in jObj[chiave])
                {
                    currentSheet.Range[lettera + j.ToString()].Value2 = valore;
                    j += 1;
                }
                i += 1;
            }
            currentSheet.Columns.AutoFit();
        }

    }
}
