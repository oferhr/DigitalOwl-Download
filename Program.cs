using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace DigitalOwl_Download
{
    class Program
    {
        private static string downloadDir;
        private static string CurrentBLine;
        private static string excelFile;
        private static string KEY = "eyJjbGllbnRfaWQiOiI0dzFPNUlJTE9GajdSajhvckZqTkJvR3Z4RVkwNlhUQyIsImNsaWVudF9zZWNyZXQiOiJsYWNBcFd0VTFHeXRfSVNlVGZCZVdweGRBRVJ3NG94Zm9EWkNvZmw0NjI2N3p1Q3ZSRUFTRjdpSEFDWDRnSmIzIiwiYXVkaWVuY2UiOiJodHRwczovL2FwaS5kaWdpdGFsb3dsLmFwcCIsImdyYW50X3R5cGUiOiJjbGllbnRfY3JlZGVudGlhbHMifQ==";
        private static Dictionary<string, string> bLines = new Dictionary<string, string>
        {
            {"defBlName", "914aa316-2243-4efb-aeea-a61758772b38" },
            {"defBlTemp", "9ff4ab50-58ee-4f3e-9c95-7479c6e02529"}
        };
        
        static async Task Main(string[] args)
        {
            downloadDir = ConfigurationManager.AppSettings["downloadDir"];
            excelFile = ConfigurationManager.AppSettings["excelFile"];
            CurrentBLine = ConfigurationManager.AppSettings["buisnessLine"];

            if (string.IsNullOrEmpty(downloadDir) ||  string.IsNullOrEmpty(excelFile) || string.IsNullOrEmpty(CurrentBLine))
            {
                throw new Exception("פרטי קונפיגורציה חסרים");
            }
            if (!Directory.Exists(downloadDir))
            {
                Directory.CreateDirectory(downloadDir);
            }
            

            var list = await GetCasesAsync();
            if(list.Count() == 0)
            {
                return;
            }
            for (int i = 0; i < list.Count(); i++)
            {
                var item = list[i];
                var download = await DownloadFile(item.id, item.name);
                if (download)
                {
                    WriteToExcel(item);
                    await ArchiveCase(item.name, item.id);
                }
                
            }
        }

        private static void WriteToExcel(Completed item)
        {
            Excel._Worksheet xlWorksheet = null;
            Excel.Workbook xlWorkbook = null;
            Excel.Application xlApp = null;
            try
            {
                xlApp = new Excel.Application();
                xlApp.Visible = false;
                xlWorkbook = xlApp.Workbooks.Open(excelFile);
                xlWorksheet = (Excel._Worksheet)xlWorkbook.ActiveSheet;
                var lastRow = xlWorksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;
                bool found = false;
                for (int i = lastRow; i > 1; i--)
                {
                    var name = xlWorksheet.Range[E_NAME + i, E_NAME + i].Value2.ToString();
                    var status = xlWorksheet.Range[E_STATUS + i, E_STATUS + i].Value2.ToString();
                    if (name == item.name && status == "העלה")
                    {
                        found = true;
                        xlWorksheet.Range[E_DATEDOWNLOAD + i, E_DATEDOWNLOAD + i].Value2 = FormatExcelDate(DateTime.Now);
                        xlWorksheet.Range[E_NUMPAGES + i, E_NUMPAGES + i].Value2 = item.pages.ToString();
                        xlWorksheet.Range[E_STATUS + i, E_STATUS + i].Value2 = "הורדה";

                    }
                }
                if (!found)
                {
                    var row = lastRow + 1;
                    xlWorksheet.Range[E_DATE + row, E_DATE + row].Value2 = FormatExcelDate(DateTime.Now);
                    xlWorksheet.Range[E_NAME + row, E_NAME + row].Value2 = item.name;
                    xlWorksheet.Range[E_STATUS + row, E_STATUS + row].Value2 = "לא קיים באקסל";
                    xlWorksheet.Range[E_NUMPAGES + row, E_NUMPAGES + row].Value2 = item.pages.ToString();
                    xlWorksheet.Range[E_DATEDOWNLOAD + row, E_DATEDOWNLOAD + row].Value2 = FormatExcelDate(DateTime.Now);
                }


                xlWorkbook.Save();
                xlWorkbook.Close();
                xlApp.Quit();
                //

            }
            catch (Exception ex)
            {
                xlWorkbook.Close();
                xlApp.Quit();
                SimpleLogger.SimpleLog.Log(ex);
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();

                if (xlWorksheet != null)
                {
                    Marshal.ReleaseComObject(xlWorksheet);
                }
                //close and release
                if (xlWorkbook != null)
                {
                    Marshal.ReleaseComObject(xlWorkbook);
                }

                if (xlApp != null)
                {
                    //quit and release

                    Marshal.ReleaseComObject(xlApp);
                }
            }
        }
        static void ErrorToExcel(DirData data)
        {
            Excel._Worksheet xlWorksheet = null;
            Excel.Workbook xlWorkbook = null;
            Excel.Application xlApp = null;
            try
            {
                xlApp = new Excel.Application();
                xlApp.Visible = false;
                xlWorkbook = xlApp.Workbooks.Open(excelFile);
                xlWorksheet = (Excel._Worksheet)xlWorkbook.ActiveSheet;
                var lastRow = xlWorksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;

                for (int i = lastRow; i > 1; i--)
                {
                    var name = xlWorksheet.Range[E_NAME + i, E_NAME + i].Value2.ToString();
                    var status = xlWorksheet.Range[E_STATUS + i, E_STATUS + i].Value2.ToString();
                    if (name == data.name && status == data.type)
                    {
                        xlWorksheet.Range[E_REMARK + i, E_REMARK + i].Value2 = data.date;
                        xlWorksheet.Range[E_STATUS + i, E_STATUS + i].Value2 = data.status;
                    }
                }



                xlWorkbook.Save();
                xlWorkbook.Close();
                xlApp.Quit();
                //

            }
            catch (Exception ex)
            {
                xlWorkbook.Close();
                xlApp.Quit();
                SimpleLogger.SimpleLog.Log(ex);
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();

                if (xlWorksheet != null)
                {
                    Marshal.ReleaseComObject(xlWorksheet);
                }
                //close and release
                if (xlWorkbook != null)
                {
                    Marshal.ReleaseComObject(xlWorkbook);
                }

                if (xlApp != null)
                {
                    //quit and release

                    Marshal.ReleaseComObject(xlApp);
                }
            }
        }

        private static async Task<bool> DownloadFile(string id, string name)
        {
            try
            {
                using (var client = new HttpClient())
                {
                    var request = new HttpRequestMessage()
                    {
                        RequestUri = new Uri("https://api.digitalowl.app/cases/" + id + "/summary"),
                        Method = HttpMethod.Get

                    };
                    client.DefaultRequestHeaders.Accept.Clear();
                    client.DefaultRequestHeaders.Add("Authorization", "Bearer " + KEY);
                    client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
                    using (var response = await client.SendAsync(request))
                    {
                        response.EnsureSuccessStatusCode();
                        HttpContent content = response.Content;
                        var contentStream = await content.ReadAsStreamAsync();
                        using (var fs = new FileStream(Path.Combine(downloadDir, name + ".pdf"), FileMode.CreateNew))
                        {
                            await content.CopyToAsync(fs);
                        }
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                SimpleLogger.SimpleLog.Info("Error while downloading file");
                SimpleLogger.SimpleLog.Log(ex);
                BuildError(name, "Error while downloading file. - " + ex.Message);
                //return "ERROR";
                return false;
            }
        }
        private static async Task<List<Completed>> GetCasesAsync()
        {
            try
            {
                string bLineID;
                if (!bLines.TryGetValue(CurrentBLine, out bLineID))
                {
                    bLineID = await GetBLineID(CurrentBLine);
                }
                if (bLineID == null || bLineID == "ERROR")
                {
                    SimpleLogger.SimpleLog.Info("No BuisnessLine exist for provided value");
                    throw new Exception("No BuisnessLine exist for provided value");
                }
                using (var client = new HttpClient())
                {
                    var request = new HttpRequestMessage()
                    {
                        RequestUri = new Uri("https://api.digitalowl.app/cases"),
                        Method = HttpMethod.Get

                    };

                    client.DefaultRequestHeaders.Add("Authorization", "Bearer " + KEY);

                    using (var response = await client.SendAsync(request))
                    {
                        var data = await response.Content.ReadAsStringAsync();
                        var root = (JArray)JsonConvert.DeserializeObject(data);
                        if (root.Count == 0)
                        {
                            return null;
                        }
                        var query = root
                            .Where(r => (string)r["businessLineId"] == bLineID && (string)r["externalStatus"] == "completed")
                            .Select(s => new Completed { id = (string)s["id"], pages = (int)s["pageCount"], name = (string)s["name"] })
                            .ToList();


                        return query;
                    }
                }
            }
                    
            catch (Exception ex)
            {
                SimpleLogger.SimpleLog.Info("Error while checking for completed cases");
                SimpleLogger.SimpleLog.Log(ex);
                //return "ERROR";
                return new List<Completed>();
            }
        }
        private static async Task<string> GetBLineID(string bline)
        {
            try
            {
                using (var client = new HttpClient())
                {
                    var request = new HttpRequestMessage()
                    {
                        RequestUri = new Uri("https://api.digitalowl.app/businessLines"),
                        Method = HttpMethod.Get,

                    };
                    client.DefaultRequestHeaders.Add("Authorization", "Bearer " + KEY);

                    using (var response = await client.SendAsync(request))
                    {
                        var data = await response.Content.ReadAsStringAsync();
                        var oData = (JArray)JsonConvert.DeserializeObject(data);
                        if (oData.Count == 0)
                        {
                            return null;
                        }
                        var obj = oData.Children<JObject>().FirstOrDefault(f => f["name"] != null && f["name"].ToString() == bline);
                        if (obj != null && obj.Count > 0)
                        {
                            return obj["id"].ToString();

                        }
                        return null;
                    }
                }
                    
            }
            catch (Exception ex)
            {
                SimpleLogger.SimpleLog.Info("Error while checking id for buisness line ID - " + bline);
                SimpleLogger.SimpleLog.Log(ex);
                return "ERROR";
            }
        }
        private static async Task ArchiveCase(string name, string caseId)
        {
            try
            {
                using (var client = new HttpClient())
                {
                    var request = new HttpRequestMessage()
                    {
                        RequestUri = new Uri("https://api.digitalowl.app/cases/" + caseId + "/archive"),
                        Method = HttpMethod.Put,

                    };
                    client.DefaultRequestHeaders.Add("Authorization", "Bearer " + KEY);

                    using (var response = await client.SendAsync(request))
                    {
                        response.EnsureSuccessStatusCode();
                    }
                }

            }
            catch (Exception ex)
            {

                SimpleLogger.SimpleLog.Info("Error while archiving a case. Case ID - " + name);
                SimpleLogger.SimpleLog.Log(ex);
                BuildError(name, "Error while archiving a case. - " + ex.Message, "הורדה");
            }
        }
        private static void BuildError(string name, string msg, string type = "העלה")
        {
            var errorStatus = new DirData
            {
                remark = msg,
                name = name,
                status = "שגיאה",
                type = type
            };
            ErrorToExcel(errorStatus);
        }
        static string FormatExcelDate(DateTime dt)
        {
            return dt.ToString("dd/MM/yyyy H:mm:ss");
        }
        private static string E_DATE = "A";
        private static string E_NAME = "B";
        private static string E_NUMDOCS = "C";
        private static string E_NUMPAGES = "D";
        private static string E_STATUS = "E";
        private static string E_DATEUPLOAD = "F";
        private static string E_DATEDOWNLOAD = "G";
        private static string E_CONTINUE = "H";
        private static string E_REMARK = "I";
    }
    public  class Completed
    {
        public string id { get; set; }
        public int pages { get; set; }
        public string name { get; set; }
    }
    public class DirData
    {
        public string date { get; set; }
        public string name { get; set; }
        public string docs { get; set; }
        public string status { get; set; }
        public string remark { get; set; }

        public string type { get; set; }
    }
}
