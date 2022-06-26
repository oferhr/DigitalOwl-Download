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
        private static string ArchiveText = "ארכיון";
        private static string ArchivePortalText = "archived";
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
                    var row = WriteToExcel(item);
                    if (row == -1)
                    {
                        throw new Exception("Failed to write to excel file");
                    }
                    var fok = await ArchiveCase(item.name, item.id);
                    if (fok)
                    {
                        await UpdateExcelArchived(row);
                    }
                }
                
            }
        }

        static Task UpdateExcelArchived(int row)
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
                xlWorksheet.Range[E_STATUS + row, E_STATUS + row].Value2 = ArchiveText;
                xlWorksheet.Range[E_OWLSTATUS + row, E_OWLSTATUS + row].Value2 = ArchivePortalText;
                
                xlWorkbook.Save();
                xlWorkbook.Close();
                xlApp.Quit();
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

            return Task.CompletedTask;
        }

        private static int WriteToExcel(Completed item)
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
                var actualRow = lastRow;
                for (int i = lastRow; i > 1; i--)
                {
                    var name = xlWorksheet.Range[E_NAME + i, E_NAME + i].Value2.ToString();
                    var status = xlWorksheet.Range[E_STATUS + i, E_STATUS + i].Value2.ToString();
                    if (name == item.name && status != ArchiveText)
                    {
                        found = true;
                        xlWorksheet.Range[E_DATEDOWNLOAD + i, E_DATEDOWNLOAD + i].Value2 = FormatExcelDate(DateTime.Now);
                        xlWorksheet.Range[E_NUMPAGES + i, E_NUMPAGES + i].Value2 = item.stats.pages.ToString();
                        xlWorksheet.Range[E_MEDICALDATA + i, E_MEDICALDATA + i].Value2 = item.stats.mediaData.ToString();
                        xlWorksheet.Range[E_HANDWRITTEN + i, E_HANDWRITTEN + i].Value2 = item.stats.handWritten.ToString();
                        xlWorksheet.Range[E_OWLSTATUS + i, E_OWLSTATUS + i].Value2 = item.stats.status.ToString();
                        xlWorksheet.Range[E_STATUS + i, E_STATUS + i].Value2 = "הורדה";

                    }
                }
                if (!found)
                {
                    var row = lastRow + 1;
                    actualRow = row;
                    xlWorksheet.Range[E_DATE + row, E_DATE + row].Value2 = FormatExcelDate(DateTime.Now);
                    xlWorksheet.Range[E_NAME + row, E_NAME + row].Value2 = item.name;
                    xlWorksheet.Range[E_STATUS + row, E_STATUS + row].Value2 = "לא קיים באקסל";
                    xlWorksheet.Range[E_NUMPAGES + row, E_NUMPAGES + row].Value2 = item.stats.pages.ToString();
                    xlWorksheet.Range[E_MEDICALDATA + row, E_MEDICALDATA + row].Value2 = item.stats.mediaData.ToString();
                    xlWorksheet.Range[E_HANDWRITTEN + row, E_HANDWRITTEN + row].Value2 = item.stats.handWritten.ToString();
                    xlWorksheet.Range[E_OWLSTATUS + row, E_OWLSTATUS + row].Value2 = item.stats.status.ToString();
                    xlWorksheet.Range[E_BLINE + row, E_BLINE + row].Value2 = item.bline.ToString();
                    xlWorksheet.Range[E_DATEDOWNLOAD + row, E_DATEDOWNLOAD + row].Value2 = FormatExcelDate(DateTime.Now);
                }


                xlWorkbook.Save();
                xlWorkbook.Close();
                xlApp.Quit();
                return actualRow;

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
            return -1;
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
                        var fileName = Path.Combine(downloadDir, name + ".pdf");
                        if (File.Exists(fileName))
                        {
                            File.Delete(fileName);
                        }
                        var contentStream = await content.ReadAsStreamAsync();
                        using (var fs = new FileStream(fileName, FileMode.CreateNew))
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
            var cases = new List<Completed>();
            try
            {
                var lst = await GetAllBuisnessLines();


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
                        
                        foreach (KeyValuePair<string, string> entry in lst)
                        {
                            var bLineID = entry.Value;
                            var query = root
                                .Where(r => (string)r["businessLineId"] == bLineID && (string)r["externalStatus"] == "completed")
                                .Select(s => new Completed { id = (string)s["id"], name = (string)s["name"], bline = entry.Key })
                                .ToList();


                            cases.AddRange(query);
                        }
                    }
                }

                for (int i = 0; i < cases.Count(); i++)
                {
                    var completedCase = cases[i];
                    using (var client = new HttpClient())
                    {
                        try
                        {
                            var request = new HttpRequestMessage()
                            {
                                RequestUri = new Uri("https://api.digitalowl.app/cases/" + completedCase.id + "/statistics"),
                                Method = HttpMethod.Get

                            };

                            client.DefaultRequestHeaders.Add("Authorization", "Bearer " + KEY);

                            using (var response = await client.SendAsync(request))
                            {
                                response.EnsureSuccessStatusCode();
                                var data = await response.Content.ReadAsStringAsync();
                                var json = JObject.Parse(data);

                                var stats = new Stats
                                {
                                    pages = (int)json["totalPageCount"],
                                    mediaData = (int)json["pagesWithMedicalDataCount"],
                                    handWritten = (int)json["handWrittenPageCount"],
                                    status = (string)json["status"]
                                };


                                completedCase.stats = stats;
                            }
                        }
                        catch
                        {
                            completedCase.stats = new Stats
                            {
                                pages = -1,
                                mediaData = -1,
                                handWritten = -1,
                                status = ""
                            };
                        }

                    }
                }
                return cases;
            }

            catch (Exception ex)
            {
                SimpleLogger.SimpleLog.Info("Error while checking for completed cases");
                SimpleLogger.SimpleLog.Log(ex);
                //return "ERROR";
                return new List<Completed>();
            }
        }
        private static async Task<string> GetBLineIdFromOwl(string bline)
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
        private static async Task<bool> ArchiveCase(string name, string caseId)
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
                return true;

            }
            catch (Exception ex)
            {

                SimpleLogger.SimpleLog.Info("Error while archiving a case. Case ID - " + name);
                SimpleLogger.SimpleLog.Log(ex);
                BuildError(name, "Error while archiving a case. - " + ex.Message, "הורדה");
                return false;
            }
        }
        private static async Task<Dictionary<string,string>> GetAllBuisnessLines()
        {
            var lst = new Dictionary<string, string>();
            string bLineID;
            if (!bLines.TryGetValue(CurrentBLine, out bLineID))
            {
                SimpleLogger.SimpleLog.Info("Default BuisnessLine is corrupt or not set in Config file");
                throw new Exception("Default BuisnessLine is corrupt or not set in Config file");
            }
            lst.Add(CurrentBLine,bLineID);
            Excel._Worksheet xlWorksheet = null;
            Excel.Workbook xlWorkbook = null;
            Excel.Application xlApp = null;
            var excelFileName = Path.GetFileName(excelFile);
            var bLineExcelFile = excelFile.Replace(excelFileName, "BusinessLine.csv");
            try
            {
                xlApp = new Excel.Application();
                xlApp.Visible = false;
                xlWorkbook = xlApp.Workbooks.Open(bLineExcelFile);
                xlWorksheet = (Excel._Worksheet)xlWorkbook.ActiveSheet;
                var lastRow = xlWorksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;
                for (int i = 2; i <= lastRow; i++)
                {
                    var name = xlWorksheet.Range["A" + i, "A" + i].Value2.ToString();
                    var bline = xlWorksheet.Range["B" + i, "B" + i].Value2.ToString();
                    bLineID = await GetBLineIdFromOwl(bline);
                    lst.Add(name, bLineID);
                }
               //xlWorkbook.Save();
                xlWorkbook.Close();
                xlApp.Quit();
                return lst;
            }
            catch (Exception ex)
            {
                SimpleLogger.SimpleLog.Log(ex);
                xlWorkbook.Close();
                xlApp.Quit();

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
            return null;
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
        private static string E_MEDICALDATA = "E";
        private static string E_HANDWRITTEN = "F";
        private static string E_STATUS = "G";
        private static string E_OWLSTATUS = "H";
        private static string E_DATEUPLOAD = "I";
        private static string E_DATEDOWNLOAD = "J";
        private static string E_CONTINUE = "K";
        private static string E_BLINE = "L";
        private static string E_REMARK = "M";
    }
    public  class Completed
    {
        public string id { get; set; }
        public Stats stats { get; set; }
        public string name { get; set; }
        public string bline { get; set; }
    }
    public class Stats
    {
        public int pages { get; set; }
        public int mediaData { get; set; }
        public int handWritten { get; set; }
        public string status { get; set; }
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
