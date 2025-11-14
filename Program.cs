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
using Azure.Security.KeyVault.Secrets;
using Azure.Identity;

namespace DigitalOwl_Download
{
    class Program
    {
        private static string downloadDir;
        private static string CurrentBLine;
        private static string excelFile;
        private static string baseURL;
        private static string keyVaultUrl;
        private static string keyVaultSecretName;
        private static string ArchiveText = "ארכיון";
        private static string ArchivePortalText = "archived";
        private static string KEY = string.Empty;
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
            baseURL = ConfigurationManager.AppSettings["baseUrl"];
            keyVaultUrl = ConfigurationManager.AppSettings["keyVaultUrl"];
            keyVaultSecretName = ConfigurationManager.AppSettings["keyVaultSecretName"];

            if (string.IsNullOrEmpty(downloadDir) ||  string.IsNullOrEmpty(excelFile) || string.IsNullOrEmpty(CurrentBLine) ||
                string.IsNullOrEmpty(baseURL) || string.IsNullOrEmpty(keyVaultUrl) || string.IsNullOrEmpty(keyVaultSecretName))
            {
                throw new Exception("פרטי קונפיגורציה חסרים - Missing configuration details");
            }
            if (!Directory.Exists(downloadDir))
            {
                Directory.CreateDirectory(downloadDir);
            }

            KEY = await GetKeyFromAzureKeyVault();
            if (string.IsNullOrEmpty(KEY))
            {
                throw new Exception("בעיה בזמן נסיון לקבל את מחרוזת הרישיו - Failed to retrieve API key from Azure Key Vault");
            }

            var list = await GetCasesAsync();
            if(list.Count() == 0)
            {
                return;
            }
            SimpleLogger.SimpleLog.Info("looping cases");
            for (int i = 0; i < list.Count(); i++)
            {
                var item = list[i];
                SimpleLogger.SimpleLog.Info("downloading case - " + item.name);
                var download = await DownloadFile(item.id, item.name);
                if (download)
                {
                    SimpleLogger.SimpleLog.Info("download OK");
                    var row = WriteToExcel(item);
                    SimpleLogger.SimpleLog.Info("write to excel row - " + row);
                    if (row == -1)
                    {
                        throw new Exception("Failed to write to excel file");
                    }
                    
                    var fok = await ArchiveCase(item.name, item.id);
                    SimpleLogger.SimpleLog.Info("archiving result - "  + fok.ToString());
                    if (fok)
                    {
                        SimpleLogger.SimpleLog.Info("updating archive text for row - " + row);
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
                SimpleLogger.SimpleLog.Info("write to excel case - " + item.name);
                xlApp = new Excel.Application();
                xlApp.Visible = false;
                xlWorkbook = xlApp.Workbooks.Open(excelFile);
                xlWorksheet = (Excel._Worksheet)xlWorkbook.ActiveSheet;
                var lastRow = xlWorksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;
                bool found = false;
                var actualRow = lastRow;
                for (int i = lastRow; i > 1; i--)
                {
                    var name = xlWorksheet.Range[E_NAME + i, E_NAME + i].Value2.ToString().Trim();
                    var status = xlWorksheet.Range[E_STATUS + i, E_STATUS + i].Value2.ToString().Trim();
                    if (name == item.name.Trim() && status != ArchiveText)
                    {
                        SimpleLogger.SimpleLog.Info("case name found in excel - " + item.name);
                        found = true;
                        xlWorksheet.Range[E_DATEDOWNLOAD + i, E_DATEDOWNLOAD + i].Value2 = FormatExcelDate(DateTime.Now);
                        xlWorksheet.Range[E_NUMPAGES + i, E_NUMPAGES + i].Value2 = item.stats.pages.ToString();
                        xlWorksheet.Range[E_MEDICALDATA + i, E_MEDICALDATA + i].Value2 = item.stats.mediaData.ToString();
                        xlWorksheet.Range[E_HANDWRITTEN + i, E_HANDWRITTEN + i].Value2 = item.stats.handWritten.ToString();
                        xlWorksheet.Range[E_OWLSTATUS + i, E_OWLSTATUS + i].Value2 = item.stats.status.ToString();
                        xlWorksheet.Range[E_STATUS + i, E_STATUS + i].Value2 = "הורדה";
                        actualRow = i;
                    }
                }
                if (!found)
                {
                    SimpleLogger.SimpleLog.Info("case name NOT found in excel - " + item.name);
                    var row = lastRow + 1;
                    actualRow = row;
                    xlWorksheet.Range[E_DATE + row, E_DATE + row].Value2 = FormatExcelDate(DateTime.Now);
                    xlWorksheet.Range[E_NAME + row, E_NAME + row].Value2 = item.name.Trim();
                    xlWorksheet.Range[E_STATUS + row, E_STATUS + row].Value2 = "לא קיים באקסל";
                    xlWorksheet.Range[E_NUMPAGES + row, E_NUMPAGES + row].Value2 = item.stats.pages.ToString();
                    xlWorksheet.Range[E_MEDICALDATA + row, E_MEDICALDATA + row].Value2 = item.stats.mediaData.ToString();
                    xlWorksheet.Range[E_HANDWRITTEN + row, E_HANDWRITTEN + row].Value2 = item.stats.handWritten.ToString();
                    xlWorksheet.Range[E_OWLSTATUS + row, E_OWLSTATUS + row].Value2 = item.stats.status.ToString();
                    xlWorksheet.Range[E_BLINE + row, E_BLINE + row].Value2 = item.bline.ToString().Trim();
                    xlWorksheet.Range[E_DATEDOWNLOAD + row, E_DATEDOWNLOAD + row].Value2 = FormatExcelDate(DateTime.Now);
                }

                SimpleLogger.SimpleLog.Info("write to excel sccessfuy - closing excel");
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
                    var name = xlWorksheet.Range[E_NAME + i, E_NAME + i].Value2.ToString().Trim();
                    var status = xlWorksheet.Range[E_STATUS + i, E_STATUS + i].Value2.ToString().Trim();
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
                        RequestUri = new Uri(baseURL + "/cases/" + id + "/summary"),
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
                        RequestUri = new Uri(baseURL + "/cases"),
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
                                .Where(r => (string)r["externalStatus"] == "completed")
                                .Select(s => new Completed { id = (string)s["id"], name = (string)s["name"], bline = getBlineFromList((string)s["businessLineId"], lst) })
                                .ToList();


                        cases.AddRange(query);
                        //foreach (KeyValuePair<string, string> entry in lst)
                        //{
                        //    var bLineID = entry.Value;
                        //    var query = root
                        //        .Where(r => (string)r["businessLineId"] == bLineID && (string)r["externalStatus"] == "completed")
                        //        .Select(s => new Completed { id = (string)s["id"], name = (string)s["name"], bline = entry.Key })
                        //        .ToList();


                        //    cases.AddRange(query);
                        //}
                    }
                }
                SimpleLogger.SimpleLog.Info("Number of cases found - " + cases.Count());
                for (int i = 0; i < cases.Count(); i++)
                {
                    
                    var completedCase = cases[i];

                    SimpleLogger.SimpleLog.Info("case number " + (i + 1) + " - " + completedCase.name);
                    using (var client = new HttpClient())
                    {
                        try
                        {
                            var request = new HttpRequestMessage()
                            {
                                RequestUri = new Uri(baseURL + "/cases/" + completedCase.id + "/statistics"),
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

        private static string getBlineFromList(string blId, Dictionary<string, string> lst)
        {
            foreach (KeyValuePair<string, string> entry in lst)
            {
                var bLineID = entry.Value;
                var bLineName = entry.Key;
                if (bLineID == blId)
                {
                    return bLineName;
                }
            }
            return blId;
        }
        private static async Task<string> GetBLineIdFromOwl(string bline)
        {
            try
            {
                using (var client = new HttpClient())
                {
                    var request = new HttpRequestMessage()
                    {
                        RequestUri = new Uri(baseURL + "/businessLines"),
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
                SimpleLogger.SimpleLog.Info("archiving case - " + name);
                using (var client = new HttpClient())
                {
                    var request = new HttpRequestMessage()
                    {
                        RequestUri = new Uri(baseURL + "/cases/" + caseId + "/archive"),
                        Method = HttpMethod.Put,

                    };
                    client.DefaultRequestHeaders.Add("Authorization", "Bearer " + KEY);

                    using (var response = await client.SendAsync(request))
                    {
                        var msg = await response.Content.ReadAsStringAsync();
                        SimpleLogger.SimpleLog.Info("response body");
                        SimpleLogger.SimpleLog.Log(msg);
                        response.EnsureSuccessStatusCode();
                    }
                }
                return true;

            }
            catch (Exception ex)
            {

                SimpleLogger.SimpleLog.Info("Error while archiving a case. Case ID - " + name);
                SimpleLogger.SimpleLog.Info("URL - " + baseURL + "/cases/" + caseId + "/archive");
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


        /// <summary>
        /// Retrieves the API key from Azure Key Vault using managed identity or Azure CLI authentication
        /// </summary>
        /// <returns>The API key as a string</returns>
        private static async Task<string> GetKeyFromAzureKeyVault()
        {
            try
            {
                SimpleLogger.SimpleLog.Info("Attempting to retrieve API key from Azure Key Vault");

                // Create a SecretClient using DefaultAzureCredential
                // This will attempt authentication in the following order:
                // 1. Environment variables (AZURE_CLIENT_ID, AZURE_TENANT_ID, AZURE_CLIENT_SECRET)
                // 2. Managed Identity (when running in Azure)
                // 3. Visual Studio authentication
                // 4. Azure CLI authentication
                // 5. Azure PowerShell authentication
                var credential = new DefaultAzureCredential();
                var client = new SecretClient(new Uri(keyVaultUrl), credential);

                // Retrieve the secret
                KeyVaultSecret secret = await client.GetSecretAsync(keyVaultSecretName);

                if (secret == null || string.IsNullOrEmpty(secret.Value))
                {
                    SimpleLogger.SimpleLog.Error($"Secret '{keyVaultSecretName}' retrieved but value is empty");
                    return string.Empty;
                }

                SimpleLogger.SimpleLog.Info("Successfully retrieved API key from Azure Key Vault");
                return secret.Value;
            }
            catch (Azure.RequestFailedException ex)
            {
                SimpleLogger.SimpleLog.Error($"Azure Key Vault request failed: {ex.Status} - {ex.Message}");
                SimpleLogger.SimpleLog.Log(ex);

                // Provide helpful error messages
                if (ex.Status == 401 || ex.Status == 403)
                {
                    SimpleLogger.SimpleLog.Error("Authentication/Authorization failed. Ensure the application has proper access to Key Vault.");
                    SimpleLogger.SimpleLog.Error("Required permissions: Get secrets from Key Vault access policy or RBAC role 'Key Vault Secrets User'");
                }
                else if (ex.Status == 404)
                {
                    SimpleLogger.SimpleLog.Error($"Secret '{keyVaultSecretName}' not found in Key Vault '{keyVaultUrl}'");
                }

                return string.Empty;
            }
            catch (Exception ex)
            {
                SimpleLogger.SimpleLog.Error("Unexpected error while retrieving API key from Azure Key Vault");
                SimpleLogger.SimpleLog.Log(ex);
                return string.Empty;
            }
        }

        #region Legacy GetKey Method - Deprecated
        /// <summary>
        /// [DEPRECATED] Legacy method to get key from Word document - replaced by Azure Key Vault
        /// Kept for reference only - DO NOT USE
        /// </summary>
        [Obsolete("This method is deprecated. Use GetKeyFromAzureKeyVault() instead.")]
        private static string GetKeyFromWordDocument(string keyFilePath)
        {
            Microsoft.Office.Interop.Word.Application word = null;
            Microsoft.Office.Interop.Word.Document doc = null;
            try
            {
                string key = string.Empty;
                word = new Microsoft.Office.Interop.Word.Application();
                doc = word.Documents.Open(keyFilePath);
                foreach (Microsoft.Office.Interop.Word.Paragraph objParagraph in doc.Paragraphs)
                {
                    key = objParagraph.Range.Text.Trim();
                }
                doc.Close();
                word.Quit();
                return key;
            }
            catch (Exception ex)
            {
                SimpleLogger.SimpleLog.Info("Error while trying to get the license key from file - " + ex.Message);
                SimpleLogger.SimpleLog.Log(ex);
                return string.Empty;
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();

                if (doc != null)
                {
                    Marshal.ReleaseComObject(doc);
                }

                if (word != null)
                {
                    Marshal.ReleaseComObject(word);
                }
            }
        }
        #endregion


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
