using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using JsonExcel.Models;
using Newtonsoft.Json;
using OfficeOpenXml;

namespace JsonExcel
{
    class Program
    { 
        static string url = "http://fund.eastmoney.com/Data/FundDataPortfolio_Interface.aspx?dt=14&mc=returnjson&ft=all&pn=5000&pi=1&sc=abbname&st=asc";
        static string outputFile = "DataStore/fundmanager.xlsx";

        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            try
            {
                string jsonResult = FetchJsonFromUrl(url);
                if (!string.IsNullOrEmpty(jsonResult))
                {
                    try
                    {
                        JsonStrToXlsx(jsonResult, outputFile);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"发生错误：{ex.Message}");
                    }
                  
                    Console.WriteLine("数据成功导出到Excel");
                }
                else
                {
                    Console.WriteLine("无法从网址获取JSON数据");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"发生错误：{ex.Message}");
            }
        }

        static string FetchJsonFromUrl(string apiUrl)
        {
            using (var client = new WebClient())
            {
                client.Encoding = System.Text.Encoding.UTF8; // 设置编码为 UTF-8
                string result = client.DownloadString(apiUrl);
                return result.Replace("var returnjson= ", "");
            }

            #region 方式二
            //HttpWebRequest request = (HttpWebRequest)WebRequest.Create(apiUrl);
            //request.Method = "GET";
            //request.ContentType = "application/json;charset=UTF-8";

            //var httpResponse = (HttpWebResponse)request.GetResponse();
            //using (var dataStream = new StreamReader(httpResponse.GetResponseStream()))
            //{
            //    result = dataStream.ReadToEnd().Replace("var returnjson= ", "");
            //}
            #endregion
        }

        static List<FundManager> ParseJsonToFundManagers(string jsonStr)
        {
            FundManagerList fmList = JsonConvert.DeserializeObject<FundManagerList>(jsonStr);
            List<FundManager> fManagers = fmList.data.Select(data => new FundManager
            {
                Id = data[0],
                ManagerName = data[1],
                CompanyId = data[2],
                Company = data[3],
                Code = data[4],
                FundName = data[5],
                Totaldays = data[6],
                BestProfit = data[7],
                BestFundCode = data[8],
                BestFundName = data[9],
                TotalMoney = data[10]
            }).ToList();

            return fManagers;
        }
        static void JsonStrToXlsx(string jsonStr, string filePath)
        {
            List<FundManager> fundManagers = ParseJsonToFundManagers(jsonStr);
            if (File.Exists(filePath))
                File.Delete(filePath);

            using (var excel = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = excel.Workbook.Worksheets.Add("Fund Managers");
                int row = 1;
                int col = 1;

                // Write headers
                var headers = new List<string> { "Id", "ManagerName", "CompanyId", "Company", "Code", "FundName", "Totaldays", "BestProfit", "BestFundCode", "BestFundName", "TotalMoney" };
                foreach (var header in headers)
                {
                    worksheet.Cells[row, col++].Value = header;
                }

                // Write data
                row++;
                foreach (FundManager manager in fundManagers)
                {
                    col = 1;
                    worksheet.Cells[row, col++].Value = manager.Id;
                    worksheet.Cells[row, col++].Value = manager.ManagerName;
                    worksheet.Cells[row, col++].Value = manager.CompanyId;
                    worksheet.Cells[row, col++].Value = manager.Company;
                    worksheet.Cells[row, col++].Value = manager.Code;
                    worksheet.Cells[row, col++].Value = manager.FundName;
                    worksheet.Cells[row, col++].Value = manager.Totaldays;
                    worksheet.Cells[row, col++].Value = manager.BestProfit;
                    worksheet.Cells[row, col++].Value = manager.BestFundCode;
                    worksheet.Cells[row, col++].Value = manager.BestFundName;
                    worksheet.Cells[row, col++].Value = manager.TotalMoney;
                    row++;
                }

                for (int i = 1; i <= headers.Count; i++)
                {
                    worksheet.Column(i).AutoFit();
                }

                excel.Save();
            }
        }
    }
}