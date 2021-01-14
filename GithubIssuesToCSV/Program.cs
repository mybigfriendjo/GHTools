using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Text;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;

namespace GithubIssuesToCSV
{
    public class Program
    {
        public static void Main(string[] args)
        {
            if (args == null || args.Length < 2)
            {
                Console.WriteLine("not enough arguments");
                return;
            }

            string url = args[0];
            string exportPath = args[1];

            if (string.IsNullOrWhiteSpace(url))
            {
                Console.WriteLine("empty argument 'url'");
                return;
            }

            if (string.IsNullOrWhiteSpace(exportPath))
            {
                Console.WriteLine("empty argument 'exportPath'");
                return;
            }

            HttpClient client = new HttpClient();
            client.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("GHTools",
                                                                                  Assembly.GetExecutingAssembly().GetName().Version
                                                                                          .ToString()));
            string issueJSON = client.GetStringAsync(new Uri(url)).Result;

            dynamic issues = JArray.Parse(issueJSON);

            if (File.Exists(exportPath))
            {
                File.Delete(exportPath);
            }

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage package = new ExcelPackage(new FileInfo(exportPath)))
            {
                ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Git Issues");

                sheet.SetValue(1, 1, "number");
                sheet.SetValue(1, 2, "url");
                sheet.SetValue(1, 3, "title");
                sheet.SetValue(1, 4, "labels");
                sheet.SetValue(1, 5, "state");
                sheet.SetValue(1, 6, "milestone");
                sheet.SetValue(1, 7, "created");
                sheet.SetValue(1, 8, "updated");
                sheet.SetValue(1, 9, "body");

                int currentRow = 2;

                foreach (dynamic issue in issues)
                {
                    sheet.SetValue(currentRow, 1, issue.number.ToString());
                    sheet.SetValue(currentRow, 2, issue.html_url.ToString());
                    sheet.SetValue(currentRow, 3, issue.title.ToString());

                    StringBuilder buf = new StringBuilder();
                    foreach (dynamic label in issue.labels)
                    {
                        buf.Append(label.name).Append(",");
                    }

                    buf.Length--;

                    sheet.SetValue(currentRow, 4, buf.ToString());
                    sheet.SetValue(currentRow, 5, issue.state.ToString());

                    if (issue.milestone != null)
                    {
                        sheet.SetValue(currentRow, 6, issue.milestone.title.ToString());
                    }

                    sheet.SetValue(currentRow, 7, issue.created_at.ToString());
                    sheet.SetValue(currentRow, 8, issue.updated_at.ToString());
                    sheet.SetValue(currentRow, 9, issue.body.ToString());

                    currentRow++;
                }

                for (int i = 1; i < 10; i++)
                {
                    sheet.Column(i).AutoFit();
                }

                package.Save();
            }
        }
    }
}