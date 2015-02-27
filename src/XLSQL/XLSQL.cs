using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Dapper;
using System.Data;
using System.Reflection;
using Newtonsoft.Json;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Diagnostics;
using System.Globalization;

namespace XLSQL
{
    public static class XLSQL
    {

        private static System.Data.OleDb.OleDbConnection GetConnection(FileInfo source)
        {
            string connStr = string.Format("provider=Microsoft.ACE.OLEDB.12.0;Data Source='{0}';Extended Properties='Excel 12.0;HDR=YES;IMEX=1';", source.FullName);
            return new System.Data.OleDb.OleDbConnection(connStr);
        }

        public static IEnumerable<string> GetSheetNames(FileInfo source)
        {
            using (var db = GetConnection(source))
            {
                db.Open();
                IEnumerable<string> sheets = null;
                using (DataTable schema = db.GetSchema("Tables"))
                {
                    sheets = schema.Rows[0].ItemArray.Where(m => m.ToString().EndsWith("$")).Select(m => string.Format("[{0}]", m.ToString()));
                }
                db.Close();
                return sheets;
            }

        }

        public static void SQLtoExcel(FileInfo source, FileInfo dest, string sql)
        {
            var queryResult = QuerySheet(source, sql);
            using (ExcelPackage pkg = new ExcelPackage(dest)) {
                using (var planilha = pkg.Workbook.Worksheets.Add("Planilha1")) {

                    int l = 1;
                    var firstItem = queryResult.First();
                    for (int c = 0; c < firstItem.Keys.Count(); c++)
                    {
                        var key = firstItem.Keys.ToArray()[c];

                        planilha.Cells[l, c+1].IsRichText = true;
                        ExcelRichText richtext = planilha.Cells[l, c+1].RichText.Add(key);
                        richtext.Bold = true;

                    }
                    l++;
                    foreach (var item in queryResult)
                    {
                        for (int c = 0; c < item.Keys.Count(); c++) {
                            
                            var key = item.Keys.ToArray()[c];
                            var value = item[key]; 

                            var cell = planilha.Cells[l, c + 1];
                            cell.Value = value;

                            if (value is DateTime)
                            {
                                cell.Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                            }
                        }
                        l++;
                    }
                    pkg.Save();
                }

            }
        }

        private static IEnumerable<Dictionary<string, object>> QuerySheet(FileInfo arquivo, string sql)
        {
            using (var db = GetConnection(arquivo))
            {
                db.Open();
                var rows = db.Query(sql);
                var json = JsonConvert.SerializeObject(rows);
                IEnumerable<Dictionary<string, object>> obj = JsonConvert.DeserializeObject<IEnumerable<Dictionary<string, object>>>(json);
                db.Close();
                return obj;
            }
        }


    }
}
