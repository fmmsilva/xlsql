using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XLSQLDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            FileInfo source = new FileInfo(@"c:\temp\exemplo.xlsx");
            FileInfo dest = new FileInfo(@"c:\temp\destino.xlsx");

            // Get Sheet Names
            var sheets = XLSQL.XLSQL.GetSheetNames(source);

            string sql = @"select Nome, count(Nome)
                            from [Planilha1$] 
                            group by Nome
                            order by Nome desc";
            if (dest.Exists)
            {
                dest.Delete();
            }
            XLSQL.XLSQL.SQLtoExcel(source, dest, sql);
            Process.Start(dest.FullName);

        }

    }
}
