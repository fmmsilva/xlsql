# xlsql [![Build status](https://ci.appveyor.com/api/projects/status/9hgk0u4pskjwjlii?svg=true)](https://ci.appveyor.com/project/fmmsilva/xlsql)
Use SQL to query Excel tables

```c#
FileInfo source = new FileInfo(@"c:\temp\source.xlsx");
FileInfo dest = new FileInfo(@"c:\temp\dest.xlsx");

// Get Sheet Names
var sheets = XLSQL.XLSQL.GetSheetNames(source);

string sql = @"select Col, count(Col)
                from [Sheet1$] 
                group by Col
                order by Col desc";

XLSQL.XLSQL.SQLtoExcel(source, dest, sql);
Process.Start(dest.FullName);
```
