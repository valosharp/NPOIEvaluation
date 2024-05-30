// See https://aka.ms/new-console-template for more information
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

using System.Diagnostics;
using System.Reflection;


var basePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)!;
var src = Path.Combine(basePath, "Data/Funding Calc bad.xlsx");
var target = Path.Combine(basePath, "Data/Funding Calc output.xlsx");

using Stream rfs = File.OpenRead(src);
using IWorkbook workbook = new XSSFWorkbook(rfs);
using FileStream fs = File.Create(target);
workbook.Write(fs, false);

Process.Start(
            new ProcessStartInfo
            {
                FileName = target,
                UseShellExecute = true
            }
        );