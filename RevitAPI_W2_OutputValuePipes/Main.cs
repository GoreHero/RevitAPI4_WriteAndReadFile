using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPI_W2_OutputValuePipes
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            //сбор всех труб
            var pipes = new FilteredElementCollector(doc)
                .OfClass(typeof(Pipe))
                .Cast<Pipe>()
                .ToList();
            //создание файла
            string excelPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "pipes.xlsx");
            using (FileStream stream = new FileStream(excelPath, FileMode.Create, FileAccess.Write))
            {
                IWorkbook workbook = new XSSFWorkbook(); //создание виртуальной книги
                ISheet sheet = workbook.CreateSheet("pipeInfo"); //создание листа
                int rowIndex = 0; //номер первой строки
                foreach (var pipe in pipes) //проходимся по строкам
                {
                    sheet.SetCellValue(rowIndex, columnIndex: 0, pipe.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString());
                    sheet.SetCellValue(rowIndex, columnIndex: 1, pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER).AsValueString());
                    sheet.SetCellValue(rowIndex, columnIndex: 2, pipe.get_Parameter(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM).AsValueString());
                    sheet.SetCellValue(rowIndex, columnIndex: 3, pipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsValueString());

                    rowIndex++;

                }
                //Закрытие файла
                workbook.Write(stream);
                workbook.Close();
            }
            System.Diagnostics.Process.Start(excelPath); //автоматическое открые файла

            return Result.Succeeded;
        }
    }
}
