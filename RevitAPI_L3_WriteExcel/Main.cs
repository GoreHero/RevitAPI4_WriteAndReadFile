using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPI_L5_WriteExcel
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            string roomInfo = string.Empty; //переменная куда будут записанны данные
            var rooms = new FilteredElementCollector(doc)  //сбор самих помещений
                .OfCategory(BuiltInCategory.OST_Rooms) //собираем по категории
                .Cast<Room>() //получаемые элементы преобразуем в помещения
                .ToList(); //записываем в список

            string excelPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "rooms.xlsx"); //путь сохранения файла (рабочий стол) и название файла
            //создаем файл
            using (FileStream stream = new FileStream(excelPath, FileMode.Create, FileAccess.Write)) //FileMode.Create - создает новый файл
            {
                IWorkbook workbook = new XSSFWorkbook(); //создали виртальную книгу(файл экселя)
                ISheet sheet = workbook.CreateSheet("Лист1"); // создаем лист 
                                                              //далее нужно заполнить данными, нужно создать метод

                int rowIndex = 0;
                foreach (var room in rooms) //проходим по каждому помещению которое есть в модели и записываем в строку 0 в файле эксель
                {
                    sheet.SetCellValue(rowIndex, ColumnIndex: 0, room.Name);
                    sheet.SetCellValue(rowIndex, ColumnIndex: 1, room.Number);
                    sheet.SetCellValue(rowIndex, ColumnIndex: 2, room.Area);
                    rowIndex++;
                }
                //после всех действий файл нужно закрыть=>
                workbook.Write(stream);
                workbook.Close();

            }
            System.Diagnostics.Process.Start(excelPath); //автоматическое открые файла

            return Result.Succeeded;
        }
    }
}
