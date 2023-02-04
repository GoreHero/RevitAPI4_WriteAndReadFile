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
using System.Windows.Forms;

namespace RevitAPI_L6_ReadExcel
{
    [Transaction(TransactionMode.Manual)]

    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            //конструкция вызывает диалоговое окно
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop), //путь по умолчанию
                Filter = "Excel files(*.xlsx) | */xlsx" //фильтр расширения фалов
            };
            string filePath = string.Empty;
            if (openFileDialog1.ShowDialog() == DialogResult.OK) //сохраняем путь если все ОК
            {
                filePath = openFileDialog1.FileName;
            }
            if (string.IsNullOrEmpty(filePath)) //если путь не задан
                return Result.Cancelled; //завершаем 

            var rooms = new FilteredElementCollector(doc) //собираем все помещения 
                .OfCategory(BuiltInCategory.OST_Rooms)
                .Cast<Room>()
                .ToList();

            //Далее нужно прочитать данные из файла
            //открываем файл
            using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read)) //FileStream(путь, открыть существующий, только чтение)
            {
                IWorkbook workbook = new XSSFWorkbook(filePath); //путь к книге
                ISheet sheet = workbook.GetSheetAt(index: 0); //выбор листа

                //далее проходим построчно  и собираем данные
                int rowIndex = 0; //первая строка
                while (sheet.GetRow(rowIndex) != null) //до тех пор пока d строке есть данные, продолжается считывание
                {
                    
                    if (sheet.GetRow(rowIndex).GetCell(0) == null || //если в указанной строке и столбце ячейка пустая
                        sheet.GetRow(rowIndex).GetCell(1) == null) //либо во втором столбце
                    {
                        rowIndex++;
                        continue; 
                    } 
                    //считываем данные из файла Excel
                    string name = sheet.GetRow(rowIndex).GetCell(0).StringCellValue; //лист.строка.ячейка(0).тип данных ячейки(строка)
                    string number = sheet.GetRow(rowIndex).GetCell(1).StringCellValue; //лист.строка.ячейка(1).тип данных(строка)
                    //из всех собраных элементов ищем с указанным номером в экселе

                    var room = rooms.FirstOrDefault(r => r.Number.Equals(number)); //первое помещения с подобным номером 
                    if (room != null)//если помещение не найдено, продолжаем цикл
                    {
                        rowIndex++;
                        continue;
                    }
                    using (var ts = new Transaction(doc, "Set parameter")) //транзакция
                    {
                        ts.Start();
                        room.get_Parameter(BuiltInParameter.ROOM_NAME).Set(name);//берем параметр по внутреннему назанию. Вписываем новые данные
                        ts.Commit();
                    }

                    rowIndex++; //переход к новой строке
                }
            }

            return Result.Succeeded;
        }
    }
}
