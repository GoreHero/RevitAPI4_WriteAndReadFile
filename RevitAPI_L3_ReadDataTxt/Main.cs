using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using RevitAPITrainingReadWrite;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RevitAPI_L3_ReadDataTxt
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;



            OpenFileDialog openFileDialog = new OpenFileDialog(); //переменная для того чтобы выбрать файл
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop); //директория по умолчанию рабочий стол
            openFileDialog.Filter = "All files(*.*)|*.*"; //фильтр фалов(все файлы)

            string filePath = string.Empty; //переменная для сохранения пути
            if (openFileDialog.ShowDialog() == DialogResult.OK)//проверка выбран ли файл
            {
                filePath = openFileDialog.FileName;
            }
            if (string.IsNullOrEmpty(filePath)) //если путь не найден, заканчиваем выполнение программы
                return Result.Cancelled;

            var lines = File.ReadAllLines(filePath).ToList(); //ReadAllLines=>забираем все строки из заданного файла

            List<RoomData> roomDataList = new List<RoomData>(); //создаем переменную списка с RoomData(будет заполняться список[имя,номер],[],[]....[]
            foreach (var line in lines)//проходим по строкам файла
            {
                List<string> values = line.Split(';').ToList(); //создаем еще один список разделяя данные по знаку табуляции, каждая строка будет одним списком
                roomDataList.Add(new RoomData //создаем экземпляры класса RoomData с заполнеными полями
                {
                    Name = values[0],
                    Number = values[1]
                });

            }


            //будем собирать данные о помещении 
            string roomInfo = string.Empty; //переменная куда будут записанны данные
            var rooms = new FilteredElementCollector(doc)  //сбор самих помещений
                .OfCategory(BuiltInCategory.OST_Rooms) //собираем по категории
                .Cast<Room>() //получаемые элементы преобразуем в помещения
                .ToList(); //записываем в список

            using (var ts = new Transaction(doc, "Set parameters")) // записываем параметры
            {
                ts.Start();
                foreach (RoomData roomData in roomDataList) //RoomData список помещений из текстового файла
                {

                    Room room = rooms.FirstOrDefault(r => r.Number.Equals(roomData.Number));//FirstOrDefault обращаемся ко всем помещениям в модели(находим то помещение которое совпадает по списку)
                    //находим помещение с тем номером которое указанно у нас в текстовом файле
                    if (room == null) //если помещения с таким номером нет
                        continue;

                    room.get_Parameter(BuiltInParameter.ROOM_NAME).Set(roomData.Name); //обращаемся к параметру имени помщения и устанавливаем новое значение которое попределено в румДата
                }

                ts.Commit();
            }


            return Result.Succeeded;
        }
    }
}
