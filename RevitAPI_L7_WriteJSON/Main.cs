using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPI_L7_WriteJSON
{
    [Transaction(TransactionMode.Manual)]

    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            var rooms = new FilteredElementCollector(doc) //собираем все помещения 
                .OfCategory(BuiltInCategory.OST_Rooms)
                .Cast<Room>()
                .ToList();
            var roomDataList = new List<RoomData>(); 
            //преобразовываем все в список румдата
            foreach (var room in rooms) //заполняю данные
            {
                roomDataList.Add(new RoomData //собираем список списком имя-номер
                {
                    Name = room.Name,
                    Number = room.Number
                });
            }
            //переводим список румДатаЛист в формат JSON
            string json = JsonConvert.SerializeObject(roomDataList, Formatting.Indented);//SerializeObject(что записывать, Indented(с переводом на новую строку))

            //сохраняем в файл
            File.WriteAllText(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop),"data.json"),json);


            return Result.Succeeded;
        }
    }
}
