using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;


//Запись данных в текстовый файл+запроспути сохранения файла
namespace RevitAPI_L2_WriteDataTxt
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            //будем собирать данные о помещении 
            string roomInfo = string.Empty; //переменная куда будут записанны данные
            var rooms = new FilteredElementCollector(doc)  //сбор самих помещений
                .OfCategory(BuiltInCategory.OST_Rooms) //собираем по категории
                .Cast<Room>() //получаемые элементы преобразуем в помещения
                .ToList(); //записываем в список

            //Заполняем данный параметр
            foreach (Room room in rooms)
            {
                string roomName = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString(); //Извлекаем имя помещения
                roomInfo += $"{roomName}\t{room.Number}\t{room.Area}{Environment.NewLine}"; //Заполняем переменную инфо новыми данными //{Environment.NewLine} новая строка
            }
            //Запрос пути сохранения
            var saveDialog = new SaveFileDialog //создаем переменную которая отвечает за путь 
            {
                OverwritePrompt = true, //запрос на перезапись
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop), //сохранение по умолчанию(Рабочий стол)
                Filter = "All files (*.*)|*.*", //файлы всех форматов
                FileName = "roomInfo.csv", //Имя по умолчанию, при сохранении можно изменить
                DefaultExt = ".csv" //Значения расширения по умолчанию

            };
            string selectedFilePath = string.Empty;//переменная с выбранным пользователем ссылкой
            if (saveDialog.ShowDialog() == DialogResult.OK) //если путь был указан то забираем путь в переменную
            {
                selectedFilePath = saveDialog.FileName;
            }
            if (string.IsNullOrEmpty(selectedFilePath)) //если путь не задан, возращаемся обратно
                return Result.Cancelled;

            File.WriteAllText(selectedFilePath, roomInfo); //сохраняем файл


            return Result.Succeeded;
        }
    }
}

