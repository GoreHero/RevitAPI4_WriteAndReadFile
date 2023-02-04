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
using System.Windows.Forms;


namespace RevitAPI_L8_ReadJSON
{
    [Transaction(TransactionMode.Manual)]

    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            // указываем файл с которым будем работать
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                Filter = "Json files (*.json) | *.json"
            };
            string filePath = string.Empty;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                filePath = openFileDialog.FileName;
            }
            if (string.IsNullOrEmpty(filePath))
                return Result.Cancelled;

            //собираем все помещения 
            var rooms = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .Cast<Room>()
                .ToList();
            //приобразуем формат JSON
            string json = File.ReadAllText(filePath);
            //получаем список
            List<RoomData> roomDataList = JsonConvert.DeserializeObject<List<RoomData>>(json);

            return Result.Succeeded;
        }
    }
}
