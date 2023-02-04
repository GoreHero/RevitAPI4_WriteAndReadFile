using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RevitAPI_W1_OutputValue
{
    [Transaction(TransactionMode.Manual)]

    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            string wallInfo = string.Empty; //для записи данных
            //сбор всех стен
            var walls = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Walls)
                .Cast<Wall>()
                .ToList(); 
            //сбор параметров
            foreach (Wall wall in walls)
            {
                string wallType = wall.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsString();
                double wallVolume = wall.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED).AsDouble();
                wallInfo += $"{wallType}\t{wallVolume}\t {Environment.NewLine}";
            }

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

            File.WriteAllText(selectedFilePath, wallInfo);

            return Result.Succeeded;
        }
    }
}
