using Autodesk.Revit.DB;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPI_W2_OutputValuePipes
{
    public static class SheetExts
    {
        public static void SetCellValue<T>(this ISheet sheet, int rowIndex, int columnIndex, T value)


        {
            var cellReference = new CellReference(rowIndex, columnIndex);//ссылка на ячейку
            //проверка существования ячейки
            var row = sheet.GetRow(cellReference.Row);
            if (row == null)
                row = sheet.CreateRow(cellReference.Row);

            var cell = row.GetCell(cellReference.Col);
            if (cell == null)
                cell = row.CreateCell(cellReference.Col);

            //проверка типа введенного значения на тип
            if (value is string)
            { cell.SetCellValue((string)(object)value); }
            else if (value is double)
            { cell.SetCellValue((double)(object)value); }
            else if (value is int)
            { cell.SetCellValue((int)(object)value); }
            else if (value is ElementId)
            { cell.SetCellValue((IRichTextString)(ElementId)(object)value); }

        }
    }
}