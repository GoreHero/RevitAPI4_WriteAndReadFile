using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPI_L5_WriteExcel
{
    public static class SheetExts //в этом классе создаем метод расширения для работы с листом
    {
        public static void SetCellValue<T>(this ISheet sheet, int rowIndex, int ColumnIndex, T value) //джинерик метод для объектов типа Sheet, номер строки, номер столбца, значение
        {
            var cellReference = new CellReference(rowIndex, ColumnIndex); //ссылка на ячейку
            // проверяем если ячейки не существует 
            var row = sheet.GetRow(cellReference.Row);
            if (row == null)
                row = sheet.CreateRow(cellReference.Row); //если строки нет, создаем
            var cell = row.GetCell(cellReference.Col); //указываем конкретную ячейку, если ее нет создаем
            if (row == null)
                cell = row.CreateCell(cellReference.Col);

            //проверяем тип введенного значения
            if (value is string)
                cell.SetCellValue((string)(object)value);
            else if (value is double)
                cell.SetCellValue((double)(object)value);
            else if (value is int)
                cell.SetCellValue((int)(object)value);
        }
    }
}