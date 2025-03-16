using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using System.Collections.Generic;
using System.IO;

namespace LabWork
{
    public class DataWriter
    {
        public static void WriteProductMovements(string filePath, List<ProductMovement> movements)
        {
            HSSFWorkbook workbook;

            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                workbook = new HSSFWorkbook(stream); 
            }

            var worksheet = workbook.GetSheetAt(0); 

            for (int row = worksheet.LastRowNum; row >= 1; row--)
            {
                var currentRow = worksheet.GetRow(row);
                if (currentRow != null) worksheet.RemoveRow(currentRow);
            }

            // новые данные
            for (int i = 0; i < movements.Count; i++)
            {
                var row = worksheet.CreateRow(i + 1); 
                row.CreateCell(0).SetCellValue(movements[i].OperationID);
                row.CreateCell(1).SetCellValue(movements[i].Date);
                row.CreateCell(2).SetCellValue(movements[i].StoreID);
                row.CreateCell(3).SetCellValue(movements[i].Article);
                row.CreateCell(4).SetCellValue(movements[i].OperationType);
                row.CreateCell(5).SetCellValue(movements[i].PackageCount);
                row.CreateCell(6).SetCellValue(movements[i].HasCustomerCard ? "Да" : "Нет");
            }

            // Сохранение
            using (var stream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(stream);
            }
        }
    }
}