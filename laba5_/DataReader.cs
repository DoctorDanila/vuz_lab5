using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Globalization;

namespace LabWork
{
    public class DataReader
    {
        public static List<ProductMovement> ReadProductMovements(string filePath)
        {
            var movements = new List<ProductMovement>();
            try
            {
                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    var workbook = new HSSFWorkbook(stream); 
                    var worksheet = workbook.GetSheetAt(0); 

                    for (int row = 1; row <= worksheet.LastRowNum; row++) 
                    {
                        var currentRow = worksheet.GetRow(row);
                        if (currentRow != null)
                            movements.Add(new ProductMovement
                            {
                                OperationID = GetIntValue(currentRow.GetCell(0)),
                                Date = GetDateTimeValue(currentRow.GetCell(1)),
                                StoreID = GetStringValue(currentRow.GetCell(2)),
                                Article = GetStringValue(currentRow.GetCell(3)),
                                OperationType = GetStringValue(currentRow.GetCell(4)),
                                PackageCount = GetIntValue(currentRow.GetCell(5)),
                                HasCustomerCard = GetStringValue(currentRow.GetCell(6)) == "Да"
                            });
                    }
                }
            }
            catch (IOException ex)
            {
                Console.WriteLine($"Ошибка при чтении файла Excel: {ex.Message}");
                Console.WriteLine("Убедитесь, что файл не открыт в другой программе.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Произошла ошибка: {ex.Message}");
            }
            return movements;
        }

        public static List<Category> ReadCategories(string filePath)
        {
            var categories = new List<Category>();
            try
            {
                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    var workbook = new HSSFWorkbook(stream); 
                    var worksheet = workbook.GetSheetAt(2); 

                    for (int row = 1; row <= worksheet.LastRowNum; row++) 
                    {
                        var currentRow = worksheet.GetRow(row);
                        if (currentRow != null)
                            categories.Add(new Category
                            {
                                ID = GetIntValue(currentRow.GetCell(0)),
                                Name = GetStringValue(currentRow.GetCell(1)),
                                AgeRestriction = GetStringValue(currentRow.GetCell(2))
                            });
                    }
                }
            }
            catch (IOException ex)
            {
                Console.WriteLine($"Ошибка при чтении файла Excel: {ex.Message}");
                Console.WriteLine("Убедитесь, что файл не открыт в другой программе.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Произошла ошибка: {ex.Message}");
            }
            return categories;
        }

        public static List<Store> ReadStores(string filePath)
        {
            var stores = new List<Store>();
            try
            {
                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    var workbook = new HSSFWorkbook(stream); 
                    var worksheet = workbook.GetSheetAt(3); 

                    for (int row = 1; row <= worksheet.LastRowNum; row++) 
                    {
                        var currentRow = worksheet.GetRow(row);
                        if (currentRow != null)
                            stores.Add(new Store
                            {
                                ID = GetStringValue(currentRow.GetCell(0)),
                                District = GetStringValue(currentRow.GetCell(1)),
                                Address = GetStringValue(currentRow.GetCell(2))
                            });
                    }
                }
            }
            catch (IOException ex)
            {
                Console.WriteLine($"Ошибка при чтении файла Excel: {ex.Message}");
                Console.WriteLine("Убедитесь, что файл не открыт в другой программе.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Произошла ошибка: {ex.Message}");
            }
            return stores;
        }

        public static List<Product> ReadProducts(string filePath)
        {
            var products = new List<Product>();
            try
            {
                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    var workbook = new HSSFWorkbook(stream);
                    var worksheet = workbook.GetSheetAt(1); 

                    for (int row = 1; row <= worksheet.LastRowNum; row++)
                    {
                        var currentRow = worksheet.GetRow(row);
                        if (currentRow != null)
                            products.Add(new Product
                            {
                                Article = GetStringValue(currentRow.GetCell(0)),
                                CategoryID = GetIntValue(currentRow.GetCell(1)),
                                Name = GetStringValue(currentRow.GetCell(2)),
                                Unit = GetStringValue(currentRow.GetCell(3)),
                                PackageQuantity = GetIntValue(currentRow.GetCell(4)),
                                PackagePrice = GetDecimalValue(currentRow.GetCell(5))
                            });
                    }
                }
            }
            catch (IOException ex)
            {
                Console.WriteLine($"Ошибка при чтении файла Excel: {ex.Message}");
                Console.WriteLine("Убедитесь, что файл не открыт в другой программе.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Произошла ошибка: {ex.Message}");
            }
            return products;
        }

        private static int GetIntValue(ICell cell)
        {
            if (cell == null || cell.CellType == CellType.Blank) return 0;

            if (cell.CellType == CellType.Numeric) return (int)cell.NumericCellValue;

            if (cell.CellType == CellType.String && int.TryParse(cell.StringCellValue, out int result)) return result;

            return 0;
        }

        private static DateTime GetDateTimeValue(ICell cell)
        {
            if (cell == null || cell.CellType == CellType.Blank)
            {
                Console.WriteLine("Предупреждение: Пустая ячейка с датой. Используется значение по умолчанию (DateTime.MinValue).");
                return DateTime.MinValue;
            }

            try
            {
                if (cell.CellType == CellType.Numeric && DateUtil.IsCellDateFormatted(cell)) return cell.DateCellValue ?? DateTime.MinValue; 

                if (cell.CellType == CellType.String)
                {
                    string cellValue = cell.StringCellValue;

                    if (DateTime.TryParseExact(cellValue, "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime result)) return result;

                    if (DateTime.TryParseExact(cellValue, "M.d.yy", CultureInfo.InvariantCulture, DateTimeStyles.None, out result)) return result;
                }

                if (cell.CellType == CellType.Numeric)
                {
                    double excelDateValue = cell.NumericCellValue;
                    return DateTime.FromOADate(excelDateValue); 
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при чтении даты: {ex.Message}. Используется значение по умолчанию (DateTime.MinValue).");
            }

            Console.WriteLine("Предупреждение: Некорректное значение даты. Используется значение по умолчанию (DateTime.MinValue).");
            return DateTime.MinValue;
        }
        private static string GetStringValue(ICell cell)
        {
            if (cell == null || cell.CellType == CellType.Blank) return string.Empty;

            if (cell.CellType == CellType.String) return cell.StringCellValue;

            if (cell.CellType == CellType.Numeric) return cell.NumericCellValue.ToString();

            return string.Empty;
        }

        private static decimal GetDecimalValue(ICell cell)
        {
            if (cell == null || cell.CellType == CellType.Blank) return 0m;

            if (cell.CellType == CellType.Numeric) return (decimal)cell.NumericCellValue;

            if (cell.CellType == CellType.String && decimal.TryParse(cell.StringCellValue, out decimal result)) return result;

            return 0m;
        }
    }
}