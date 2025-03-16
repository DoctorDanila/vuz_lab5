using System;

namespace LabWork
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = "exel_tables.xls";

            Console.WriteLine("Создать новый файл логов? (y/n)");
            bool createNewLog = Console.ReadLine().ToLower() == "y";

            string logFilePath = createNewLog ? GetLogFileName() : "log.txt";

            var logger = new Logger(logFilePath, !createNewLog);
            var app = new Application(filePath, logger);

            while (true)
            {
                Console.WriteLine("\tВыберите действие:");
                Console.WriteLine("\t1. Просмотр данных");
                Console.WriteLine("\t2. Удалить элемент");
                Console.WriteLine("\t3. Обновить элемент");
                Console.WriteLine("\t4. Добавить элемент");
                Console.WriteLine("\t5. Выполнить запросы");
                Console.WriteLine("\t6. Сохранить изменения и выйти");
                Console.WriteLine("\t7. Выйти без сохранения");

                string choice = Console.ReadLine();

                switch (choice)
                {
                    case "1":
                        app.ViewData();
                        break;
                    case "2":
                        Console.WriteLine("Введите OperationID для удаления из movements:");
                        if (int.TryParse(Console.ReadLine(), out int deleteId)) app.DeleteElement(deleteId);
                        else Console.WriteLine("Некорректный ввод.");
                        break;
                    case "3":
                        Console.WriteLine("Введите OperationID для обновления movements:");
                        if (int.TryParse(Console.ReadLine(), out int updateId))
                        {
                            var newMovement = app.GetMovementFromUser();
                            app.UpdateElement(updateId, newMovement);
                        }
                        else Console.WriteLine("Некорректный ввод.");
                        break;
                    case "4":
                        var movement = app.GetMovementFromUser();
                        app.AddElement(movement);
                        break;
                    case "5":
                        app.ExecuteQueries();
                        break;
                    case "6":
                        app.SaveChanges(filePath);
                        logger.Close();
                        return;
                    case "7":
                        logger.Close();
                        return;
                    default:
                        Console.WriteLine("Команда не найдена.");
                        break;
                }
            }
        }

        private static string GetLogFileName()
        {
            Console.WriteLine("Введите название файла для логов (например, log_log.txt):");
            return Console.ReadLine();
        }
    }
}