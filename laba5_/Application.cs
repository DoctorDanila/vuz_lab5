using System;
using System.Collections.Generic;
using System.Linq;

namespace LabWork
{
    public class Application
    {
        private List<ProductMovement> movements;
        private List<Category> categories;
        private List<Store> stores;
        private List<Product> products;
        private Logger logger;

        public Application(string filePath, Logger logger)
        {
            this.logger = logger;
            movements = DataReader.ReadProductMovements(filePath);
            categories = DataReader.ReadCategories(filePath);
            stores = DataReader.ReadStores(filePath);
            products = DataReader.ReadProducts(filePath);
            logger.Log("Данные загружены.");
        }

        public void ViewData()
        {
            Console.WriteLine("\tВыберите таблицу для просмотра:");
            Console.WriteLine("\t1 - movements");
            Console.WriteLine("\t2 - categories");
            Console.WriteLine("\t3 - stores");
            Console.WriteLine("\t4 - products");
            string choice = Console.ReadLine();
            switch (choice) 
            {
                case "1":
                    logger.Log("Просмотр данных movements.");
                    Console.WriteLine("Просмотр данных movements.");
                    foreach (var movement in movements) Console.WriteLine(movement);
                    break;
                case "2":
                    logger.Log("Просмотр данных categories.");
                    Console.WriteLine("Просмотр данных categories.");
                    foreach (var category in categories) Console.WriteLine(category);
                    break;
                case "3":
                    logger.Log("Просмотр данных stores.");
                    Console.WriteLine("Просмотр данных stores.");
                    foreach (var store in stores) Console.WriteLine(store);
                    break;
                case "4":
                    logger.Log("Просмотр данных products.");
                    Console.WriteLine("Просмотр данных products.");
                    foreach (var product in products) Console.WriteLine(product);
                    break;
                default:
                    Console.WriteLine("Такую таблицу выбрать нельзя :(");
                    break;
            }
        }

        public void DeleteElement(int operationID)
        {
            var movement = movements.FirstOrDefault(m => m.OperationID == operationID);
            if (movement != null)
            {
                movements.Remove(movement);
                logger.Log($"Элемент с OperationID {operationID} удален.");
                Console.WriteLine($"Элемент с OperationID {operationID} удален.");
            }
            else
            {
                logger.Log($"Элемент с OperationID {operationID} не найден.");
                Console.WriteLine($"Элемент с OperationID {operationID} не найден.");
            }
        }

        public void UpdateElement(int operationID, ProductMovement newMovement)
        {
            var movement = movements.FirstOrDefault(m => m.OperationID == operationID);
            if (movement != null)
            {
                movement.Date = newMovement.Date;
                movement.StoreID = newMovement.StoreID;
                movement.Article = newMovement.Article;
                movement.OperationType = newMovement.OperationType;
                movement.PackageCount = newMovement.PackageCount;
                movement.HasCustomerCard = newMovement.HasCustomerCard;
                logger.Log($"Элемент с OperationID {operationID} обновлен.");
                Console.WriteLine($"Элемент с OperationID {operationID} обновлен.");
            }
            else
            {
                logger.Log($"Элемент с OperationID {operationID} не найден.");
                Console.WriteLine($"Элемент с OperationID {operationID} не найден.");
            }
        }

        public void AddElement(ProductMovement movement)
        {
            movements.Add(movement);
            logger.Log($"Добавлен новый элемент с OperationID {movement.OperationID}.");
            Console.WriteLine($"Добавлен новый элемент с OperationID {movement.OperationID}.");
        }

        public ProductMovement GetMovementFromUser()
        {
            Console.WriteLine("Введите данные для нового элемента movement:");

            Console.Write("OperationID: ");
            int operationID = int.Parse(Console.ReadLine());

            Console.Write("Дата (дд.мм.гггг): ");
            DateTime date = DateTime.Parse(Console.ReadLine());

            Console.Write("StoreID: ");
            string storeID = Console.ReadLine();

            Console.Write("Article: ");
            string article = Console.ReadLine();

            Console.Write("Тип операции (Продажа/Поступление/Возврат): ");
            string operationType = Console.ReadLine();

            Console.Write("Количество упаковок: ");
            int packageCount = int.Parse(Console.ReadLine());

            Console.Write("Наличие карты клиента (Да/Нет): ");
            bool hasCustomerCard = Console.ReadLine().ToLower() == "да";

            return new ProductMovement
            {
                OperationID = operationID,
                Date = date,
                StoreID = storeID,
                Article = article,
                OperationType = operationType,
                PackageCount = packageCount,
                HasCustomerCard = hasCustomerCard
            };
        }

        public void ExecuteQueries()
        {
            logger.Log("Выполнение запросов.");
            
            // Запрос 1: Товары, проданные в магазине Р7
            var salesInStore = (from m in movements
                                join p in products on m.Article equals p.Article
                                where m.StoreID == "Р7" && m.OperationType == "Продажа"
                                select p).ToList();
            Console.WriteLine("Товары, проданные в магазине Р7:");
            foreach (var product in salesInStore) Console.WriteLine(product);

            // Запрос 2: Общая стоимость товаров из категории "Игрушки на радиоуправлении 12+"
            var totalCost = (from m in movements
                             join p in products on m.Article equals p.Article
                             join c in categories on p.CategoryID equals c.ID
                             where c.Name == "Игрушки на радиоуправлении" && c.AgeRestriction == "12+" && m.OperationType == "Продажа"
                             select p.PackagePrice * m.PackageCount).Sum();
            Console.WriteLine($"Общая стоимость товаров из категории 'Игрушки на радиоуправлении 12+': {totalCost}");

            // Запрос 3: Магазины, в которых продавались товары из категории "Куклы"
            var storesWithDolls = (from m in movements
                                   join p in products on m.Article equals p.Article
                                   join c in categories on p.CategoryID equals c.ID
                                   join s in stores on m.StoreID equals s.ID
                                   where c.Name == "Куклы" && m.OperationType == "Продажа"
                                   select s).Distinct().ToList();
            Console.WriteLine("Магазины, в которых продавались товары из категории 'Куклы':");
            foreach (var store in storesWithDolls) Console.WriteLine(store);
        }

        public void SaveChanges(string filePath)
        {
            try
            {
                DataWriter.WriteProductMovements(filePath, movements);
                logger.Log("Изменения сохранены в Excel.");
                Console.WriteLine("Изменения сохранены в Excel.");
            }
            catch (Exception ex)
            {
                logger.Log($"Ошибка при сохранении изменений: {ex.Message}");
                Console.WriteLine($"Ошибка при сохранении изменений: {ex.Message}");
            }
        }
    }
}