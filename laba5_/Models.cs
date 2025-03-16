using System;

namespace LabWork
{
    public class ProductMovement
    {
        public int OperationID { get; set; }
        public DateTime Date { get; set; }
        public string StoreID { get; set; }
        public string Article { get; set; }
        public string OperationType { get; set; }
        public int PackageCount { get; set; }
        public bool HasCustomerCard { get; set; }

        public override string ToString()
        {
            return $"OperationID: {OperationID}, Date: {Date.ToShortDateString()}, StoreID: {StoreID}, " +
                   $"Article: {Article}, OperationType: {OperationType}, PackageCount: {PackageCount}, " +
                   $"HasCustomerCard: {HasCustomerCard}";
        }
    }

    public class Category
    {
        public int ID { get; set; }
        public string Name { get; set; }
        public string AgeRestriction { get; set; }

        public override string ToString()
        {
            return $"ID: {ID}, Name: {Name}, AgeRestriction: {AgeRestriction}";
        }
    }

    public class Store
    {
        public string ID { get; set; }
        public string District { get; set; }
        public string Address { get; set; }

        public override string ToString()
        {
            return $"ID: {ID}, District: {District}, Address: {Address}";
        }
    }

    public class Product
    {
        public string Article { get; set; }
        public int CategoryID { get; set; }
        public string Name { get; set; }
        public string Unit { get; set; }
        public int PackageQuantity { get; set; }
        public decimal PackagePrice { get; set; }

        public override string ToString()
        {
            return $"Article: {Article}, CategoryID: {CategoryID}, Name: {Name}, " +
                   $"Unit: {Unit}, PackageQuantity: {PackageQuantity}, PackagePrice: {PackagePrice}";
        }
    }
}