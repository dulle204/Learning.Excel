using System;
using System.Collections;

namespace Learning.Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            var ListOfPersons = new ExcelDocHandler().GetByID(5);
            string s = $"{ListOfPersons.ID} | {ListOfPersons.FirstName} | {ListOfPersons.LastName} | {ListOfPersons.LastName} | {ListOfPersons.Gender.ToString()}";
            Console.WriteLine(s);

            //foreach (var item in ListOfPersons)
            //{
            //    string s = $"{item.ID} | {item.FirstName} | {item.LastName} | {item.LastName} | {item.Gender.ToString()}";
            //    Console.WriteLine(s);
            //}

            Console.ReadKey();
        }
    }
}
