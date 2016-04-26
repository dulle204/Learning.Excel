using System;
using System.Collections;

namespace Learning.Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            var ListOfPersons = new ExcelDocHandler().GetByName(Console.ReadLine());
            //string s = $"{ListOfPersons.ID} | {ListOfPersons.FirstName} | {ListOfPersons.LastName} | {ListOfPersons.Email} | {ListOfPersons.Gender.ToString()}";
            //Console.WriteLine(s);

            foreach (var item in ListOfPersons)
            {
                string s = $"{item.ID} | {item.FirstName} | {item.LastName} | {item.Email} | {item.Gender.ToString()}";
                Console.WriteLine(s);
            }

            Console.ReadKey();
        }
    }
}
