using System;

namespace Learning.Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            var ListOfPersons = new ExcelDocHandler().GetByName("Diana");

            foreach (var item in ListOfPersons)
            {
                string s = $"{item.ID} | {item.FirstName} | {item.LastName} | {item.LastName} | {item.Gender.ToString()}";
                Console.WriteLine(s);
            }

            Console.ReadKey();
        }
    }
}
