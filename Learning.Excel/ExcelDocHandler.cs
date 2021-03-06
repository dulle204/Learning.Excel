﻿using Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace Learning.Excel
{
    public class ExcelDocHandler
    {
        public List<Person> GetAll()
        {
            FileStream stream = new FileStream(@"C:\Users\d.ivanisevic\Documents\Visual Studio 2015\Projects\Learning.Excel\MOCK_DATA.xlsx", FileMode.Open, FileAccess.Read);

            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            
            excelReader.IsFirstRowAsColumnNames = true;
            DataSet result = excelReader.AsDataSet();
            List<Person> persons = new List<Person>();
            DataTable table = new DataTable();
            table = result.Tables[0];
            excelReader.Close();
            excelReader.Dispose();
            stream.Close(); stream.Dispose();
            foreach (DataRow row in table.Rows)
            {
                Person person = new Person()
                {
                    ID = (double)row.ItemArray[0],
                    FirstName = (string)row.ItemArray[1],
                    LastName = (string)row.ItemArray[2],
                    Email = row.ItemArray[3].ToString(),
                    Gender = (string)row.ItemArray[4]
                };
                persons.Add(person);
            }
            return persons;
        }

        public List<Person> GetByName(string name)
        {
            List<Person> result = new List<Person>();
            if (name != null && name != String.Empty) 
            {
                var list = this.GetAll();
                result = list.Where(x => x.FirstName == name).ToList();
            }

            return result;
        }

       public Person GetByID(int id)
        {
            try
            {
                return this.GetAll().Where(x => x.ID == (double)id).FirstOrDefault();
            }
            catch (Exception)
            {

                return null;
            }
        }
    }
}
