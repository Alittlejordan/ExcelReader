using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader
{
    internal class ExcelDataManager
    {
        private const string DatabaseFileName = "data.xlsx";

        //this function reads data from excel file and returns a list of citizens
        public List<Citizen> ReadDataFromExcel()
        {
            var assemblyLocation = Assembly.GetExecutingAssembly().Location;
            var projectDirectory = Path.GetDirectoryName(assemblyLocation);

            // Go back four levels
            for (int i = 0; i < 4; i++)
            {
                projectDirectory = Directory.GetParent(projectDirectory)?.FullName;
            }

            var FilePath = Path.Combine(projectDirectory, DatabaseFileName);
            using (var package = new ExcelPackage(new FileInfo(FilePath)))
            {

                var worksheet = package.Workbook.Worksheets.First();

                var startRow = worksheet.Dimension.Start.Row + 1;
                var endRow = worksheet.Dimension.End.Row;

                var people = new List<Citizen>();

                for (int row = startRow; row <= endRow; row++)
                {
                    var person = new Citizen
                    {
                        ID = worksheet.Cells[row, 1].GetValue<int>(),
                        Name = worksheet.Cells[row, 2].GetValue<string>(),
                        Age = worksheet.Cells[row, 3].GetValue<int>(),
                        City = worksheet.Cells[row, 4].GetValue<string>()
                    };

                    people.Add(person);
                }

                return people;
            }
        }

        //this method writes data to excel file

        public void WriteDataToExcel(List<Citizen> citizens)
        {
            var assemblyLocation = Assembly.GetExecutingAssembly().Location;
            var projectDirectory = Path.GetDirectoryName(assemblyLocation);

            // Go back four levels
            for (int i = 0; i < 4; i++)
            {
                projectDirectory = Directory.GetParent(projectDirectory)?.FullName;
            }

            var filePath = Path.Combine(projectDirectory, DatabaseFileName);


            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets.FirstOrDefault() ?? package.Workbook.Worksheets.Add("Sheet1");

                // Find the last used row in the worksheet
                int lastUsedRow = worksheet.Dimension?.End.Row ?? 1;

                // Append new data to the worksheet
                for (int i = 0; i < citizens.Count; i++)
                {
                    var citizen = citizens[i];

                    worksheet.Cells[lastUsedRow + i + 1, 1].Value = citizen.ID;
                    worksheet.Cells[lastUsedRow + i + 1, 2].Value = citizen.Name;
                    worksheet.Cells[lastUsedRow + i + 1, 3].Value = citizen.Age;
                    worksheet.Cells[lastUsedRow + i + 1, 4].Value = citizen.City;
                }

                package.Save();
            }
        }

        //this method intializes the database and deletes the existing database (if one exist) and creates a new one
        public void InitializeDatabase()
        {
            using (var context = new CitizenDbContext())
            {
                if (context.Database.CanConnect())
                {
                    Console.WriteLine("Deleting existing database...");

                    // Delete the existing database if it exists
                    context.Database.EnsureDeleted();
                }

                Console.WriteLine("Creating a new database...");

                // Create a new database
                context.Database.EnsureCreated();

                // No need to seed data from Excel, as we are fetching from the database
            }
        }

        //this method writes data to database
        public void WriteDataToDatabase(List<Citizen> citizens)
        {
            using (var context = new CitizenDbContext())
            {
                context.Database.OpenConnection();
                try
                {
                    // Enable IDENTITY_INSERT so that we can insert data into the ID column
                    context.Database.ExecuteSqlRaw("SET IDENTITY_INSERT Citizens ON");

                    context.Citizens.AddRange(citizens);
                    context.SaveChanges();
                }
                finally
                {
                    
                    context.Database.ExecuteSqlRaw("SET IDENTITY_INSERT Citizens OFF");
                    context.Database.CloseConnection();
                }
            }
        }

        //this method reads data from database
        public List<Citizen> ReadDataFromDatabase()
        {
            using (var context = new CitizenDbContext())
            {
                return context.GetAllCitizens();
            }
        }

        //this method returns the last id from the database
        public int GetLastIDFromDatabase()
        {
            using (var context = new CitizenDbContext())
            {
                var lastID = context.Citizens.Max(c => (int?)c.ID) ?? 0;
                return lastID;
            }
        }
    }
}
