using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader
{
    internal class UserInput
    {
        private ExcelDataManager excelDataManager;

        public UserInput(ExcelDataManager manager)
        {
            excelDataManager = manager;
        }

        public void Start()
        {
            Console.WriteLine("Welcome to the Excel Data Manager!");

            while (true)
            {
                Console.WriteLine("\nChoose an option:");
                Console.WriteLine("1. Add Data to Excel Sheet");
                Console.WriteLine("2. Print Data from Database");
                Console.WriteLine("3. Exit");

                string choice = Console.ReadLine();

                //this switch case is used to choose the option
                switch (choice)
                {
                    case "1":
                        AddDataToExcel();
                        break;

                    case "2":
                        PrintDataFromDatabase();
                        break;

                    case "3":
                        Console.WriteLine("Exiting the program. Goodbye!");
                        Environment.Exit(0);
                        return;

                    default:
                        Console.WriteLine("Invalid choice. Please enter a valid option.");
                        break;
                }
            }
        }

        //this method adds data to excel file
        private void AddDataToExcel()
        {
            Console.WriteLine("Enter new citizen details:");

            Console.Write("Name: ");
            string name = Console.ReadLine();

            int age;
            do
            {
                Console.Write("Age: ");
            } while (!int.TryParse(Console.ReadLine(), out age) || age <= 0);

            Console.Write("City: ");
            string city = Console.ReadLine();

            var newCitizen = new Citizen
            {
                ID = GenerateNewID(),
                Name = name,
                Age = age,
                City = city
            };

            excelDataManager.WriteDataToExcel(new List<Citizen> { newCitizen });

            Console.WriteLine("Data added to the Excel sheet.");
            Start();
        }

        //this method prints data from database
        private void PrintDataFromDatabase()
        {
            Console.WriteLine("\nFetching Data from Database:");

            var fetchedCitizens = excelDataManager.ReadDataFromDatabase();

            if (fetchedCitizens.Any())
            {
                Console.WriteLine("ID\tName\tAge\tCity");
                foreach (var citizen in fetchedCitizens)
                {
                    Console.WriteLine($"{citizen.ID}\t{citizen.Name}\t{citizen.Age}\t{citizen.City}");
                }
            }
            else
            {
                Console.WriteLine("No data found in the database.");
            }
            Start();
        }

        //this method generates new id for the new citizen
        private int GenerateNewID()
        {
            // Get the last ID from the database and increment it
            var lastID = excelDataManager.GetLastIDFromDatabase();
            return lastID + 1;
        }
    }
}

