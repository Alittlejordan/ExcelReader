using OfficeOpenXml;

namespace ExcelReader
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Initializing Database...");
            // Set the license context
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelDataManager excelDataManager = new ExcelDataManager();

            try
            {
                // Initialize the database by deleting existing, creating new, and populating with data
                excelDataManager.InitializeDatabase();

                Console.WriteLine("Database Initialized Successfully!");

                // Read data from Excel and populate the database
                List<Citizen> citizens = excelDataManager.ReadDataFromExcel();
                excelDataManager.WriteDataToDatabase(citizens);

                Console.WriteLine("Data loaded into the database!");

                 UserInput userInput = new UserInput(excelDataManager);
                 userInput.Start();

                
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error Initializing/Populating Database: {ex.Message}");
            }

        }
    }
}