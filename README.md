# Excel to Database Integration

This console application demonstrates the integration between an Excel file and a database using EPPlus and Entity Framework Core.

## Project Structure

The project follows a simple structure:

- **ExcelReader** (main project)
  - **Data** (contains data-related classes)
    - `Citizen.cs` (represents the Citizen entity)
  - **DbContexts** (contains Entity Framework DbContext)
    - `CitizenDbContext.cs` (DbContext for interacting with the database)
  - **Excel** (manages reading and writing data to Excel)
    - `ExcelDataManager.cs` (handles Excel-related operations)
  - `Program.cs` (entry point of the application)
  - `UserInput.cs` (manages user input and interaction)

## Features

- Read data from an Excel spreadsheet into a database.
- Create, delete, and seed a database.
- Print data from the database to the console.
- Add new data to the Excel file and see changes reflected in the database.
- Interactive user input for managing data.

## Prerequisites

- [.NET SDK](https://dotnet.microsoft.com/download)
- Excel (xlsx) file containing data

## Installation

1. Clone the repository:

   ```
   git clone https://github.com/your-username/excel-to-database.git	```
   
## Database Setup

### Creating the Database

1. Open a terminal or command prompt.

2. Navigate to the root directory of your project:

3. Run the following command to create the initial migration:
dotnet ef migrations add InitialCreate --project ExcelReader

4. Run the following command to apply the migration and create the database:
dotnet ef database update --project ExcelReader

These commands use the Entity Framework Core CLI (dotnet ef) to manage migrations and apply them to the database. 
Ensure that your project references the necessary NuGet packages, 
and the DbContext is correctly configured in your project.
   
Edit the appsettings.json file in the project root and replace {{Your_Connection_String_Here}} 
with your actual connection string:

##  Important Note
The application needs to be restarted to see changes in the Excel file reflected in the database.

## Contributing
Feel free to contribute by opening issues or submitting pull requests.