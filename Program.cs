using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using NanoXLSX;
using Newtonsoft.Json; // Nuget Package

namespace Work_Month_Generator
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            // Declare variables and then initialize them.
            bool isCustomYear = false; 
            bool errorFlag = false;
            bool holidayFlag = false;
            int selectedMonth = 0;
            String yearValue = "";
            String localeValue = "";
            String apiUrl = "https://date.nager.at/api/v3/publicholidays";
            List<HolidayJson> holidayList;
            String[] holidayDays = null;
            String rifSog = "";
            String job = "";
            String[] employees = null;

            // Retrieve configuration files txt
            job = readJob();
            rifSog = readRif();
            employees = readEmployees();

            // Retrieve all data from user
            isCustomYear = displayTitleAndMenu(errorFlag);
            yearValue = selectYear(isCustomYear, errorFlag);
            selectedMonth = selectMonth(errorFlag);
            holidayFlag = selectHolidays(errorFlag);
            localeValue = selectLocale(errorFlag, holidayFlag);

            // Retrieve holidays if needed
            if (holidayFlag)
            {
                holidayList = await retrieveHolidaysAPI($"{apiUrl}/{yearValue}/{localeValue}");
                holidayDays = extractHolidays(holidayList);
            }

            // Generate Excel and save it on the computer
            generateExcel(yearValue, selectedMonth, holidayFlag, holidayDays);

            // Close program
            closeProgram(); 
        }

        public static async Task<List<HolidayJson>> retrieveHolidaysAPI(string url)
        {
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    var response = await client.GetAsync(url);
                    if (response != null)
                    {
                        var jsonString = await response.Content.ReadAsStringAsync();
                        return JsonConvert.DeserializeObject<List<HolidayJson>>(jsonString);
                    }
                }
            }
            catch (Exception)
            {
                Console.WriteLine("--> REQUEST ERROR: An error has occurred when calling the API to retrieve the holidays.");
                closeProgram();
            }
            return null;
        }

        public static bool displayTitleAndMenu(bool errorFlag)
        {
            // Display title
            Console.WriteLine("---------------------------\r");
            Console.WriteLine("Work Month Generator - 2024\r");
            Console.WriteLine("---------------------------\n");

            // Display menu
            Console.WriteLine("--> Choose an option: \n");
            Console.WriteLine("     1. Create a sheet for the current year\r");
            Console.WriteLine("     2. Create a sheet for a custom year");

            // Use a do while loop and switch statement to let the user decide if they want to create a sheet for current or custom year.
            do
            {
                Console.Write("\n--> Your option? ");
                errorFlag = false;
                switch (Console.ReadLine())
                {
                    case "1":
                        return false;
                    case "2":
                        return true;
                    default:
                        Console.WriteLine("This option doesn't exist");
                        errorFlag = true;
                        break;
                }
            } while (errorFlag);
            return false;
        }

        public static String selectYear(bool isCustomYear, bool errorFlag)
        {
            // Use an if statement to redirect to requested functionality
            if (!isCustomYear)
            {
                return DateTime.Now.Year.ToString();
            }
            else if (isCustomYear)
            {
                do
                {
                    Console.Write("\n--> Input a valid year: ");
                    String yearValue = Console.ReadLine();

                    var inputNumber = 0;
                    var parsingFlag = Int32.TryParse(yearValue, out inputNumber);
                    if ((inputNumber < 1000) || (inputNumber > 9999) || !parsingFlag)
                    {
                        Console.WriteLine("Invalid entry.");
                        errorFlag = true;
                    }
                    else
                    {
                        errorFlag = false;
                        return yearValue;
                    }
                } while (errorFlag);
            }
            return "";
        }

        public static int selectMonth(bool errorFlag)
        {
            // Use a do while loop and switch statement to let the user decide the month to be generated.
            Console.WriteLine("\n--> Select a month: \n");
            Console.WriteLine("     1. Gennaio\r");
            Console.WriteLine("     2. Febbraio\r");
            Console.WriteLine("     3. Marzo\r");
            Console.WriteLine("     4. Aprile\r");
            Console.WriteLine("     5. Maggio\r");
            Console.WriteLine("     6. Giugno\r");
            Console.WriteLine("     7. Luglio\r");
            Console.WriteLine("     8. Agosto\r");
            Console.WriteLine("     9. Settembre\r");
            Console.WriteLine("     10. Ottobre\r");
            Console.WriteLine("     11. Novembre\r");
            Console.WriteLine("     12. Dicembre\r");

            do
            {
                Console.Write("\n--> Your option? ");
                String selectedMonth = Console.ReadLine();
                errorFlag = false;

                var inputNumber = 0;
                var parsingFlag = Int32.TryParse(selectedMonth, out inputNumber);
                if ((inputNumber < 1) || (inputNumber > 12) || !parsingFlag)
                {
                    Console.WriteLine("Invalid entry.");
                    errorFlag = true;
                }
                else
                {
                    return inputNumber;
                }
            } while (errorFlag);
            return 0;
        }

        public static bool selectHolidays(bool errorFlag)
        {
            // Use a do while loop and switch statement to let the user decide if they want to include holidays (call to API).
            do
            {
                Console.Write("\n--> Do you want to include holidays in this sheet? (Y/N): ");
                errorFlag = false;
                switch (Console.ReadLine())
                {
                    case "Y":
                    case "y":
                        return true;
                    case "N":
                    case "n":
                        return false;
                    default:
                        Console.WriteLine("This option doesn't exist");
                        errorFlag = true;
                        break;
                }
            } while (errorFlag);
            return false;
        }

        public static String selectLocale(bool errorFlag, bool holidayFlag)
        {
            if (holidayFlag)
            {
                Console.Write("\n--> Please enter your locale ID (Italy=IT, United States=US etc.): ");
                return Console.ReadLine();  
            } else
            {
                return "";
            }         
        }

        public static void closeProgram()
        {
            // Wait for the user to respond before closing.
            Console.Write("\n--> Press any key to close the Work Month Generator app...");
            Console.ReadKey();
            Environment.Exit(0);
        }

        public static String[] extractHolidays(List<HolidayJson>holidayList)
        {
            bool isNullOrEmpty = holidayList?.Any() != true;
            List<string> dates = new List<string>();

            if (!isNullOrEmpty)
            {
                foreach (var holiday in holidayList)
                {
                    //Console.Write(holiday.date + "--" + holiday.localName + "\n");
                    dates.Add(holiday.date);
                }
                return dates.ToArray();
            }
            else
            {
                Console.WriteLine("--> LIST EMPTY ERROR: List of holidays is empty. Check locale ID.");
                closeProgram();
                return null;
            }
        }

        public static void generateExcel(String yearValue, int selectedMonth, bool holidayFlag, String[]holidayDays)
        {
            //Console.WriteLine($"FULL: {yearValue}/{selectedMonth} -- Holiday: {holidayFlag}");

            Workbook workbook = new Workbook("myWorkbook.xlsx", "Sheet1");         // Create new workbook with a worksheet called Sheet1
            workbook.CurrentWorksheet.AddNextCell("Some Data");                    // Add cell A1
            workbook.CurrentWorksheet.AddNextCell(42);                             // Add cell B1
            workbook.CurrentWorksheet.GoToNextRow();                               // Go to row 2
            workbook.CurrentWorksheet.AddNextCell(DateTime.Now);                   // Add cell A2
            workbook.Save();                                                       // Save the workbook as myWorkbook.xlsx
        }

        public static String readRif()
        {
            var rif = Path.Combine(Directory.GetCurrentDirectory(), "rif.txt");

            try
            {
                // Check if file exists.
                if (!File.Exists(rif))
                {
                    Console.WriteLine("--> rif.txt file does not exist. It will be automatically generated. Please fill it with correct data.");
                    FileStream fs = File.Create(rif);
                    closeProgram();
                }
                else
                {
                    // Check if file empty.
                    if (new FileInfo(rif).Length == 0)
                    {
                        Console.WriteLine("--> rif.txt file is empty. Please fill it with correct data.");
                        closeProgram();
                    }
                    else
                    {
                        // Open the stream and read it back.
                        using (StreamReader sr = File.OpenText(rif))
                        {
                            string s = "";
                            while ((s = sr.ReadLine()) != null)
                            {
                                return s;
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                Console.WriteLine("--> An error has occurred when generating rif file. Please exit from the application.");
                closeProgram();
            }
            return "";
        }

        public static String[] readEmployees()
        {
            var employeesListFile = Path.Combine(Directory.GetCurrentDirectory(), "employeesList.txt");
            List<String> employeesList = new List<String>();

            try
            {
                // Check if file exists.
                if (!File.Exists(employeesListFile))
                {
                    Console.WriteLine("--> employeesList.txt file does not exist. It will be automatically generated. Please fill it with correct data.");
                    FileStream fs = File.Create(employeesListFile);
                    closeProgram();
                }
                else
                {
                    // Check if file empty.
                    if (new FileInfo(employeesListFile).Length == 0)
                    {
                        Console.WriteLine("--> employeesList.txt file is empty. Please fill it with correct data.");
                        closeProgram();
                    }
                    else
                    {
                        // Open the stream and read it back.
                        using (StreamReader sr = File.OpenText(employeesListFile))
                        {
                            string s = "";
                            while ((s = sr.ReadLine()) != null)
                            {
                                employeesList.Add(s);
                            }
                            return employeesList.ToArray();
                        }
                    }
                }
            }
            catch (Exception)
            {
                Console.WriteLine("--> An error has occurred when generating employeesList file. Please exit from the application.");
                closeProgram();
            }
            return null;
        }

        public static String readJob()
        {
            var job = Path.Combine(Directory.GetCurrentDirectory(), "job.txt");

            try
            {
                // Check if file exists.
                if (!File.Exists(job))
                {
                    Console.WriteLine("--> job.txt file does not exist. It will be automatically generated. Please fill it with correct data.");
                    FileStream fs = File.Create(job);
                    closeProgram();
                }
                else
                {
                    // Check if file empty.
                    if (new FileInfo(job).Length == 0)
                    {
                        Console.WriteLine("--> job.txt file is empty. Please fill it with correct data.");
                        closeProgram();
                    }
                    else
                    {
                        // Open the stream and read it back.
                        using (StreamReader sr = File.OpenText(job))
                        {
                            string s = "";
                            while ((s = sr.ReadLine()) != null)
                            {
                                return s;
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                Console.WriteLine("--> An error has occurred when generating rif file. Please exit from the application.");
                closeProgram();
            }
            return "";
        }
    }
}
