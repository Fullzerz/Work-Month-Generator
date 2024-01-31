using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using NanoXLSX;
using NanoXLSX.Styles;
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
            String rif = "";
            String job = "";
            String[] employees = null;
            String[] sheetTitles = null;

            // Retrieve configuration files txt
            job = readJob();
            rif = readRif();
            employees = readEmployees();
            sheetTitles = readSheetTitles();

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
            generateExcel(yearValue, selectedMonth, holidayFlag, holidayDays, job, rif, employees, sheetTitles);

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

        public static void generateExcel(String yearValue, int selectedMonth, bool holidayFlag, String[]holidayDays, String job, String rif, String[] employees, String[] sheetTitles)
        {

            Workbook workbook = new Workbook($"{yearValue}_{selectedMonth.ToString("00")}.xlsx", sheetTitles[0]);
            Worksheet presenzeSheet = new Worksheet(sheetTitles[1]);
            int daysInMonth = DateTime.DaysInMonth(Int32.Parse(yearValue), selectedMonth);

            // RIF COLUMN
            presenzeSheet.AddNextCell($"Rif.{job}");
            Address address1 = new Address(0, 0);
            Address address2 = new Address(0, 1);

            Range range1 = new Range(address1, address2);
            presenzeSheet.MergeCells(range1);
            // END RIF COLUMN

            // RISORSA COLUMN
            presenzeSheet.AddNextCell("Risorsa");
            Address address3 = new Address(1, 0);
            Address address4 = new Address(1, 1);

            Range range2 = new Range(address3, address4);
            presenzeSheet.MergeCells(range2);
            // END RISORSA COLUMN 

            // MONTH COLUMN
            switch (selectedMonth){
                case 1:
                    presenzeSheet.AddNextCell($"Jan-{yearValue}");
                    break;
                case 2:
                    presenzeSheet.AddNextCell($"Feb-{yearValue}");
                    break;
                case 3:
                    presenzeSheet.AddNextCell($"Mar-{yearValue}");
                    break;
                case 4:
                    presenzeSheet.AddNextCell($"Apr-{yearValue}");
                    break;
                case 5:
                    presenzeSheet.AddNextCell($"May-{yearValue}");
                    break;
                case 6:
                    presenzeSheet.AddNextCell($"Jun-{yearValue}");
                    break;
                case 7:
                    presenzeSheet.AddNextCell($"Jul-{yearValue}");
                    break;
                case 8:
                    presenzeSheet.AddNextCell($"Aug-{yearValue}");
                    break;
                case 9:
                    presenzeSheet.AddNextCell($"Sep-{yearValue}");
                    break;
                case 10:
                    presenzeSheet.AddNextCell($"Oct-{yearValue}");
                    break;
                case 11:
                    presenzeSheet.AddNextCell($"Nov-{yearValue}");
                    break;
                case 12:
                    presenzeSheet.AddNextCell($"Dec-{yearValue}");
                    break;
            }
            // END MONTH COLUMN

            // TOTALE COLUMN
            presenzeSheet.AddCell("totale gg/u", daysInMonth+3, 0);
            // END TOTALE COLUMN

            // MERGE FIRST ROW
            Address address5 = new Address(2, 0);
            Address address6 = new Address(daysInMonth+2, 0);

            Range range3 = new Range(address5, address6);
            presenzeSheet.MergeCells(range3);
            // END MERGE FIRST ROW

            // ---- STYLES ----
            Style style = new Style();
            style.CurrentCellXf.TextRotation = 90;
            style.CurrentCellXf.Locked = false;
            //style.CurrentCellXf.Alignment = (CellXf.TextBreakValue)CellXf.HorizontalAlignValue.left;
            // ---- END OF STYLES ---
            //workbook.SetWorkbookProtection(true, false, false, null);

            // PRINT DAYS
            presenzeSheet.GoToNextRow();
            presenzeSheet.AddNextCell("");
            presenzeSheet.AddNextCell("");
            
            for(int i = 0; i < daysInMonth; i++)
            {
                presenzeSheet.AddNextCell($"{(i+1).ToString("00")}/{selectedMonth.ToString("00")}");
                
                Address address = new Address(presenzeSheet.GetCurrentColumnNumber()-2, presenzeSheet.GetCurrentRowNumber());
                Cell cell = presenzeSheet.GetCell(address);
                cell.SetStyle(style);
            }
            presenzeSheet.AddNextCell("totale gg/u");
            presenzeSheet.AddNextCell("inserire i valori in GG");
            // END PRINT DAYS

            // PRINT EMPLOYEES
            presenzeSheet.GoToNextRow();
            foreach(String employee in employees)
            {
                presenzeSheet.AddNextCell(rif);
                presenzeSheet.AddNextCell(employee);
                presenzeSheet.GoToNextRow();
            }
            presenzeSheet.AddNextCell("totale Giorni");
            // END PRINT EMPLOYEES

            // Add the second sheet to the workbook and save
            workbook.AddWorksheet(presenzeSheet);
            workbook.Save();

            Console.WriteLine($"\n--> Excel file generated.");
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

        public static String[] readSheetTitles()
        {
            var sheetTitlesFile = Path.Combine(Directory.GetCurrentDirectory(), "sheetTitles.txt");
            List<String> sheetTitles = new List<String>();

            try
            {
                // Check if file exists.
                if (!File.Exists(sheetTitlesFile))
                {
                    Console.WriteLine("--> sheetTitles.txt file does not exist. It will be automatically generated. Please fill it with correct data.");
                    FileStream fs = File.Create(sheetTitlesFile);
                    closeProgram();
                }
                else
                {
                    // Check if file empty.
                    if (new FileInfo(sheetTitlesFile).Length == 0)
                    {
                        Console.WriteLine("--> sheetTitles.txt file is empty. Please fill it with correct data.");
                        closeProgram();
                    }
                    else
                    {
                        // Open the stream and read it back.
                        using (StreamReader sr = File.OpenText(sheetTitlesFile))
                        {
                            string s = "";
                            while ((s = sr.ReadLine()) != null)
                            {
                                sheetTitles.Add(s);
                            }
                            return sheetTitles.ToArray();
                        }
                    }
                }
            }
            catch (Exception)
            {
                Console.WriteLine("--> An error has occurred when generating sheetTitles file. Please exit from the application.");
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
