using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Globalization;

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
            }
            else
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

        public static String[] extractHolidays(List<HolidayJson> holidayList)
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

        public static void generateExcel(String yearValue, int selectedMonth, bool holidayFlag, String[] holidayDays, String job, String rif, String[] employees, String[] sheetTitles)
        {
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;
            Excel.Range oRng;

            try
            {
                //Start Excel and get Application object.
                oXL = new Excel.Application();
                oXL.Visible = false;

                //Initialize variables
                int daysInMonth = DateTime.DaysInMonth(Int32.Parse(yearValue), selectedMonth);
                int count = 0;

                //Get a new workbook and create sheets
                oWB = oXL.Workbooks.Add(Missing.Value);
                oSheet = oWB.ActiveSheet;
                oSheet.Name = $"{sheetTitles[1]}";

                //Add table headers going cell by cell.
                oSheet.Cells[1, 1] = $"Rif. {job}";
                oSheet.Cells[1, 2] = "Risorsa";

                //Write selected month into worksheet
                oSheet.Cells[1, 3] = $"{GetMonthName(selectedMonth)}-{yearValue}";
                oSheet.Range[oSheet.Cells[1, 3], oSheet.Cells[1, 3]].NumberFormat = "mmm-yyyy";

                //Styles for the first row
                oSheet.Range[oSheet.Cells[1, 1], oSheet.Cells[1, 3]].Font.Bold = true;
                oSheet.Range[oSheet.Cells[1, 1], oSheet.Cells[1, 3]].Font.Size = 9;
                oSheet.Range[oSheet.Cells[1, 1], oSheet.Cells[1, 3]].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(-4194304));
                oSheet.Range[oSheet.Cells[1, 1], oSheet.Cells[1, 3]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(-855310));
                oSheet.Range[oSheet.Cells[1, 1], oSheet.Cells[1, 3]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                oSheet.Range[oSheet.Cells[1, 1], oSheet.Cells[1, 3]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                // Merge cells and write last cell
                oSheet.get_Range("A1", "A2").MergeCells = true;
                oSheet.get_Range("B1", "B2").MergeCells = true;
                oSheet.Range[oSheet.Cells[1, 3], oSheet.Cells[1, daysInMonth + 3]].MergeCells = true;
                oSheet.Cells[1, daysInMonth + 4] = "Totale gg/u";
                oSheet.Range[oSheet.Cells[1, daysInMonth + 4], oSheet.Cells[1, daysInMonth + 4]].Font.Bold = true;
                oSheet.Range[oSheet.Cells[1, daysInMonth + 4], oSheet.Cells[1, daysInMonth + 4]].Font.Size = 9;
                oSheet.Range[oSheet.Cells[1, daysInMonth + 4], oSheet.Cells[1, daysInMonth + 4]].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(-4194304));
                oSheet.Range[oSheet.Cells[1, daysInMonth + 4], oSheet.Cells[1, daysInMonth + 4]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(-855310));
                oSheet.Range[oSheet.Cells[1, daysInMonth + 4], oSheet.Cells[1, daysInMonth + 4]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                oSheet.Range[oSheet.Cells[1, daysInMonth + 4], oSheet.Cells[1, daysInMonth + 4]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                // Print days
                for (int i = 0; i < daysInMonth; i++)
                {
                    oSheet.Cells[2, i + 3] = $"{(i + 1).ToString("00")}-{GetMonthName(selectedMonth)}-{yearValue}";
                    oSheet.Range[oSheet.Cells[2, i + 3], oSheet.Cells[2, i + 3]].NumberFormat = "dd/mm";
                    oSheet.Range[oSheet.Cells[2, i + 3], oSheet.Cells[2, i + 3]].Cells.Orientation = Excel.XlOrientation.xlUpward;
                    oSheet.Range[oSheet.Cells[2, i + 3], oSheet.Cells[2, i + 3]].Font.Size = 9;
                    oSheet.Range[oSheet.Cells[2, i + 3], oSheet.Cells[2, i + 3]].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(-4194304));
                    oSheet.Range[oSheet.Cells[2, i + 3], oSheet.Cells[2, i + 3]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(-855310));
                    oSheet.Range[oSheet.Cells[2, i + 3], oSheet.Cells[2, i + 3]].ColumnWidth = 5.71;
                }
                oSheet.Cells[2, daysInMonth + 3] = "totale gg/u";
                oSheet.Range[oSheet.Cells[2, daysInMonth + 3], oSheet.Cells[2, daysInMonth + 3]].Font.Bold = true;
                oSheet.Range[oSheet.Cells[2, daysInMonth + 3], oSheet.Cells[2, daysInMonth + 3]].Font.Size = 9;
                oSheet.Range[oSheet.Cells[2, daysInMonth + 3], oSheet.Cells[2, daysInMonth + 3]].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(-4194304));
                oSheet.Range[oSheet.Cells[2, daysInMonth + 3], oSheet.Cells[2, daysInMonth + 3]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(-855310));
                oSheet.Range[oSheet.Cells[2, daysInMonth + 3], oSheet.Cells[2, daysInMonth + 3]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                oSheet.Range[oSheet.Cells[2, daysInMonth + 3], oSheet.Cells[2, daysInMonth + 3]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                oSheet.Range[oSheet.Cells[2, daysInMonth + 3], oSheet.Cells[2, daysInMonth + 3]].WrapText = true;
                oSheet.Range[oSheet.Cells[2, daysInMonth + 3], oSheet.Cells[2, daysInMonth + 3]].ColumnWidth = 5.71;

                oSheet.Cells[2, daysInMonth + 4] = "inserire i valori";
                oSheet.Range[oSheet.Cells[2, daysInMonth + 4], oSheet.Cells[2, daysInMonth + 4]].Font.Bold = true;
                oSheet.Range[oSheet.Cells[2, daysInMonth + 4], oSheet.Cells[2, daysInMonth + 4]].Font.Size = 11;
                oSheet.Range[oSheet.Cells[2, daysInMonth + 4], oSheet.Cells[2, daysInMonth + 4]].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(-16777216));
                oSheet.Range[oSheet.Cells[2, daysInMonth + 4], oSheet.Cells[2, daysInMonth + 4]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(-256));
                oSheet.Range[oSheet.Cells[2, daysInMonth + 4], oSheet.Cells[2, daysInMonth + 4]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                oSheet.Range[oSheet.Cells[2, daysInMonth + 4], oSheet.Cells[2, daysInMonth + 4]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                oSheet.Range[oSheet.Cells[2, daysInMonth + 4], oSheet.Cells[2, daysInMonth + 4]].WrapText = true;
                oSheet.Range[oSheet.Cells[2, daysInMonth + 4], oSheet.Cells[2, daysInMonth + 4]].Borders.LineStyle = Excel.XlLineStyle.xlDouble;

                // Print employees and red column for total hours
                count = 1;
                foreach (String employee in employees)
                {
                    oSheet.Cells[count + 2, 1] = rif;
                    oSheet.Cells[count + 2, 2] = employee;
                    oSheet.Range[oSheet.Cells[count + 2, 1], oSheet.Cells[count + 2, 2]].Font.Size = 9;

                    oSheet.Cells[count + 2, daysInMonth + 3].Formula = "=Sum(" + oSheet.Cells[count + 2, 3].Address + ":" + oSheet.Cells[count + 2, daysInMonth + 2].Address + ")";
                    oSheet.Range[oSheet.Cells[count + 2, daysInMonth + 3], oSheet.Cells[count + 2, daysInMonth + 3]].Font.Size = 9;
                    oSheet.Range[oSheet.Cells[count + 2, daysInMonth + 3], oSheet.Cells[count + 2, daysInMonth + 3]].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(-1));
                    oSheet.Range[oSheet.Cells[count + 2, daysInMonth + 3], oSheet.Cells[count + 2, daysInMonth + 3]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(-4194304));
                    oSheet.Range[oSheet.Cells[count + 2, daysInMonth + 3], oSheet.Cells[count + 2, daysInMonth + 3]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    oSheet.Range[oSheet.Cells[count + 2, daysInMonth + 3], oSheet.Cells[count + 2, daysInMonth + 3]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                    count++;
                }
                oSheet.Cells[count + 2, daysInMonth + 3].Formula = "=Sum(" + oSheet.Cells[count + 2, 3].Address + ":" + oSheet.Cells[count + 2, daysInMonth + 2].Address + ")";
                oSheet.Range[oSheet.Cells[count + 2, daysInMonth + 3], oSheet.Cells[count + 2, daysInMonth + 3]].Font.Size = 9;
                oSheet.Range[oSheet.Cells[count + 2, daysInMonth + 3], oSheet.Cells[count + 2, daysInMonth + 3]].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(-1));
                oSheet.Range[oSheet.Cells[count + 2, daysInMonth + 3], oSheet.Cells[count + 2, daysInMonth + 3]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(-4194304));
                oSheet.Range[oSheet.Cells[count + 2, daysInMonth + 3], oSheet.Cells[count + 2, daysInMonth + 3]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                oSheet.Range[oSheet.Cells[count + 2, daysInMonth + 3], oSheet.Cells[count + 2, daysInMonth + 3]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                oSheet.Cells[count + 2, 1] = "totale Giorni";
                oSheet.Range[oSheet.Cells[count + 2, 1], oSheet.Cells[count + 2, 2]].MergeCells = true;
                oSheet.Range[oSheet.Cells[count + 2, 1], oSheet.Cells[count + 2, 2]].Font.Bold = true;
                oSheet.Range[oSheet.Cells[count + 2, 1], oSheet.Cells[count + 2, 2]].Font.Size = 9;
                oSheet.Range[oSheet.Cells[count + 2, 1], oSheet.Cells[count + 2, 2]].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(-4194304));
                oSheet.Range[oSheet.Cells[count + 2, 1], oSheet.Cells[count + 2, 2]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(-855310));
                oSheet.Range[oSheet.Cells[count + 2, 1], oSheet.Cells[count + 2, 2]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                oSheet.Range[oSheet.Cells[count + 2, 1], oSheet.Cells[count + 2, 2]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                // Auto fit first 2 columns
                oRng = oSheet.get_Range("A1", "B1");
                oRng.EntireColumn.AutoFit();

                // Create borders for all written cells, unlock all written cells
                Excel.Range last = oSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                oSheet.Range["A1", last].Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                oSheet.Range["A1", last].Cells.Locked = false;

                // Remove borders from unnecessary cells
                count = 1;
                foreach (String employee in employees)
                {
                    oSheet.Range[oSheet.Cells[count + 2, daysInMonth + 4], oSheet.Cells[count + 2, daysInMonth + 4]].Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                    count++;
                }
                oSheet.Range[oSheet.Cells[count + 2, daysInMonth + 4], oSheet.Cells[count + 2, daysInMonth + 4]].Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;

                // Holidays and weekends + lock cells
                for (int i = 0; i < daysInMonth; i++)
                {
                    String date = oSheet.Cells[2, i + 3].Value.ToString();
                    //Console.WriteLine($"--> TEST: {date}");
                    var cultureInfo = new CultureInfo("it-IT");
                    DateTime dateTime = DateTime.Parse(date, cultureInfo);
                    if ((dateTime.DayOfWeek == DayOfWeek.Saturday) || (dateTime.DayOfWeek == DayOfWeek.Sunday))
                    {
                        count = 1;
                        //Console.WriteLine("This is a weekend");
                        foreach (String employee in employees)
                        {
                            oSheet.Range[oSheet.Cells[count + 2, i + 3], oSheet.Cells[count + 2, i + 3]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(-4210753));
                            oSheet.Range[oSheet.Cells[count + 2, i + 3], oSheet.Cells[count + 2, i + 3]].Cells.Locked = true;
                            count++;
                        }
                        oSheet.Range[oSheet.Cells[count + 2, i + 3], oSheet.Cells[count + 2, i + 3]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(-4210753));
                        oSheet.Range[oSheet.Cells[count + 2, i + 3], oSheet.Cells[count + 2, i + 3]].Cells.Locked = true;
                    }

                    if (holidayFlag)
                    {
                        foreach (String holiday in holidayDays)
                        {
                            DateTime holidayDateTime = DateTime.Parse(holiday, cultureInfo);
                            if (DateTime.Compare(holidayDateTime, dateTime) == 0)
                            {
                                count = 1;
                                //Console.WriteLine("This is a holiday");
                                foreach (String employee in employees)
                                {
                                    oSheet.Range[oSheet.Cells[count + 2, i + 3], oSheet.Cells[count + 2, i + 3]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(-4210753));
                                    oSheet.Range[oSheet.Cells[count + 2, i + 3], oSheet.Cells[count + 2, i + 3]].Cells.Locked = true;
                                    count++;
                                }
                                oSheet.Range[oSheet.Cells[count + 2, i + 3], oSheet.Cells[count + 2, i + 3]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(-4210753));
                                oSheet.Range[oSheet.Cells[count + 2, i + 3], oSheet.Cells[count + 2, i + 3]].Cells.Locked = true;
                            }
                        }
                    }
                }

                //Protect the sheet
                oSheet.Protect();

                //Create a copy of the current sheet, using another title
                oSheet.Copy(oSheet);
                oSheet = oWB.ActiveSheet;
                oSheet.Name = $"{sheetTitles[0]}";

                //Make sure Excel is visible and give the user the save prompt
                oXL.Visible = true;
                oXL.UserControl = true;
                oXL.GetSaveAsFilename($"{rif}-{yearValue}_{selectedMonth.ToString("00")}");
                Console.WriteLine($"\n--> Excel file generated.");
            }
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);
                Console.WriteLine($"\n--> A fatal error has occurred: {errorMessage}");
            }
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

        public static String GetMonthName(int month)
        {
            return CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(month);
        }
    }
}
