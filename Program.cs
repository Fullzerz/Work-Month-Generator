using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Http;
using System.Threading.Tasks;
using Newtonsoft.Json; // Nuget Package

namespace Work_Month_Generator
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            // Declare variables and then initialize them.
            bool currentYear = false; 
            bool customYear = false; 
            bool errorFlag = false;
            bool holidayFlag = false;
            String yearValue = "";
            String localeValue = "IT";
            String apiUrl = "https://date.nager.at/api/v3/publicholidays";
            List<HolidayJson> holidayList = null;

            // Display title
            Console.WriteLine("-----------------------------------\r");
            Console.WriteLine("Work Month Generator - 2024 - ITALY\r");
            Console.WriteLine("-----------------------------------\n");

            // Display menu
            Console.WriteLine("Choose an option: \n");
            Console.WriteLine("     1. Create a sheet for the current year\r");
            Console.WriteLine("     2. Create a sheet for a custom year");

            // Use a do while loop and switch statement to select the correct functionality.
            do
            {
                Console.Write("\nYour option? ");
                errorFlag = false;
                switch (Console.ReadLine())
                {
                    case "1":
                        currentYear = true;
                        break;
                    case "2":
                        customYear = true;
                        break;
                    default:
                        Console.WriteLine("This option doesn't exist");
                        errorFlag = true;
                        break;
                }
            } while (errorFlag);

            // Use a do while loop and switch statement to let the user decide if they want to include holidays (call to API).
            do
            {
                Console.Write("\nDo you want to include holidays in this sheet? (Y/N): ");
                errorFlag = false;
                switch (Console.ReadLine())
                {
                    case "Y": case "y":
                        holidayFlag = true;
                        break;
                    case "N": case "n":
                        holidayFlag = false;
                        break;
                    default:
                        Console.WriteLine("This option doesn't exist");
                        errorFlag = true;
                        break;
                }
            } while (errorFlag);

            // Use an if statement to redirect to requested functionality
            if (currentYear)
            {
                yearValue = DateTime.Now.Year.ToString();

                if (holidayFlag)
                {
                    holidayList = await retrieveHolidaysAPI($"{apiUrl}/{yearValue}/{localeValue}");
                    foreach(var holiday in holidayList){
                        Console.Write(holiday.date + "--" + holiday.localName + "\n");
                    }
                    // TODO: HERE GOES THE CODE THAT WILL DECODE THE JSON
                }

                // TODO: HERE GOES THE CODE THAT WILL GENERATE THE EXCEL
            }
            else if (customYear)
            {
                do
                {
                    Console.Write("\nInput a valid year: ");
                    yearValue = Console.ReadLine();

                    var inputNumber = 0;
                    var parsingFlag = Int32.TryParse(yearValue, out inputNumber);
                    if ((inputNumber < 1000) || (inputNumber > 9999) || !parsingFlag)
                    {
                        Console.WriteLine("Invalid entry.");
                        errorFlag = true;
                    } else
                    {
                        errorFlag = false;
                    }
                } while (errorFlag);
                


                if (holidayFlag)
                {
                    holidayList = await retrieveHolidaysAPI($"{apiUrl}/{yearValue}/{localeValue}");
                    foreach (var holiday in holidayList)
                    {
                        Console.Write(holiday.date + "--" + holiday.localName + "\n");
                    }
                    // TODO: HERE GOES THE CODE THAT WILL DECODE THE JSON
                }

                // TODO: HERE GOES THE CODE THAT WILL GENERATE THE EXCEL
            }

            // Wait for the user to respond before closing.
            Console.Write("Press any key to close the Work Month Generator app...");
            Console.ReadLine(); //TODO: Rimettere ReadKey
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
            catch (Exception ex)
            {
                Console.WriteLine("An error has occurred when calling the API to retrieve the holidays. Exception: " + ex);
            }
            return null;
        }
    }
}
