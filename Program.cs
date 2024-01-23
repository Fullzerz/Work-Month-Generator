using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Work_Month_Generator
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Declare variables and then initialize to zero.
            int currentYear = 0; int customYear = 0; int errorFlag = 0;

            // Display title
            Console.WriteLine("---------------------------\r");
            Console.WriteLine("Work Month Generator - 2024\r");
            Console.WriteLine("---------------------------\n");

            // Display menu
            Console.WriteLine("Choose an option: \n");
            Console.WriteLine("     1. Create a sheet for the current year\r");
            Console.WriteLine("     2. Create a sheet for a custom year");

            // Use a do while loop and switch statement to select the correct functionality.
            do
            {
                Console.Write("\nYour option? ");
                errorFlag = 0;
                switch (Console.ReadLine())
                {
                    case "1":
                        currentYear = 1;
                        break;
                    case "2":
                        customYear = 1;
                        break;
                    default:
                        Console.WriteLine("This option doesn't exist");
                        errorFlag = 1;
                        break;
                }
            } while (errorFlag == 1);

            // Use an if statement to redirect to requested functionality
            if (currentYear == 1)
            {
                Console.WriteLine($"\ncurrentYear value: {currentYear} -- customYear value: {customYear} -- errorFlag value: {errorFlag}");
            }
            else if (customYear == 1)
            {
                Console.WriteLine($"\ncurrentYear value: {currentYear} -- customYear value: {customYear} -- errorFlag value: {errorFlag}");
            }

            // Wait for the user to respond before closing.
            Console.Write("Press any key to close the Work Month Generator app...");
            Console.ReadKey();
        }
    }
}
