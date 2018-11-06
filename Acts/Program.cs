using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Acts
{
    class Program
    {
        // 0 - path to template
        // 1 - path to Excel
        // 2 - path to Reasons' file
        static void Main(string[] args)
        {
            Console.WriteLine("Welcome to Act Genereator!!!");
            Console.WriteLine();

            string[] correctArgs = new string[3];
            if (args.Length > 0)
            {
                correctArgs = SetArgs(args);
            }
            else
            {
                Console.WriteLine("Please, enter the path to Template:");
                correctArgs[0] = Console.ReadLine();
                Console.WriteLine("Thanks. And now, please, enter the path to Values:");
                correctArgs[1] = Console.ReadLine();
                Console.WriteLine("Thanks. And one more: please, enter the path to Reasons:");
                correctArgs[2] = Console.ReadLine();
            }

            CheckArgs(ref correctArgs);
            new Docs(correctArgs[0], correctArgs[1], correctArgs[2]).Execute();
        }

        
        private static string[] SetArgs (string[] args)
        {
            string[] correctArgs = new string[3];

            switch (args.Length)
            {
                case 1:
                    correctArgs[0] = args[0];
                    correctArgs[1] = "C:\\temp\\Values.xlsx"; //default
                    correctArgs[2] = "C:\\temp\\Reasons.xlsx"; //default
                    return correctArgs;
                case 2:
                    correctArgs[0] = args[0];
                    correctArgs[1] = args[1];
                    correctArgs[2] = "C:\\temp\\Reasons.xlsx"; //default
                    return correctArgs;
                case 3:
                    correctArgs[0] = args[0];
                    correctArgs[1] = args[1];
                    correctArgs[2] = args[2];
                    return correctArgs;
                default:
                    throw new Exception("Your arguments aren't valid");
            }
        }

        private static void CheckArgs (ref string[] args)
        {
            while (true)
            {
                if (File.Exists(args[0]) && (args[0].EndsWith(".docx") || args[0].EndsWith(".doc")))
                {
                    if (File.Exists(args[1]) && (args[1].EndsWith(".xlsx") || args[1].EndsWith(".xls")))
                    {
                        if (File.Exists(args[2]) && (args[2].EndsWith(".xlsx") || args[2].EndsWith(".xls")))
                        {
                            break;
                        }
                        else
                        {
                            Console.WriteLine("The path to Reasons is incorrect. Please, enter the correct path");
                            args[2] = Console.ReadLine();
                            continue;
                        }
                    }
                    else
                    {
                        Console.WriteLine("The path to Values is incorrect. Please, enter the correct path");
                        args[1] = Console.ReadLine();
                        continue;
                    }
                }
                else
                {
                    Console.WriteLine("The path to template is incorrect. Please, enter the correct path");
                    args[0] = Console.ReadLine();
                    continue;
                }
            }
        }
    }
}
