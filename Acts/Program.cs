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
        static void Main(string[] args)
        {
            Console.WriteLine("Welcome to Act Genereator!!!");
            Console.WriteLine();

            string[] correctArgs = new string[2];
            if (args.Length > 0)
            {
                correctArgs = SetArgs(args);
            }
            else
            {
                Console.WriteLine("Please, enter the path to Template:");
                correctArgs[0] = Console.ReadLine();
                Console.WriteLine("Thanks. And now, please, enter the path to Excel:");
                correctArgs[1] = Console.ReadLine();
            }

            correctArgs = CheckArgs(correctArgs);
            new Docs(correctArgs[0], correctArgs[1]).Execute();
        }

        
        private static string[] SetArgs (string[] args)
        {
            string[] correctArgs = new string[2];

            if (args.Length >= 2)
            {
                correctArgs[0] = args[0];
                correctArgs[1] = args[1];
            }
            else
            {
                correctArgs[0] = args[0];
                correctArgs[1] = "C:\\temp\\Template.docx"; //default
            } 

            return correctArgs;
        }

        private static string[] CheckArgs (string[] args)
        {
            while (true)
            {
                if (File.Exists(args[0]) && (args[0].EndsWith(".docx") || args[0].EndsWith(".doc")))
                {
                    if (File.Exists(args[1]) &&( args[1].EndsWith(".xlsx") || args[1].EndsWith(".xls")))
                    {
                        return args;
                    }
                    else
                    {
                        Console.WriteLine("The path to Excel is incorrect. Please, enter the correct path");
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
