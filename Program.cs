using System;
using System.IO;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Inventor;

namespace InventorModelStateCleanup
{
    // ================================================================
    // PART 1 OF 3: PROGRAM (entry point + workflow)
    // Asks for Excel, reads names, connects to Inventor, loops states.
    // ================================================================
    class Program
    {
        // ============================================================
        // BLOCK 0: CONFIGURATION
        // ============================================================

        static readonly string InputFolder  = @"C:\Bharat_H_S\Nikhil\Task";
        static readonly string OutputFolder = @"C:\Bharat_H_S\Nikhil\Task\Output";

        const bool OverwriteOutput = true; // true = replace old output copies

        static void Main(string[] args)
        {
            try
            {
                // ====================================================
                // BLOCK 1: SELECT EXCEL FILE
                // ====================================================

                Console.WriteLine("MODEL STATE CLEANUP");
                Console.WriteLine("-------------------");
                Console.Write("Enter the full path of the Excel file: ");

                string excelPath = (Console.ReadLine() ?? "").Trim().Trim('"');

                if (string.IsNullOrWhiteSpace(excelPath) || !File.Exists(excelPath))
                {
                    Console.WriteLine("Excel file not found.");
                    return;
                }

                if (!Directory.Exists(InputFolder))
                {
                    Console.WriteLine("Input folder not found.");
                    return;
                }

                Directory.CreateDirectory(OutputFolder);

                // ====================================================
                // BLOCK 2: READ REQUIRED MODEL STATE NAMES (Part 2)
                // ====================================================

                List<string> requiredStates = ExcelReader.ReadModelStateNames(excelPath);

                if (requiredStates.Count == 0)
                {
                    Console.WriteLine("No Model State names found in Excel.");
                    return;
                }

                // ====================================================
                // BLOCK 3: CONNECT TO INVENTOR (must already be running)
                // ====================================================

                Inventor.Application inventor = (Inventor.Application)
                    Marshal.GetActiveObject("Inventor.Application");

                // ====================================================
                // BLOCK 4: PROCESS EACH STATE (Part 3)
                // ====================================================

                ModelStateProcessor processor =
                    new ModelStateProcessor(inventor, InputFolder, OutputFolder, OverwriteOutput);

                foreach (string stateName in requiredStates)
                {
                    processor.Process(stateName);
                }

                // ====================================================
                // BLOCK 5: FINISH
                // ====================================================

                Console.WriteLine();
                Console.WriteLine("Processing completed.");
                Console.WriteLine("Output folder: " + OutputFolder);
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: " + ex.Message);
            }

            Console.WriteLine("Press Enter to exit.");
            Console.ReadLine();
        }
    }
}
