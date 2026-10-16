using System;
using System.IO;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Inventor;
using VaultFileExtraction;

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

        // Vault settings
        const bool   UseVault      = true;                // false = use IPTs already in InputFolder
        const string VaultServer   = "YOUR_VAULT_SERVER"; // <-- change
        const string VaultName     = "YOUR_VAULT_NAME";   // <-- change
        const bool   ForceRedownload = false;             // true = always fetch latest from Vault

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
                // BLOCK 2B: DOWNLOAD IPT FILES FROM VAULT
                // Only files not already in InputFolder (unless
                // ForceRedownload = true).
                // ====================================================

                if (UseVault)
                {
                    DownloadFromVault(requiredStates);
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

        // ============================================================
        // BLOCK 6B: VAULT DOWNLOAD STEP
        // ============================================================

        static void DownloadFromVault(List<string> requiredStates)
        {
            // One download per part (many states can share one IPT).
            List<string> baseNames = new List<string>();

            foreach (string stateName in requiredStates)
            {
                string baseName = ModelStateProcessor.GetBaseFileName(stateName);

                if (baseName != null && !baseNames.Contains(baseName))
                    baseNames.Add(baseName);
            }

            Console.WriteLine();
            Console.WriteLine("Connecting to Vault...");

            Vault vault = new Vault();

            try
            {
                var connection = vault.Connect(VaultServer, VaultName);
                VaultDownloader downloader = new VaultDownloader(connection);

                foreach (string baseName in baseNames)
                {
                    string fileName = baseName + ".ipt";
                    string localPath = Path.Combine(InputFolder, fileName);

                    if (File.Exists(localPath) && !ForceRedownload)
                    {
                        Console.WriteLine("VAULT: already local, skipping " + fileName);
                        continue;
                    }

                    downloader.Download(fileName, InputFolder);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("VAULT ERROR: " + ex.Message);
                Console.WriteLine("Continuing with IPT files already in the input folder.");
            }
            finally
            {
                try { vault.Disconnect(); } catch { }
            }
        }
    }
}
