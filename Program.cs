using System;
using System.IO;

namespace InventorStepExport
{
    class Program
    {
        static void Main(string[] args)
        {
            string inputFolder = AskForFolder();
            string outputFolder = Path.Combine(inputFolder, "STEP_Output");

            Run(inputFolder, outputFolder);

            Console.WriteLine();
            Console.WriteLine("Press any key to exit.");
            Console.ReadKey();
        }

        static string AskForFolder()
        {
            while (true)
            {
                Console.Write("Enter the folder containing IPT files: ");

                // Trim spaces and quotes (e.g. from "Copy as path").
                string path = (Console.ReadLine() ?? "").Trim().Trim('"');

                if (Directory.Exists(path))
                    return path;

                Console.WriteLine("Folder not found. Please try again.");
            }
        }

        static void Run(string inputFolder, string outputFolder)
        {
            try
            {
                if (!Directory.Exists(inputFolder))
                {
                    Console.WriteLine("Input folder not found: " + inputFolder);
                    return;
                }

                Directory.CreateDirectory(outputFolder);

                var session = new InventorSession();
                session.Connect();

                Console.WriteLine("Connected to Autodesk Inventor.");
                Console.WriteLine("Output: " + outputFolder);

                var exporter = new StepExporter(
                    session.Application,
                    session.GetStepTranslator(),
                    outputFolder);

                ExportResult result = exporter.ExportFolder(inputFolder);

                Console.WriteLine();
                Console.WriteLine("========== EXPORT SUMMARY ==========");
                Console.WriteLine("IPT files found: " + result.FilesFound);
                Console.WriteLine("STEP files exported: " + result.Exported);
                Console.WriteLine("Failures: " + result.Failed);
                Console.WriteLine("Output folder: " + outputFolder);
            }
            catch (Exception ex)
            {
                Console.WriteLine("FATAL ERROR: " + ex.Message);
            }
        }
    }
}
