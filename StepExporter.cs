using System;
using System.IO;
using Inventor;

namespace InventorStepExport
{
    class ExportResult
    {
        public int FilesFound;
        public int Exported;
        public int Failed;
    }

    /// <summary>
    /// Opens each IPT in a folder and exports every Model State to STEP.
    /// </summary>
    class StepExporter
    {
        private readonly Inventor.Application _inventor;
        private readonly TranslatorAddIn _translator;
        private readonly string _outputFolder;

        public StepExporter(
            Inventor.Application inventor,
            TranslatorAddIn translator,
            string outputFolder)
        {
            _inventor = inventor;
            _translator = translator;
            _outputFolder = outputFolder;
        }

        public ExportResult ExportFolder(string inputFolder)
        {
            // Only IPT files directly inside the input folder.
            string[] iptFiles = Directory.GetFiles(
                inputFolder, "*.ipt", SearchOption.TopDirectoryOnly);

            var result = new ExportResult { FilesFound = iptFiles.Length };

            foreach (string iptPath in iptFiles)
                ExportPart(iptPath, result);

            return result;
        }

        private void ExportPart(string iptPath, ExportResult result)
        {
            PartDocument partDoc = null;
            ModelState originalState = null;

            try
            {
                Console.WriteLine();
                Console.WriteLine("Opening: " + Path.GetFileName(iptPath));

                partDoc = (PartDocument)_inventor.Documents.Open(iptPath, false);

                ModelStates states = partDoc.ComponentDefinition.ModelStates;

                // Store state names before activating states.
                string[] stateNames = new string[states.Count];

                for (int i = 1; i <= states.Count; i++)
                {
                    stateNames[i - 1] = states[i].Name;

                    if (states[i].Active)
                        originalState = states[i];
                }

                foreach (string stateName in stateNames)
                    ExportState(partDoc, states, iptPath, stateName, result);

                // Restore the state that was active originally.
                if (originalState != null)
                {
                    try
                    {
                        originalState.Activate();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(
                            "Could not restore original state: " + ex.Message);
                    }
                }

                partDoc.Close(false);
                partDoc = null;
            }
            catch (Exception ex)
            {
                Console.WriteLine("IPT ERROR: " + ex.Message);
                result.Failed++;

                if (partDoc != null)
                {
                    try { partDoc.Close(false); } catch { }
                }
            }
        }

        private void ExportState(
            PartDocument partDoc,
            ModelStates states,
            string iptPath,
            string stateName,
            ExportResult result)
        {
            try
            {
                ModelState state = states[stateName];

                Console.WriteLine("Exporting state: " + stateName);

                // Activate the requested Model State.
                state.Activate();
                partDoc.Update2(true);

                string outputName = BuildOutputName(iptPath, stateName);
                string stepPath = Path.Combine(_outputFolder, outputName);

                // Avoid overwrite confirmation dialogs.
                if (File.Exists(stepPath))
                    File.Delete(stepPath);

                SaveAsStep(partDoc, stepPath);

                if (File.Exists(stepPath) && new FileInfo(stepPath).Length > 0)
                {
                    Console.WriteLine("SUCCESS: " + outputName);
                    result.Exported++;
                }
                else
                {
                    Console.WriteLine("FAILED: STEP file not created.");
                    result.Failed++;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("FAILED state " + stateName + ": " + ex.Message);
                result.Failed++;
            }
        }

        private void SaveAsStep(PartDocument partDoc, string stepPath)
        {
            TranslationContext context =
                _inventor.TransientObjects.CreateTranslationContext();
            context.Type = IOMechanismEnum.kFileBrowseIOMechanism;

            NameValueMap options =
                _inventor.TransientObjects.CreateNameValueMap();

            _translator.GetSaveCopyAsOptions(partDoc, context, options);

            DataMedium data = _inventor.TransientObjects.CreateDataMedium();
            data.FileName = stepPath;

            _translator.SaveCopyAs(partDoc, context, options, data);
        }

        private static string BuildOutputName(string iptPath, string stateName)
        {
            // Primary exists in every IPT, so include the IPT name
            // to avoid collisions.
            bool isPrimary =
                stateName.Equals("[Primary]", StringComparison.OrdinalIgnoreCase) ||
                stateName.Equals("Primary", StringComparison.OrdinalIgnoreCase);

            if (isPrimary)
                return Path.GetFileNameWithoutExtension(iptPath) + "_Primary.step";

            return MakeSafeFileName(stateName) + ".step";
        }

        private static string MakeSafeFileName(string name)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');

            return name.Trim();
        }
    }
}
