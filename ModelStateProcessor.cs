using System;
using System.IO;
using Inventor;

namespace InventorModelStateCleanup
{
    // ================================================================
    // PART 3 OF 3: MODEL STATE PROCESSOR
    // For one requested state: find IPT, copy, clean states, save.
    // ================================================================
    class ModelStateProcessor
    {
        readonly Inventor.Application _inventor;
        readonly string _inputFolder;
        readonly string _outputFolder;
        readonly bool _overwriteOutput;

        public ModelStateProcessor(Inventor.Application inventor,
                                   string inputFolder,
                                   string outputFolder,
                                   bool overwriteOutput)
        {
            _inventor = inventor;
            _inputFolder = inputFolder;
            _outputFolder = outputFolder;
            _overwriteOutput = overwriteOutput;
        }

        // ============================================================
        // BLOCK 9: PROCESS ONE REQUESTED STATE
        // ============================================================

        public void Process(string stateName)
        {
            Console.WriteLine();
            Console.WriteLine("Requested state: " + stateName);

            // --------------------------------------------------------
            // BLOCK 10: FIND MATCHING IPT FILE
            // Example: 3215932P01 -> 3215932.ipt
            // --------------------------------------------------------

            int pIndex = stateName.LastIndexOf('P');

            if (pIndex <= 0)
            {
                Console.WriteLine("SKIPPED: No 'P' found in state name: " + stateName);
                return;
            }

            string baseFileName = stateName.Substring(0, pIndex);
            string sourcePath = Path.Combine(_inputFolder, baseFileName + ".ipt");

            if (!File.Exists(sourcePath))
            {
                Console.WriteLine("SKIPPED: IPT file not found: " + sourcePath);
                return;
            }

            Console.WriteLine("Matching IPT file: " + sourcePath);

            string outputPath = Path.Combine(_outputFolder, Path.GetFileName(sourcePath));

            if (File.Exists(outputPath) && !_overwriteOutput)
            {
                Console.WriteLine("SKIPPED: Output already exists: " + outputPath);
                return;
            }

            bool success = false;
            Document doc = null;

            try
            {
                // ----------------------------------------------------
                // BLOCK 11: CREATE OUTPUT COPY
                // ----------------------------------------------------

                if (File.Exists(outputPath))
                {
                    File.SetAttributes(outputPath, FileAttributes.Normal);
                    File.Delete(outputPath);
                }

                File.Copy(sourcePath, outputPath, false);
                File.SetAttributes(outputPath, FileAttributes.Normal); // remove read-only

                // ----------------------------------------------------
                // BLOCK 12: OPEN THE OUTPUT COPY (invisible)
                // ----------------------------------------------------

                doc = _inventor.Documents.Open(outputPath, false);

                PartDocument part = doc as PartDocument;
                if (part == null)
                    throw new Exception("Output is not an Inventor part.");

                ModelStates states = part.ComponentDefinition.ModelStates;

                // ----------------------------------------------------
                // BLOCK 13: LIST ALL STATES + FIND PRIMARY & REQUESTED
                // ----------------------------------------------------

                Console.WriteLine("Model states in file (" + states.Count + "):");

                int primaryIndex = -1;
                int requestedIndex = -1;

                for (int i = 1; i <= states.Count; i++)
                {
                    string name = states[i].Name;
                    Console.WriteLine("  [" + i + "] '" + name + "'");

                    if (primaryIndex < 0 && IsPrimaryName(name))
                        primaryIndex = i;

                    if (requestedIndex < 0 && NamesMatch(name, stateName))
                        requestedIndex = i;
                }

                // Primary is always the first model state in Inventor.
                if (primaryIndex < 0 && states.Count >= 1)
                    primaryIndex = 1;

                if (primaryIndex < 0)
                    throw new Exception("Primary Model State not found.");

                if (requestedIndex < 0)
                    throw new Exception("Requested Model State '" + stateName + "' not found in file.");

                string primaryName = states[primaryIndex].Name;
                string requestedName = states[requestedIndex].Name;

                // ----------------------------------------------------
                // BLOCK 14: ACTIVATE PRIMARY (active state can't be deleted)
                // ----------------------------------------------------

                states[primaryIndex].Activate();
                Console.WriteLine("Activated: " + primaryName);

                // ----------------------------------------------------
                // BLOCK 15: DELETE EVERYTHING EXCEPT PRIMARY + REQUESTED
                // Loop backwards because indexes shift on delete.
                // ----------------------------------------------------

                for (int i = states.Count; i >= 1; i--)
                {
                    string currentName = states[i].Name;

                    if (string.Equals(currentName, primaryName, StringComparison.Ordinal) ||
                        string.Equals(currentName, requestedName, StringComparison.Ordinal))
                    {
                        Console.WriteLine("Keeping: " + currentName);
                        continue;
                    }

                    try
                    {
                        Console.WriteLine("Deleting: " + currentName);
                        states[i].Delete();
                    }
                    catch (Exception delEx)
                    {
                        throw new Exception(
                            "Could not delete state '" + currentName + "': " + delEx.Message);
                    }
                }

                // ----------------------------------------------------
                // BLOCK 16: SAVE CLEANED IPT
                // ----------------------------------------------------

                part.Update2(true);
                doc.Save();
                success = true;

                Console.WriteLine("Saved: " + outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("FAILED: " + stateName + " - " + ex.Message);
            }
            finally
            {
                if (doc != null)
                {
                    try { doc.Close(true); } catch { }
                }

                // Remove half-processed copy so no bad files are left behind
                if (!success && File.Exists(outputPath))
                {
                    try { File.Delete(outputPath); } catch { }
                }
            }
        }

        // ============================================================
        // BLOCK 17: NAME HELPERS
        // ============================================================

        // Matches "[Primary]" or "Primary"
        static bool IsPrimaryName(string name)
        {
            if (name == null) return false;
            string n = name.Trim().TrimStart('[').TrimEnd(']').Trim();
            return string.Equals(n, "Primary", StringComparison.OrdinalIgnoreCase);
        }

        // Trim + case-insensitive compare
        static bool NamesMatch(string a, string b)
        {
            if (a == null || b == null) return false;
            return string.Equals(a.Trim(), b.Trim(), StringComparison.OrdinalIgnoreCase);
        }
    }
}
