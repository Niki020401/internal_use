using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ACW = Autodesk.Connectivity.WebServices;
using Autodesk.DataManagement.Client.Framework.Currency;
using Autodesk.DataManagement.Client.Framework.Vault.Currency.Connections;
using Autodesk.DataManagement.Client.Framework.Vault.Currency.Entities;
using Autodesk.DataManagement.Client.Framework.Vault.Settings;

namespace InventorModelStateCleanup
{
    // ================================================================
    // VAULT DOWNLOADER
    // Finds the latest version of a file in Vault by name and
    // downloads it to a local folder.
    // ================================================================
    class VaultDownloader
    {
        readonly Connection _conn;
        ACW.PropDef _nameDef;

        public VaultDownloader(Connection conn)
        {
            _conn = conn;
        }

        // ============================================================
        // BLOCK V1: DOWNLOAD ONE FILE
        // Returns true if the file exists locally after the download.
        // ============================================================

        public bool Download(string fileName, string destFolder)
        {
            try
            {
                ACW.File vaultFile = FindLatestFile(fileName);

                if (vaultFile == null)
                {
                    Console.WriteLine("VAULT: not found: " + fileName);
                    return false;
                }

                Console.WriteLine("VAULT: found " + vaultFile.Name +
                                  " (version " + vaultFile.VerNum + ")");

                AcquireFilesSettings settings = new AcquireFilesSettings(_conn);
                settings.LocalPath = new FolderPathAbsolute(destFolder);
                settings.OptionsResolution.OverwriteOption =
                    AcquireFilesSettings.AcquireFileResolutionOptions.OverwriteOptions.ForceOverwriteAll;

                settings.AddFileToAcquire(
                    new FileIteration(_conn, vaultFile),
                    AcquireFilesSettings.AcquisitionOption.Download);

                _conn.FileManager.AcquireFiles(settings);

                string localPath = Path.Combine(destFolder, fileName);
                bool ok = System.IO.File.Exists(localPath);

                Console.WriteLine(ok
                    ? "VAULT: downloaded to " + localPath
                    : "VAULT: download finished but file not found at " + localPath);

                return ok;
            }
            catch (Exception ex)
            {
                Console.WriteLine("VAULT: FAILED " + fileName + " - " + ex.Message);
                return false;
            }
        }

        // ============================================================
        // BLOCK V2: SEARCH VAULT BY EXACT FILE NAME (latest version)
        // ============================================================

        ACW.File FindLatestFile(string fileName)
        {
            if (_nameDef == null)
            {
                ACW.PropDef[] defs = _conn.WebServiceManager.PropertyService
                    .GetPropertyDefinitionsByEntityClassId("FILE");

                _nameDef = defs.FirstOrDefault(p => p.SysName == "ClientFileName");

                if (_nameDef == null)
                    throw new Exception("File name property definition not found in Vault.");
            }

            ACW.SrchCond cond = new ACW.SrchCond();
            cond.PropDefId = _nameDef.Id;
            cond.PropTyp = ACW.PropertySearchType.SingleProperty;
            cond.SrchOper = 3; // 3 = "is exactly"
            cond.SrchRule = ACW.SearchRuleType.Must;
            cond.SrchTxt = fileName;

            List<ACW.File> matches = new List<ACW.File>();
            string bookmark = string.Empty;
            ACW.SrchStatus status = null;

            while (status == null || matches.Count < status.TotalHits)
            {
                ACW.File[] page = _conn.WebServiceManager.DocumentService
                    .FindFilesBySearchConditions(
                        new ACW.SrchCond[] { cond },
                        null,   // no sort
                        null,   // all folders
                        true,   // recurse
                        true,   // latest version only
                        ref bookmark,
                        out status);

                if (page == null) break;
                matches.AddRange(page);
            }

            ACW.File[] exact = matches
                .Where(f => string.Equals(f.Name, fileName, StringComparison.OrdinalIgnoreCase))
                .ToArray();

            if (exact.Length > 1)
                Console.WriteLine("VAULT: WARNING " + exact.Length +
                                  " files named " + fileName + ", using the first.");

            return exact.FirstOrDefault();
        }
    }
}
