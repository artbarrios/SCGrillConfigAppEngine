using SCGrillConfigAppEngine.Models;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace SCGrillConfigAppEngine
{
    class ReportManager
    {
        // spin up a copy of Word for use during this session
        public static Application app = new Application();
        // specify the directory where all reports are to be located
        private static string fileSaveDirectory = AppCommon.GetFileSaveDirectory();
        // specify the base address for the WebAPI uri
        private static string webApiAddress = AppCommon.GetRemoteWebApiUrl();

        public static void GenerateReport(Report report)
        {
            // generates the specified report in the specified format
            // gives the file the specified filename and stores it in the specified directory

            // check for valid input
            if (report.Name.Length == 0)
            {
                throw new Exception("GenerateReport: No report.Name specified.");
            }
            if (report.Filename.Length == 0)
            {
                throw new Exception("GenerateReport: No report.Filename specified.");
            }
            if (report.Url.Length == 0)
            {
                throw new Exception("GenerateReport: No report.Url specified.");
            }

            // generate the specified report
            AppCommon.Log("Generating report " + report.Name + ".", EventLogEntryType.Information);
            switch (report.Name.ToUpper())
            {
                case "SAMPLEREPORT":
                    Reports.SampleReport.Generate(report, fileSaveDirectory, app);
                    break; // SAMPLEREPORT
                case "GRILLSIZESINDEXPRINTERFRIENDLY":
                    Reports.GrillSizesIndexPrinterFriendlyReport.Generate(report, fileSaveDirectory, app);
                    break;
                case "SIDEBURNERTYPESINDEXPRINTERFRIENDLY":
                    Reports.SideBurnerTypesIndexPrinterFriendlyReport.Generate(report, fileSaveDirectory, app);
                    break;
                case "BUILDTASKSINDEXPRINTERFRIENDLY":
                    Reports.BuildTasksIndexPrinterFriendlyReport.Generate(report, fileSaveDirectory, app);
                    break;
                case "COLORSINDEXPRINTERFRIENDLY":
                    Reports.ColorsIndexPrinterFriendlyReport.Generate(report, fileSaveDirectory, app);
                    break;
                case "MATERIALSINDEXPRINTERFRIENDLY":
                    Reports.MaterialsIndexPrinterFriendlyReport.Generate(report, fileSaveDirectory, app);
                    break;
                case "FUELSINDEXPRINTERFRIENDLY":
                    Reports.FuelsIndexPrinterFriendlyReport.Generate(report, fileSaveDirectory, app);
                    break;
                case "GRILLTYPESINDEXPRINTERFRIENDLY":
                    Reports.GrillTypesIndexPrinterFriendlyReport.Generate(report, fileSaveDirectory, app);
                    break;
                case "GRILLCONFGURATIONSINDEXPRINTERFRIENDLY":
                    Reports.GrillConfgurationsIndexPrinterFriendlyReport.Generate(report, fileSaveDirectory, app);
                    break;

            }

            // purge old files older than the specified number of hours if purge enabled
            int countOfPurgedFiles = 0;
            if (AppCommon.IsPurgeOldFilesEnabled())
            {
                countOfPurgedFiles = AppCommon.PurgeOldFiles(fileSaveDirectory, AppCommon.GetPurgeAgeHours());
                if (countOfPurgedFiles > 0)
                {
                    AppCommon.Log("Purged " + countOfPurgedFiles.ToString() + " files from " + fileSaveDirectory + " .", EventLogEntryType.Information);
                }
            }

        } // GenerateReport

    }
}

