using SCGrillConfig.Models;
using SCGrillConfigAppEngine.Models;
using SCGrillConfigAppEngine.Web_Data;
using Microsoft.Office.Interop.Word;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Http;

namespace SCGrillConfigAppEngine
{
    public class ReportsController : ApiController
    {

        // GET /api/reports/SampleReport
        [Route("api/reports/SampleReport")]
        [HttpGet]
        public IHttpActionResult SampleReport()
        {
            try
            {
                // create report object, Url is the public location where it can be viewed with a browser
                Report report = new Report();
                report.Name = "SampleReport";
                report.Filename = "SampleReport";
                report.SaveFormat = WdSaveFormat.wdFormatPDF;
                report.Extension = AppCommon.GetExtensionFromWdSaveFormat(report.SaveFormat);
                report.Url = AppCommon.BuildUrl(AppCommon.GetAppEngineUrl(), report.Filename + "." + AppCommon.GetExtensionFromWdSaveFormat(report.SaveFormat), AppCommon.GetAppEnginePort());
                // generate the report
                ReportManager.GenerateReport(report);
                // return the report properties
                return Ok(report);
            }
            catch (Exception e)
            {
                string message = AppCommon.AppendInnerExceptionMessages("ReportsController.SampleReport = " + e.Message, e);
                AppCommon.Log(message, EventLogEntryType.Error);
                throw new Exception(message);
            }
        } // SampleReport()

        // GET /api/reports/GrillSizesIndexPrinterFriendly
        [Route("api/reports/GrillSizesIndexPrinterFriendly")]
        [HttpGet]
        public IHttpActionResult GrillSizesIndexPrinterFriendly()
        {
            try
            {
                // create report object, Url is the public location where it can be viewed with a browser
                Report report = new Report();
                report.Name = "GrillSizesIndexPrinterFriendly";
                report.Filename = "GrillSizesIndexPrinterFriendly";
                report.SaveFormat = WdSaveFormat.wdFormatPDF;
                report.Extension = AppCommon.GetExtensionFromWdSaveFormat(report.SaveFormat);
                report.Url = AppCommon.BuildUrl(AppCommon.GetAppEngineUrl(), report.Filename + "." + AppCommon.GetExtensionFromWdSaveFormat(report.SaveFormat), AppCommon.GetAppEnginePort());
                // generate the report
                ReportManager.GenerateReport(report);
                // return the report properties
                return Ok(report);
            }
            catch (Exception e)
            {
                string message = AppCommon.AppendInnerExceptionMessages("ReportsController.GrillSizesIndexPrinterFriendly = " + e.Message, e);
                AppCommon.Log(message, EventLogEntryType.Error);
                throw new Exception(message);
            }
        } // GrillSizesIndexPrinterFriendly()

        // GET /api/reports/SideBurnerTypesIndexPrinterFriendly
        [Route("api/reports/SideBurnerTypesIndexPrinterFriendly")]
        [HttpGet]
        public IHttpActionResult SideBurnerTypesIndexPrinterFriendly()
        {
            try
            {
                // create report object, Url is the public location where it can be viewed with a browser
                Report report = new Report();
                report.Name = "SideBurnerTypesIndexPrinterFriendly";
                report.Filename = "SideBurnerTypesIndexPrinterFriendly";
                report.SaveFormat = WdSaveFormat.wdFormatPDF;
                report.Extension = AppCommon.GetExtensionFromWdSaveFormat(report.SaveFormat);
                report.Url = AppCommon.BuildUrl(AppCommon.GetAppEngineUrl(), report.Filename + "." + AppCommon.GetExtensionFromWdSaveFormat(report.SaveFormat), AppCommon.GetAppEnginePort());
                // generate the report
                ReportManager.GenerateReport(report);
                // return the report properties
                return Ok(report);
            }
            catch (Exception e)
            {
                string message = AppCommon.AppendInnerExceptionMessages("ReportsController.SideBurnerTypesIndexPrinterFriendly = " + e.Message, e);
                AppCommon.Log(message, EventLogEntryType.Error);
                throw new Exception(message);
            }
        } // SideBurnerTypesIndexPrinterFriendly()

        // GET /api/reports/BuildTasksIndexPrinterFriendly
        [Route("api/reports/BuildTasksIndexPrinterFriendly")]
        [HttpGet]
        public IHttpActionResult BuildTasksIndexPrinterFriendly()
        {
            try
            {
                // create report object, Url is the public location where it can be viewed with a browser
                Report report = new Report();
                report.Name = "BuildTasksIndexPrinterFriendly";
                report.Filename = "BuildTasksIndexPrinterFriendly";
                report.SaveFormat = WdSaveFormat.wdFormatPDF;
                report.Extension = AppCommon.GetExtensionFromWdSaveFormat(report.SaveFormat);
                report.Url = AppCommon.BuildUrl(AppCommon.GetAppEngineUrl(), report.Filename + "." + AppCommon.GetExtensionFromWdSaveFormat(report.SaveFormat), AppCommon.GetAppEnginePort());
                // generate the report
                ReportManager.GenerateReport(report);
                // return the report properties
                return Ok(report);
            }
            catch (Exception e)
            {
                string message = AppCommon.AppendInnerExceptionMessages("ReportsController.BuildTasksIndexPrinterFriendly = " + e.Message, e);
                AppCommon.Log(message, EventLogEntryType.Error);
                throw new Exception(message);
            }
        } // BuildTasksIndexPrinterFriendly()

        // GET /api/reports/ColorsIndexPrinterFriendly
        [Route("api/reports/ColorsIndexPrinterFriendly")]
        [HttpGet]
        public IHttpActionResult ColorsIndexPrinterFriendly()
        {
            try
            {
                // create report object, Url is the public location where it can be viewed with a browser
                Report report = new Report();
                report.Name = "ColorsIndexPrinterFriendly";
                report.Filename = "ColorsIndexPrinterFriendly";
                report.SaveFormat = WdSaveFormat.wdFormatPDF;
                report.Extension = AppCommon.GetExtensionFromWdSaveFormat(report.SaveFormat);
                report.Url = AppCommon.BuildUrl(AppCommon.GetAppEngineUrl(), report.Filename + "." + AppCommon.GetExtensionFromWdSaveFormat(report.SaveFormat), AppCommon.GetAppEnginePort());
                // generate the report
                ReportManager.GenerateReport(report);
                // return the report properties
                return Ok(report);
            }
            catch (Exception e)
            {
                string message = AppCommon.AppendInnerExceptionMessages("ReportsController.ColorsIndexPrinterFriendly = " + e.Message, e);
                AppCommon.Log(message, EventLogEntryType.Error);
                throw new Exception(message);
            }
        } // ColorsIndexPrinterFriendly()

        // GET /api/reports/MaterialsIndexPrinterFriendly
        [Route("api/reports/MaterialsIndexPrinterFriendly")]
        [HttpGet]
        public IHttpActionResult MaterialsIndexPrinterFriendly()
        {
            try
            {
                // create report object, Url is the public location where it can be viewed with a browser
                Report report = new Report();
                report.Name = "MaterialsIndexPrinterFriendly";
                report.Filename = "MaterialsIndexPrinterFriendly";
                report.SaveFormat = WdSaveFormat.wdFormatPDF;
                report.Extension = AppCommon.GetExtensionFromWdSaveFormat(report.SaveFormat);
                report.Url = AppCommon.BuildUrl(AppCommon.GetAppEngineUrl(), report.Filename + "." + AppCommon.GetExtensionFromWdSaveFormat(report.SaveFormat), AppCommon.GetAppEnginePort());
                // generate the report
                ReportManager.GenerateReport(report);
                // return the report properties
                return Ok(report);
            }
            catch (Exception e)
            {
                string message = AppCommon.AppendInnerExceptionMessages("ReportsController.MaterialsIndexPrinterFriendly = " + e.Message, e);
                AppCommon.Log(message, EventLogEntryType.Error);
                throw new Exception(message);
            }
        } // MaterialsIndexPrinterFriendly()

        // GET /api/reports/FuelsIndexPrinterFriendly
        [Route("api/reports/FuelsIndexPrinterFriendly")]
        [HttpGet]
        public IHttpActionResult FuelsIndexPrinterFriendly()
        {
            try
            {
                // create report object, Url is the public location where it can be viewed with a browser
                Report report = new Report();
                report.Name = "FuelsIndexPrinterFriendly";
                report.Filename = "FuelsIndexPrinterFriendly";
                report.SaveFormat = WdSaveFormat.wdFormatPDF;
                report.Extension = AppCommon.GetExtensionFromWdSaveFormat(report.SaveFormat);
                report.Url = AppCommon.BuildUrl(AppCommon.GetAppEngineUrl(), report.Filename + "." + AppCommon.GetExtensionFromWdSaveFormat(report.SaveFormat), AppCommon.GetAppEnginePort());
                // generate the report
                ReportManager.GenerateReport(report);
                // return the report properties
                return Ok(report);
            }
            catch (Exception e)
            {
                string message = AppCommon.AppendInnerExceptionMessages("ReportsController.FuelsIndexPrinterFriendly = " + e.Message, e);
                AppCommon.Log(message, EventLogEntryType.Error);
                throw new Exception(message);
            }
        } // FuelsIndexPrinterFriendly()

        // GET /api/reports/GrillTypesIndexPrinterFriendly
        [Route("api/reports/GrillTypesIndexPrinterFriendly")]
        [HttpGet]
        public IHttpActionResult GrillTypesIndexPrinterFriendly()
        {
            try
            {
                // create report object, Url is the public location where it can be viewed with a browser
                Report report = new Report();
                report.Name = "GrillTypesIndexPrinterFriendly";
                report.Filename = "GrillTypesIndexPrinterFriendly";
                report.SaveFormat = WdSaveFormat.wdFormatPDF;
                report.Extension = AppCommon.GetExtensionFromWdSaveFormat(report.SaveFormat);
                report.Url = AppCommon.BuildUrl(AppCommon.GetAppEngineUrl(), report.Filename + "." + AppCommon.GetExtensionFromWdSaveFormat(report.SaveFormat), AppCommon.GetAppEnginePort());
                // generate the report
                ReportManager.GenerateReport(report);
                // return the report properties
                return Ok(report);
            }
            catch (Exception e)
            {
                string message = AppCommon.AppendInnerExceptionMessages("ReportsController.GrillTypesIndexPrinterFriendly = " + e.Message, e);
                AppCommon.Log(message, EventLogEntryType.Error);
                throw new Exception(message);
            }
        } // GrillTypesIndexPrinterFriendly()

        // GET /api/reports/GrillConfgurationsIndexPrinterFriendly
        [Route("api/reports/GrillConfgurationsIndexPrinterFriendly")]
        [HttpGet]
        public IHttpActionResult GrillConfgurationsIndexPrinterFriendly()
        {
            try
            {
                // create report object, Url is the public location where it can be viewed with a browser
                Report report = new Report();
                report.Name = "GrillConfgurationsIndexPrinterFriendly";
                report.Filename = "GrillConfgurationsIndexPrinterFriendly";
                report.SaveFormat = WdSaveFormat.wdFormatPDF;
                report.Extension = AppCommon.GetExtensionFromWdSaveFormat(report.SaveFormat);
                report.Url = AppCommon.BuildUrl(AppCommon.GetAppEngineUrl(), report.Filename + "." + AppCommon.GetExtensionFromWdSaveFormat(report.SaveFormat), AppCommon.GetAppEnginePort());
                // generate the report
                ReportManager.GenerateReport(report);
                // return the report properties
                return Ok(report);
            }
            catch (Exception e)
            {
                string message = AppCommon.AppendInnerExceptionMessages("ReportsController.GrillConfgurationsIndexPrinterFriendly = " + e.Message, e);
                AppCommon.Log(message, EventLogEntryType.Error);
                throw new Exception(message);
            }
        } // GrillConfgurationsIndexPrinterFriendly()

    }
}

