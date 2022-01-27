using ClosedXML.Excel;
using Newtonsoft.Json;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Services;
using TableExporter.Models;

namespace TableExporter.Controllers
{
    public class ExcelUploadController : Controller
    {
        public ActionResult ImportExcel()
        {
            return View();
        }

        [ActionName("Importexcel")]
        [HttpPost]
        public ActionResult Importexcel1()
        {

            if (Request.Files["FileUpload1"].ContentLength > 0)
            {
                string extension = System.IO.Path.GetExtension(Request.Files["FileUpload1"].FileName).ToLower();
                //string query = null;
                string connString = "";

                string[] validFileTypes = { ".xls", ".xlsx" };

                string path1 = string.Format("{0}/{1}", Server.MapPath("~/Content/Uploads"), Request.Files["FileUpload1"].FileName);
                if (!Directory.Exists(path1))
                {
                    Directory.CreateDirectory(Server.MapPath("~/Content/Uploads"));
                }
                if (validFileTypes.Contains(extension))
                {
                    if (System.IO.File.Exists(path1))
                    { System.IO.File.Delete(path1); }
                    Request.Files["FileUpload1"].SaveAs(path1);
                    //if (extension == ".csv")
                    //{
                    //    DataTable dt = Utility.ConvertCSVtoDataTable(path1);
                    //    ViewBag.Data = dt;
                    //}
                    //Connection String to Excel Workbook  
                    if (extension.Trim() == ".xls")
                    {
                        connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path1 + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
                        DataTable dt = Utility.ConvertXSLXtoDataTable(path1, connString);
                        ViewBag.Data = dt;
                    }
                    else if (extension.Trim() == ".xlsx")
                    {
                        connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path1 + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
                        DataTable dt = Utility.ConvertXSLXtoDataTable(path1, connString);
                        ViewBag.Data = dt;
                    }

                }
                else
                {
                    ViewBag.Error = "Пожалуйста, загрузите файлы в формате .xls или .xlsx";
                }

            }

            return View();
        }


        [HttpPost]
        public ActionResult ExportTable(string[] col, string[] row, string submitButton)
        {
            // создание dataTable
            DataTable dt = new DataTable();

            int t = 0;
            // формирование th
            for (int i = 0; i < col.Count(); i++)
            {
                dt.Columns.Add(col[i].ToString());
            }

            // формирование td
            for (int i = 0; i < row.Count()/col.Count(); i++)
            {
                DataRow dr = dt.NewRow();
                for (int j = 0; j < col.Count(); j++)
                {
                    dr[col[j].ToString()] = row[t].ToString();
                    t++;
                }
                dt.Rows.Add(dr);
            }

            switch (submitButton)
            {
                case "Выгрузить в DBF":
                    return (ExportToDBF(dt));
                case "Выгрузить в XLS":
                    return (ExportToXLS(dt));
                case "Выгрузить в XLSX":
                    return (ExportToXLSX(dt));
                case "Выгрузить в CSV":
                    return (ExportToCSV(dt));
                default:
                    return View("ImportExcel");
            }
        }

        [HttpPost]
        private ActionResult ExportToDBF(DataTable dt)
        {

            var resp = System.Web.HttpContext.Current.Response;
            Spire.DataExport.DBF.DBFExport DBFExport = new Spire.DataExport.DBF.DBFExport();
            DBFExport.DataSource = Spire.DataExport.Common.ExportSource.DataTable;
            DBFExport.DataTable = dt as DataTable;
            resp.AddHeader("Transfer-Encoding", "identity");
            DBFExport.SaveToHttpResponse("dbf_file_" + DateTime.Now.ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss") + ".dbf", resp);

            return View("ImportExcel");
        }

        [HttpPost]
        private ActionResult ExportToXLS(DataTable dt)
        {
            var resp = System.Web.HttpContext.Current.Response;
            Spire.DataExport.XLS.WorkSheet workSheet1 = new Spire.DataExport.XLS.WorkSheet();
            Spire.DataExport.XLS.CellExport cellExport1 = new Spire.DataExport.XLS.CellExport();
            workSheet1.DataSource = Spire.DataExport.Common.ExportSource.DataTable;
            workSheet1.DataTable = dt as DataTable;
            workSheet1.StartDataCol = ((System.Byte)(0));
            cellExport1.Sheets.Add(workSheet1);
            resp.AddHeader("Transfer-Encoding", "identity");
            cellExport1.SaveToHttpResponse("xls_file_" + DateTime.Now.ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss") + ".xls", resp);


            return View("ImportExcel");
        }

        [HttpPost]
        private ActionResult ExportToXLSX(DataTable dt)
        {
            var resp = System.Web.HttpContext.Current.Response;
            XLWorkbook wb = new XLWorkbook();
            wb.Worksheets.Add(dt, "Sheet1");
            Response.Clear();
            Response.Buffer = true;
            Response.Charset = "";
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment;filename=xlsx_file_" + DateTime.Now.ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss") + ".xlsx");
            MemoryStream MyMemoryStream = new MemoryStream();
            wb.SaveAs(MyMemoryStream);
            MyMemoryStream.WriteTo(Response.OutputStream);
            Response.Flush();
            Response.End();


            return View("ImportExcel");
        }

        [HttpPost]
        private ActionResult ExportToCSV(DataTable dt)
        {

            var resp = System.Web.HttpContext.Current.Response;
            Spire.DataExport.TXT.TXTExport txtExport1 = new Spire.DataExport.TXT.TXTExport();
            txtExport1.ExportType = Spire.DataExport.TXT.TextExportType.CSV;
            txtExport1.DataSource = Spire.DataExport.Common.ExportSource.DataTable;
            txtExport1.DataTable = dt as DataTable;
            Response.AddHeader("Transfer-Encoding", "identity");
            txtExport1.SaveToHttpResponse("csv_file_" + DateTime.Now.ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss") + ".csv", resp);

            return View("ImportExcel");
        }

    }
}