using ClosedXML.Excel;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace WebReport.Controllers
{
    public class HomeController : Controller
    {
        public static string connect = "aspnet_WebReport_20200726103735ConnectionString";
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult TienDo()
        {
           
            return View();
        }
        public ActionResult ThongTin()
        {
         

            return View();
        }

        public ActionResult CapNhat()
        {
            return View();
        }
        [HttpPost]
        public ActionResult CapNhat(HttpPostedFileBase file)
        {
            DataTable dt = new DataTable();
            //Checking file content length and Extension must be .xlsx  
            if (file != null && file.ContentLength > 0 && System.IO.Path.GetExtension(file.FileName).ToLower() == ".xlsx")
            {
                string path = System.IO.Path.Combine(Server.MapPath("~"), System.IO.Path.GetFileName(file.FileName));
                //Saving the file  
                file.SaveAs(path);
                //Started reading the Excel file.  
                using (XLWorkbook workbook = new XLWorkbook(path))
                {
                    IXLWorksheet worksheet = workbook.Worksheet(1);
                    bool FirstRow = true;
                    //Range for reading the cells based on the last cell used.  
                    string readRange = "1:1";
                    foreach (IXLRow row in worksheet.RowsUsed())
                    {
                        //If Reading the First Row (used) then add them as column name  
                        if (FirstRow)
                        {
                            //Checking the Last cellused for column generation in datatable  
                            readRange = string.Format("{0}:{1}", 1, row.LastCellUsed().Address.ColumnNumber);
                            foreach (IXLCell cell in row.Cells(readRange))
                            {
                                dt.Columns.Add(cell.Value.ToString());
                            }
                            FirstRow = false;
                        }
                        else
                        {
                            //Adding a Row in datatable  
                            dt.Rows.Add();
                            int cellIndex = 0;
                            //Updating the values of datatable  
                            foreach (IXLCell cell in row.Cells(readRange))
                            {
                                dt.Rows[dt.Rows.Count - 1][cellIndex] = cell.Value.ToString();
                                cellIndex++;
                            }
                        }
                    }
                    //If no data in Excel file  
                    if (FirstRow)
                    {
                        ViewBag.Message = "Empty Excel File!";
                    }
                }
            }
            else
            {
                //If file extension of the uploaded file is different then .xlsx  
                ViewBag.Message = "Please select file with .xlsx extension!";
            }


            return View(dt);
        }
    }
}