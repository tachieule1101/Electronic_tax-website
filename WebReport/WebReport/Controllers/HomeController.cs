using ClosedXML.Excel;
using System;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Web.Mvc;
using WebReport.Models;

namespace WebReport.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }
        public static DataTable TienDoCBQL()
        {
            DBConnection DB = new DBConnection();
            string sql;
            sql = "SELECT DISTINCT TT,MA,CHITIEU,THANG,QUY,KHQUY,TLQUY,NAM,KHNAM,TLNAM FROM TienDoCBQL ORDER BY TT ASC";
            SqlConnection con = DB.getConnection();
            SqlDataAdapter cmd = new SqlDataAdapter(sql, con);
            DataTable dt = new DataTable();
            try
            {
                con.Open();
                cmd.Fill(dt);
                cmd.Dispose();
                con.Close();
            }
            catch (Exception ex)
            {}
            return dt;
        }
        public static DataTable TienDoLT()
        {
            DBConnection DB = new DBConnection();
            string sql;
            sql = "SELECT DISTINCT TT,MA,CHITIEU,THANG,QUY,KHQUY,TLQUY,NAM,KHNAM,TLNAM FROM TienDoLT ORDER BY TT ASC";
            SqlConnection con = DB.getConnection();
            SqlDataAdapter cmd = new SqlDataAdapter(sql, con);
            DataTable dt = new DataTable();
            try
            {
                con.Open();
                cmd.Fill(dt);
                cmd.Dispose();
                con.Close();
            }
            catch (Exception ex)
            { }
            return dt;
        }
        public static DataTable TienDoMLNS()
        {
            DBConnection DB = new DBConnection();
            string sql;
            sql = "SELECT DISTINCT TT,MA,CHITIEU,THANG,QUY,KHQUY,TLQUY,NAM,KHNAM,TLNAM FROM TienDoMLNS ORDER BY TT ASC";
            SqlConnection con = DB.getConnection();
            SqlDataAdapter cmd = new SqlDataAdapter(sql, con);
            DataTable dt = new DataTable();
            try
            {
                con.Open();
                cmd.Fill(dt);
                cmd.Dispose();
                con.Close();
            }
            catch (Exception ex)
            { }
            return dt;
        }
        public static DataTable TienDoCBQL1()
        {
            DBConnection DB = new DBConnection();
            string sql;
            sql = "SELECT DISTINCT TT,MA,CHITIEU,THANG,QUY,KHQUY,TLQUY,NAM,KHNAM,TLNAM FROM TienDoCBQL1 ORDER BY TT ASC";
            SqlConnection con = DB.getConnection();
            SqlDataAdapter cmd = new SqlDataAdapter(sql, con);
            DataTable dt = new DataTable();
            try
            {
                con.Open();
                cmd.Fill(dt);
                cmd.Dispose();
                con.Close();
            }
            catch (Exception ex)
            { }
            return dt;
        }
        public static DataTable TienDoMLNS1()
        {
            DBConnection DB = new DBConnection();
            string sql;
            sql = "SELECT DISTINCT TT,MA,CHITIEU,THANG,QUY,KHQUY,TLQUY,NAM,KHNAM,TLNAM FROM TienDoMLNS1 ORDER BY TT ASC";

            SqlConnection con = DB.getConnection();
            SqlDataAdapter cmd = new SqlDataAdapter(sql, con);
            DataTable dt = new DataTable();
            try
            {
                con.Open();
                cmd.Fill(dt);
                cmd.Dispose();
                con.Close();
            }
            catch (Exception ex)
            { }
            return dt;
        }
        public static DataTable TienDoLT1()
        {
            DBConnection DB = new DBConnection();
            string sql;
            sql = "SELECT DISTINCT TT,MA,CHITIEU,THANG,QUY,KHQUY,TLQUY,NAM,KHNAM,TLNAM FROM TienDoLT1 ORDER BY TT ASC";

            SqlConnection con = DB.getConnection();
            SqlDataAdapter cmd = new SqlDataAdapter(sql, con);
            DataTable dt = new DataTable();
            try
            {
                con.Open();
                cmd.Fill(dt);
                cmd.Dispose();
                con.Close();
            }
            catch (Exception ex)
            { }
            return dt;
        }
        public ActionResult TienDo(string id)
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
        public ActionResult ImportTienDoLT(WebReport.Models.Importsql importExcel)
        {
            DBConnection DB = new DBConnection();
            if (ModelState.IsValid)
            {

                SqlConnection con = DB.getConnection();
                con.Open();
                try
                {
                    SqlCommand cmd4 = new SqlCommand("delete from TienDoLT1", con);
                    cmd4.ExecuteNonQuery();
                    SqlCommand cmd2 = new SqlCommand("SET IDENTITY_INSERT TienDoLT OFF ", con);
                    cmd2.ExecuteNonQuery();
                }
                catch (Exception ex) { }
                SqlCommand cmd1 = new SqlCommand("select * into TienDoLT1 from TienDoLT where 1=1", con);
                cmd1.ExecuteNonQuery();
                con.Close();

                string path = Server.MapPath("~/Content/Upload/" + importExcel.file.FileName);
                importExcel.file.SaveAs(path);

                string excelConnectionString = @"Provider='Microsoft.ACE.OLEDB.12.0';Data Source='" + path + "';Extended Properties='Excel 12.0 Xml;IMEX=1'";
                OleDbConnection excelConnection = new OleDbConnection(excelConnectionString);

                //Sheet Name
                excelConnection.Open();
                string tableName = excelConnection.GetSchema("Tables").Rows[0]["TABLE_NAME"].ToString();
                excelConnection.Close();
                //End

                OleDbCommand cmd = new OleDbCommand("Select * from [" + tableName + "]", excelConnection);

                excelConnection.Open();

                OleDbDataReader dReader;
                dReader = cmd.ExecuteReader();
                SqlBulkCopy sqlBulk = new SqlBulkCopy(ConfigurationManager.ConnectionStrings["Demo1"].ConnectionString);


                //Give your Destination table name
                sqlBulk.DestinationTableName = "TienDoLT1";

                //Mappings
                sqlBulk.ColumnMappings.Add("TT", "TT");
                sqlBulk.ColumnMappings.Add("MA", "MA");
                sqlBulk.ColumnMappings.Add("CHITIEU", "CHITIEU");
                sqlBulk.ColumnMappings.Add("THANG", "THANG");
                sqlBulk.ColumnMappings.Add("QUY", "QUY");
                sqlBulk.ColumnMappings.Add("KHQUY", "KHQUY");
                sqlBulk.ColumnMappings.Add("TLQUY", "TLQUY");
                sqlBulk.ColumnMappings.Add("NAM", "NAM");
                sqlBulk.ColumnMappings.Add("KHNAM", "KHNAM");
                sqlBulk.ColumnMappings.Add("TLNAM", "TLNAM");
                sqlBulk.WriteToServer(dReader);
                excelConnection.Close();

                ViewBag.Result = "Successfully Imported";
            }
            return RedirectToAction("CapNhat");
        }
        [HttpPost]
        public ActionResult ImportTienDoCBQL(WebReport.Models.Importsql importExcel)
        {
            DBConnection DB = new DBConnection();
            if (ModelState.IsValid)
            {

                SqlConnection con = DB.getConnection();
                con.Open();
                try
                {
                    SqlCommand cmd4 = new SqlCommand("delete from TienDoCBQL1", con);
                    cmd4.ExecuteNonQuery();
                    SqlCommand cmd2 = new SqlCommand("SET IDENTITY_INSERT TienDoCBQL OFF ", con);
                    cmd2.ExecuteNonQuery();
                }
                catch (Exception ex) { }
                SqlCommand cmd1 = new SqlCommand("select * into TienDoCBQL1 from TienDoCBQL where 1=1", con);
                cmd1.ExecuteNonQuery();
                con.Close();

                string path = Server.MapPath("~/Content/Upload/" + importExcel.file.FileName);
                importExcel.file.SaveAs(path);

                string excelConnectionString = @"Provider='Microsoft.ACE.OLEDB.12.0';Data Source='" + path + "';Extended Properties='Excel 12.0 Xml;IMEX=1'";
                OleDbConnection excelConnection = new OleDbConnection(excelConnectionString);

                //Sheet Name
                excelConnection.Open();
                string tableName = excelConnection.GetSchema("Tables").Rows[0]["TABLE_NAME"].ToString();
                excelConnection.Close();
                //End

                OleDbCommand cmd = new OleDbCommand("Select * from [" + tableName + "]", excelConnection);

                excelConnection.Open();

                OleDbDataReader dReader;
                dReader = cmd.ExecuteReader();
                SqlBulkCopy sqlBulk = new SqlBulkCopy(ConfigurationManager.ConnectionStrings["Demo1"].ConnectionString);


                //Give your Destination table name
                sqlBulk.DestinationTableName = "TienDoCBQL1";

                //Mappings
                sqlBulk.ColumnMappings.Add("TT", "TT");
                sqlBulk.ColumnMappings.Add("MA", "MA");
                sqlBulk.ColumnMappings.Add("CHITIEU", "CHITIEU");
                sqlBulk.ColumnMappings.Add("THANG", "THANG");
                sqlBulk.ColumnMappings.Add("QUY", "QUY");
                sqlBulk.ColumnMappings.Add("KHQUY", "KHQUY");
                sqlBulk.ColumnMappings.Add("TLQUY", "TLQUY");
                sqlBulk.ColumnMappings.Add("NAM", "NAM");
                sqlBulk.ColumnMappings.Add("KHNAM", "KHNAM");
                sqlBulk.ColumnMappings.Add("TLNAM", "TLNAM");

                sqlBulk.WriteToServer(dReader);
                excelConnection.Close();

                ViewBag.Result = "Successfully Imported";
            }
            return RedirectToAction("CapNhat");
        }
       
        [HttpPost]
        public ActionResult ResetTienDoCBQL()
        {
            DBConnection DB = new DBConnection();
            SqlConnection con = DB.getConnection();
            con.Open();
            SqlCommand cmd4 = new SqlCommand("delete from TienDoCBQL", con);
            cmd4.ExecuteNonQuery();
            try
            {
                SqlCommand cmd2 = new SqlCommand("SET IDENTITY_INSERT TienDoCBQL ON ", con);
                cmd2.ExecuteNonQuery();
            }
            catch (Exception ex) { }
            
            SqlCommand cmd1 = new SqlCommand("Insert into TienDoCBQL ([TT],[MA],[CHITIEU],[THANG],[QUY],[KHQUY],[TLQUY],[NAM],[KHNAM],[TLNAM]) select [TT],[MA],[CHITIEU],[THANG],[QUY],[KHQUY],[TLQUY],[NAM],[KHNAM],[TLNAM] from TienDoCBQL1 ", con);
            cmd1.ExecuteNonQuery();
            SqlCommand cmd3 = new SqlCommand("drop table TienDoCBQL1", con);
            cmd3.ExecuteNonQuery();
            con.Close();
            return RedirectToAction("CapNhat");
        }
        //import tien do MLNS
        [HttpPost]
        public ActionResult ImportTienDoMLNS(WebReport.Models.Importsql importExcel)
        {
            DBConnection DB = new DBConnection();
            if (ModelState.IsValid)
            {

                SqlConnection con = DB.getConnection();
                con.Open();
                try
                {
                    SqlCommand cmd4 = new SqlCommand("delete from TienDoMLNS1", con);
                    cmd4.ExecuteNonQuery();
                    SqlCommand cmd2 = new SqlCommand("SET IDENTITY_INSERT TienDoMLNS OFF ", con);
                    cmd2.ExecuteNonQuery();
                }
                catch(Exception ex) { }
                SqlCommand cmd1 = new SqlCommand("select * into TienDoMLNS1 from TienDoMLNS where 1=1", con);
                cmd1.ExecuteNonQuery();
                con.Close();

                string path = Server.MapPath("~/Content/Upload/" + importExcel.file.FileName);
                importExcel.file.SaveAs(path);

                string excelConnectionString = @"Provider='Microsoft.ACE.OLEDB.12.0';Data Source='" + path + "';Extended Properties='Excel 12.0 Xml;IMEX=1'";
                OleDbConnection excelConnection = new OleDbConnection(excelConnectionString);

                //Sheet Name
                excelConnection.Open();
                string tableName = excelConnection.GetSchema("Tables").Rows[0]["TABLE_NAME"].ToString();
                excelConnection.Close();
                //End

                OleDbCommand cmd = new OleDbCommand("Select * from [" + tableName + "]", excelConnection);

                excelConnection.Open();

                OleDbDataReader dReader;
                dReader = cmd.ExecuteReader();
                SqlBulkCopy sqlBulk = new SqlBulkCopy(ConfigurationManager.ConnectionStrings["Demo1"].ConnectionString);


                //Give your Destination table name
                sqlBulk.DestinationTableName = "TienDoMLNS1";

                //Mappings
                sqlBulk.ColumnMappings.Add("TT", "TT");
                sqlBulk.ColumnMappings.Add("MA", "MA");
                sqlBulk.ColumnMappings.Add("CHITIEU", "CHITIEU");
                sqlBulk.ColumnMappings.Add("THANG", "THANG");
                sqlBulk.ColumnMappings.Add("QUY", "QUY");
                sqlBulk.ColumnMappings.Add("KHQUY", "KHQUY");
                sqlBulk.ColumnMappings.Add("TLQUY", "TLQUY");
                sqlBulk.ColumnMappings.Add("NAM", "NAM");
                sqlBulk.ColumnMappings.Add("KHNAM", "KHNAM");
                sqlBulk.ColumnMappings.Add("TLNAM", "TLNAM");

                sqlBulk.WriteToServer(dReader);
                excelConnection.Close();

                ViewBag.Result = "Successfully Imported";
            }
            return RedirectToAction("CapNhat");
        }

        [HttpPost]
        public ActionResult ResetTienDoMLNS()
        {
            DBConnection DB = new DBConnection();
            SqlConnection con = DB.getConnection();
            con.Open();
            SqlCommand cmd4 = new SqlCommand("delete from TienDoMLNS", con);
            cmd4.ExecuteNonQuery();
            //SqlCommand cmd2 = new SqlCommand("SET IDENTITY_INSERT TienDoMLNS ON ", con);
            //cmd2.ExecuteNonQuery();
            SqlCommand cmd1 = new SqlCommand("Insert into TienDoMLNS ([TT],[MA],[CHITIEU],[THANG],[QUY],[KHQUY],[TLQUY],[NAM],[KHNAM],[TLNAM]) select [TT],[MA],[CHITIEU],[THANG],[QUY],[KHQUY],[TLQUY],[NAM],[KHNAM],[TLNAM] from TienDoMLNS1 ", con);
            cmd1.ExecuteNonQuery();
            SqlCommand cmd3 = new SqlCommand("drop table TienDoMLNS1", con);
            cmd3.ExecuteNonQuery();
            con.Close();
            return RedirectToAction("CapNhat");
        }
        //import tien do LT
        [HttpPost]
        public ActionResult ImportTienDoLT1(WebReport.Models.Importsql importExcel)
        {
            DBConnection DB = new DBConnection();
            if (ModelState.IsValid)
            {

                SqlConnection con = DB.getConnection();
                con.Open();
                try
                {
                    SqlCommand cmd4 = new SqlCommand("delete from TienDoLT1", con);
                    cmd4.ExecuteNonQuery();
                    SqlCommand cmd2 = new SqlCommand("SET IDENTITY_INSERT TienDoLT OFF ", con);
                    cmd2.ExecuteNonQuery();
                }
                catch (Exception ex) { }
                SqlCommand cmd1 = new SqlCommand("select * into TienDoLT1 from TienDoLT where 1=1", con);
                cmd1.ExecuteNonQuery();
                con.Close();

                string path = Server.MapPath("~/Content/Upload/" + importExcel.file.FileName);
                importExcel.file.SaveAs(path);

                string excelConnectionString = @"Provider='Microsoft.ACE.OLEDB.12.0';Data Source='" + path + "';Extended Properties='Excel 12.0 Xml;IMEX=1'";
                OleDbConnection excelConnection = new OleDbConnection(excelConnectionString);

                //Sheet Name
                excelConnection.Open();
                string tableName = excelConnection.GetSchema("Tables").Rows[0]["TABLE_NAME"].ToString();
                excelConnection.Close();
                //End

                OleDbCommand cmd = new OleDbCommand("Select * from [" + tableName + "]", excelConnection);

                excelConnection.Open();

                OleDbDataReader dReader;
                dReader = cmd.ExecuteReader();
                SqlBulkCopy sqlBulk = new SqlBulkCopy(ConfigurationManager.ConnectionStrings["Demo1"].ConnectionString);


                //Give your Destination table name
                sqlBulk.DestinationTableName = "TienDoLT1";

                //Mappings
                sqlBulk.ColumnMappings.Add("TT", "TT");
                sqlBulk.ColumnMappings.Add("MA", "MA");
                sqlBulk.ColumnMappings.Add("CHITIEU", "CHITIEU");
                sqlBulk.ColumnMappings.Add("THANG", "THANG");
                sqlBulk.ColumnMappings.Add("QUY", "QUY");
                sqlBulk.ColumnMappings.Add("KHQUY", "KHQUY");
                sqlBulk.ColumnMappings.Add("TLQUY", "TLQUY");
                sqlBulk.ColumnMappings.Add("NAM", "NAM");
                sqlBulk.ColumnMappings.Add("KHNAM", "KHNAM");
                sqlBulk.ColumnMappings.Add("TLNAM", "TLNAM");

                sqlBulk.WriteToServer(dReader);
                excelConnection.Close();

                ViewBag.Result = "Successfully Imported";
            }
            return RedirectToAction("CapNhat");
        }

        [HttpPost]
        public ActionResult ResetTienDoLT()
        {
            DBConnection DB = new DBConnection();
            SqlConnection con = DB.getConnection();
            con.Open();
            SqlCommand cmd4 = new SqlCommand("delete from TienDoLT", con);
            cmd4.ExecuteNonQuery();
            //SqlCommand cmd2 = new SqlCommand("SET IDENTITY_INSERT TienDoMLNS ON ", con);
            //cmd2.ExecuteNonQuery();
            SqlCommand cmd1 = new SqlCommand("Insert into TienDoLT ([TT],[MA],[CHITIEU],[THANG],[QUY],[KHQUY],[TLQUY],[NAM],[KHNAM],[TLNAM]) select [TT],[MA],[CHITIEU],[THANG],[QUY],[KHQUY],[TLQUY],[NAM],[KHNAM],[TLNAM] from TienDoLT1 ", con);
            cmd1.ExecuteNonQuery();
            SqlCommand cmd3 = new SqlCommand("drop table TienDoLT1", con);
            cmd3.ExecuteNonQuery();
            con.Close();
            return RedirectToAction("CapNhat");
        }
        //file mau
        [HttpPost]
        public ActionResult WriteDataToExcel()
        {
            DataTable dt = new DataTable();
            dt.TableName = "File example";
            //Add Columns  
            dt.Columns.Add("TT", typeof(int));
            dt.Columns.Add("MA", typeof(string));
            dt.Columns.Add("CHITIEU", typeof(string));
            dt.Columns.Add("THANG", typeof(string));
            dt.Columns.Add("QUY", typeof(string));
            dt.Columns.Add("KHQUY", typeof(string));
            dt.Columns.Add("TLQUY", typeof(string));
            dt.Columns.Add("NAM", typeof(string));
            dt.Columns.Add("KHNAM", typeof(string));
            dt.Columns.Add("TLNAM", typeof(string));
            dt.AcceptChanges();
            //Name of File  
            string fileName = "FileExample.xlsx";
            using (XLWorkbook wb = new XLWorkbook())
            {
                //Add DataTable in worksheet  
                wb.Worksheets.Add(dt);
                using (MemoryStream stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    //Return xlsx Excel File  
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.template", fileName);
                }
            }
        }
        [HttpPost]
        public ActionResult WriteTienDoCBQL()
        {
            DBConnection DB = new DBConnection();
            string sql;
            sql = "SELECT * FROM TienDoCBQL";
            SqlConnection con = DB.getConnection();
            SqlDataAdapter cmd = new SqlDataAdapter(sql, con);
            DataTable dt = new DataTable();
            dt.TableName = "TienDoCBQL";
            con.Open();
           cmd.Fill(dt);
           cmd.Dispose();
           con.Close();
           
            //Name of File  
            string fileName = "TienDoCBQL.xlsx";
            using (XLWorkbook wb = new XLWorkbook())
            {
                //Add DataTable in worksheet  
                wb.Worksheets.Add(dt);
                using (MemoryStream stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    //Return xlsx Excel File  
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.template", fileName);
                }
            }
        }
        [HttpPost]
        public ActionResult WriteTienDoMLNS()
        {
            DBConnection DB = new DBConnection();
            string sql;
            sql = "SELECT * FROM TienDoMLNS";
            SqlConnection con = DB.getConnection();
            SqlDataAdapter cmd = new SqlDataAdapter(sql, con);
            DataTable dt = new DataTable();
            dt.TableName = "TienDoMLNS";
          
            con.Open();
            cmd.Fill(dt);
            cmd.Dispose();
            con.Close();
            
            //Name of File  
            string fileName = "TienDoMLNS.xlsx";
            using (XLWorkbook wb = new XLWorkbook())
            {
                //Add DataTable in worksheet  
                wb.Worksheets.Add(dt);
                using (MemoryStream stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    //Return xlsx Excel File  
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.template", fileName);
                }
            }
        }
        [HttpPost]
        public ActionResult WriteTienDoLT()
        {
            DBConnection DB = new DBConnection();
            string sql;
            sql = "SELECT * FROM TienDoLT";
            SqlConnection con = DB.getConnection();
            SqlDataAdapter cmd = new SqlDataAdapter(sql, con);
            DataTable dt = new DataTable();
            dt.TableName = "TienDoLT";

            con.Open();
            cmd.Fill(dt);
            cmd.Dispose();
            con.Close();

            //Name of File  
            string fileName = "TienDoLT.xlsx";
            using (XLWorkbook wb = new XLWorkbook())
            {
                //Add DataTable in worksheet  
                wb.Worksheets.Add(dt);
                using (MemoryStream stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    //Return xlsx Excel File  
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.template", fileName);
                }
            }
        }
        [HttpPost]
        public ActionResult ExitTienDoLT()
        {
            DBConnection DB = new DBConnection();
            SqlConnection con = DB.getConnection();
            con.Open();
          
            SqlCommand cmd3 = new SqlCommand("drop table TienDoLT1", con);
            cmd3.ExecuteNonQuery();
            con.Close();
            return RedirectToAction("CapNhat");
        }
        [HttpPost]
        public ActionResult ExitTienDoCBQL()
        {
            DBConnection DB = new DBConnection();
            SqlConnection con = DB.getConnection();
            con.Open();
           
            SqlCommand cmd3 = new SqlCommand("drop table TienDoCBQL1", con);
            cmd3.ExecuteNonQuery();
            con.Close();
            return RedirectToAction("CapNhat");
        }
        [HttpPost]
        public ActionResult ExitTienDoMLNS()
        {
            DBConnection DB = new DBConnection();
            SqlConnection con = DB.getConnection();
            con.Open();
           
            SqlCommand cmd3 = new SqlCommand("drop table TienDoMLNS1", con);
            cmd3.ExecuteNonQuery();
            con.Close();
            return RedirectToAction("CapNhat");
        }
    }
}