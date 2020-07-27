using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace WebReport.Models
{
    public class DBConnection
    {
        string a = WebReport.Controllers.HomeController.connect;
        string strCon;
        public DBConnection()
        {

            strCon = ConfigurationManager.ConnectionStrings[a].ConnectionString;
        }
        public SqlConnection getConnection()
        {
            return new SqlConnection(strCon);
        }
    }
}