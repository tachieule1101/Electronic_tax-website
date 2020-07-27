using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace WebReport.Models
{
    public class QLTienDoMLNS
    {

        public string TT { set; get; }
        public string MA { set; get; }
        public string CHITIEU { set; get; }
        public string THANG { set; get; }
        public string QUY { set; get; }
        public string NAM { set; get; }
        public string KHQUY { set; get; }
        public string KHNAM { set; get; }

    }
    class QLtableDSML
    {

        DBConnection DB;
        public QLtableDSML()
        {
            DB = new DBConnection();

        }
        public List<QLTienDoMLNS> get()
        {
            string sql;

            sql = "SELECT * FROM TienDoMLNS";


            List<QLTienDoMLNS> strTa = new List<QLTienDoMLNS>();
            SqlConnection con = DB.getConnection();
            SqlDataAdapter cmd = new SqlDataAdapter(sql, con);
            DataTable dt = new DataTable();

            con.Open();
            cmd.Fill(dt);
            cmd.Dispose();
            con.Close();
            QLTienDoMLNS strH;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                strH = new QLTienDoMLNS();
                strH.TT = (i + 1) + "";
                strH.MA = dt.Rows[i]["MA"].ToString();
                strH.CHITIEU = dt.Rows[i]["CHITIEU"].ToString();

                strH.THANG = Convert.ToDecimal(long.Parse(dt.Rows[i]["THANG"].ToString())).ToString("#,##0");

                strH.QUY = Convert.ToDecimal(long.Parse(dt.Rows[i]["QUY"].ToString())).ToString("#,##0");
                strH.NAM = Convert.ToDecimal(long.Parse(dt.Rows[i]["NAM"].ToString())).ToString("#,##0");
                strH.KHQUY = Convert.ToDecimal(long.Parse(dt.Rows[i]["KHQUY"].ToString())).ToString("#,##0");
                strH.KHNAM = Convert.ToDecimal(long.Parse(dt.Rows[i]["KHNAM"].ToString())).ToString("#,##0");
                strTa.Add(strH);

            }
            return strTa;
        }
        public DataTable kqTienDoMLNS()
        {
            string sql;

            sql = "SELECT * FROM TienDoMLNS";
            SqlConnection con = DB.getConnection();
            SqlDataAdapter cmd = new SqlDataAdapter(sql, con);
            DataTable dt = new DataTable();

            con.Open();
            cmd.Fill(dt);
            cmd.Dispose();
            con.Close();
            return dt;
        }

    }
}