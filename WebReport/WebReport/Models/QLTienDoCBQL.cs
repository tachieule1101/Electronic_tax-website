using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace WebReport.Models
{
    public class QLTienDoCBQL
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
    class QLtableDS
    {

        DBConnection DB;
        public QLtableDS()
        {
            DB = new DBConnection();

        }
        public List<QLTienDoCBQL> get()
        {
            string sql;

            sql = "SELECT * FROM TienDoCBQL";


            List<QLTienDoCBQL> strTa = new List<QLTienDoCBQL>();
            SqlConnection con = DB.getConnection();
            SqlDataAdapter cmd = new SqlDataAdapter(sql, con);
            DataTable dt = new DataTable();

            con.Open();
            cmd.Fill(dt);
            cmd.Dispose();
            con.Close();
            QLTienDoCBQL strH;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                strH = new QLTienDoCBQL();
                strH.TT = (i + 1) + "";
                strH.MA = dt.Rows[i]["MA"].ToString();
                strH.CHITIEU = dt.Rows[i]["CHITIEU"].ToString();
                strH.THANG = Convert.ToDecimal((long.Parse(dt.Rows[i]["THANG"].ToString())) / 1000000).ToString("#,##0");
                strH.QUY = Convert.ToDecimal((long.Parse(dt.Rows[i]["QUY"].ToString())) / 1000000).ToString("#,##0");
                strH.NAM = Convert.ToDecimal((long.Parse(dt.Rows[i]["NAM"].ToString())) / 1000000).ToString("#,##0");
                strH.KHQUY = Convert.ToDecimal((long.Parse(dt.Rows[i]["KHQUY"].ToString())) / 1000000).ToString("#,##0");
                strH.KHNAM = Convert.ToDecimal((long.Parse(dt.Rows[i]["KHNAM"].ToString())) / 1000000).ToString("#,##0");
                strTa.Add(strH);
            }
            return strTa;
        }

        public DataTable kqTienDoCBQL()
        {
            string sql;

            sql = "SELECT * FROM TienDoCBQL";
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