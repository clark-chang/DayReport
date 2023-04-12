using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using DayReport.Utility;
using System.Diagnostics;

namespace DayReport.Models
{
    public class Login
    { 
        public int Validate(string username, string password)
        {
            int rt = -1;
            Database db = new Database();
            OleDbConnection odcnn = db.GetOleDbConnection("ORACLE_DB_HO");
            OleDbCommand odcmm = new OleDbCommand();
            OleDbDataAdapter oddap = null;
            OleDbDataReader oddrd = null;
            DataTable dt = new DataTable();
            string sql =  @"	SELECT CLERK_ID, PASSWD FROM MAST.PASSWD1  " +
                          @"	WHERE 1=1                                  " +
                          @"	AND CLERK_ID = '{0}'                       " +
                          @"	AND PASSWD = '{1}'                         " ;
            string formatsql = string.Format(sql, username, password); 

            try
            {
                
                odcmm.Connection = odcnn;
                odcmm.CommandType = CommandType.Text;
                odcmm.CommandText = formatsql;
                odcnn.Open();
                oddrd = odcmm.ExecuteReader();
                while (oddrd.Read())
                {
                    rt = 1;
                }
                //oddap = new OleDbDataAdapter(odcmm);
                //oddap.Fill(dt);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Debug.WriteLine(ex.Message);
            }
            finally
            {
                if (oddap != null) { oddap.Dispose(); }
                if (odcmm != null) { odcmm.Dispose(); }
                if (odcnn != null) { odcnn.Dispose(); }
            }
            //if(dt.Rows.Count == 1)
            //{
            //    rt = 1;
            //}
            //else
            //{
            //    rt = -1;
            //}
            return rt;
        }
    }      
}
