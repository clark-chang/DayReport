using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Data.OleDb;
using DayReport.Utility;
using System.Data;

namespace DayReport.Models
{
    public class Ysh

    {
        
        public DataTable Hitodayin()
        {
            //查詢資料
            Database db = new Database();
            OleDbConnection odcnn = db.GetOleDbConnection("ORACLE_DB_HO");
            OleDbCommand odcmm = new OleDbCommand();
            OleDbDataAdapter oddap = null;
            DataTable dt = new DataTable();
            try
            {

                odcnn.Open();
                odcmm.Connection = odcnn;
                odcmm.CommandType = CommandType.Text;
                odcmm.CommandText =
                                      @"   
                                         SELECT ""院區"",                                                                     
                                                ""科別名稱"",                                                                 
                                                ""醫師名稱"",                                                                 
                                                Count(*) AS ""小計""                                                          
                                         FROM   (SELECT CASE                                                                
                                                          WHEN bed_no NOT LIKE 'B%' THEN '員榮'                             
                                                          WHEN bed_no LIKE 'B%' THEN '員生'                                 
                                                        END             AS ""院區"",                                        
                                                        b.div_full_name AS ""科別名稱"",                                     
                                                        c.doctor_name   AS ""醫師名稱""                                      
                                                 FROM   ipd.ptipd @hi a                                                     
                                                        left join mast.div b                                                
                                                               ON a.div_no = b.div_no                                       
                                                        left join mast.doctor c                                             
                                                               ON a.vs_no = c.doctor_no                                     
                                                 WHERE  1 = 1                                                               
                                                        AND A.status <> '7'                                                 
                                                        AND a.bed_no LIKE 'B%'                                          
                                                        AND To_date(( To_number(Replace(A.admit_date, '.', ''))             
                                                                      + 19110000 ), 'yyyy-mm-dd') = Trunc(SYSDATE))         
                                         GROUP  BY ""院區"",                                                                
                                                   ""科別名稱"",                                                              
                                                   ""醫師名稱""
                                      ";
                oddap = new OleDbDataAdapter(odcmm);
                oddap.Fill(dt);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (oddap != null) { oddap.Dispose(); }
                if (odcmm != null) { odcmm.Dispose(); }
                if (odcnn != null) { odcnn.Dispose(); }
            }
            //計算小計
            try
            {
                int inttotal = 0;
                string subtotal = "小計";
                foreach (DataRow row in dt.Rows)
                {
                    inttotal += Convert.ToInt32(row["小計"]);
                }
                DataRow dataRow = dt.NewRow();
                dataRow["醫師名稱"] = subtotal;
                dataRow["小計"] = inttotal;
                dt.Rows.Add(dataRow);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                ;
            }
            return dt;
        }


        public DataTable Hitodayout()
        {
            //查詢資料
            Database db = new Database();
            OleDbConnection odcnn = db.GetOleDbConnection("ORACLE_DB_HO");
            OleDbCommand odcmm = new OleDbCommand();
            OleDbDataAdapter oddap = null;
            DataTable dt = new DataTable();
            try
            {
                odcnn.Open();
                odcmm.Connection = odcnn;
                odcmm.CommandType = CommandType.Text;
                odcmm.CommandText =
                                             @"  
                                                 SELECT ""院區"",                                                                         
                                                        ""科別名稱"",                                                                     
                                                        ""醫師名稱"",                                                                     
                                                        Count(*) AS ""小計""                                                              
                                                 FROM   (SELECT CASE                                                                      
                                                                  WHEN bed_no NOT LIKE 'B%' THEN '員榮'                                   
                                                                  WHEN bed_no LIKE 'B%' THEN '員生'                                       
                                                                END             AS ""院區"",                                              
                                                                b.div_full_name AS ""科別名稱"",                                          
                                                                c.doctor_name   AS ""醫師名稱""                                           
                                                         FROM   ipd.ptipd @hi a                                                           
                                                                left join mast.div b                                                      
                                                                       ON a.div_no = b.div_no                                             
                                                                left join mast.doctor c                                                   
                                                                       ON a.vs_no = c.doctor_no                                           
                                                         WHERE  1 = 1                                                                     
                                                                AND a.discharge_date NOT IN ( '0' )                                       
                                                                AND bed_no LIKE 'B%'                                                  
                                                                AND To_date(( To_number(Replace(A.discharge_date, '.', ''))               
                                                                              + 19110000 ), 'yyyy-mm-dd') = Trunc(SYSDATE))               
                                                 GROUP  BY ""院區"",                                                                      
                                                           ""科別名稱"",                                                                  
                                                           ""醫師名稱""                                                                   
                                             ";

                oddap = new OleDbDataAdapter(odcmm);
                oddap.Fill(dt);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (oddap != null) { oddap.Dispose(); }
                if (odcmm != null) { odcmm.Dispose(); }
                if (odcnn != null) { odcnn.Dispose(); }
            }

            //計算小計
            try
            {
                int inttotal = 0;
                string subtotal = "小計";
                DataRow dr = dt.NewRow();
                foreach (DataRow row in dt.Rows)
                {
                    inttotal += Convert.ToInt32(row["小計"]);
                }
                dr["醫師名稱"] = subtotal;
                dr["小計"] = inttotal;
                dt.Rows.Add(dr);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                ;
            }
            return dt;
        }


        public DataTable Hinow()
        {
            //查詢資料
            Database db = new Database();
            OleDbConnection odcnn = db.GetOleDbConnection("ORACLE_DB_HO");
            OleDbCommand odcmm = new OleDbCommand();
            OleDbDataAdapter oddap = null;
            DataTable dt = new DataTable();

            try
            {
                odcnn.Open();
                odcmm.Connection = odcnn;
                odcmm.CommandType = CommandType.Text;
                odcmm.CommandText =
                                           @"
                                             SELECT   ""院區"",                                                                     
                                                        ""科別名稱"",                                                                 
                                                        ""醫師名稱"",                                                                 
                                                        count(""醫師名稱"") AS ""小計""                                                
                                               FROM     (                                                                             
                                                                  SELECT                                                              
                                                                            CASE                                                      
                                                                                      WHEN bed_no NOT LIKE 'B%' THEN '員榮'           
                                                                                      WHEN bed_no LIKE 'B%' THEN '員生'               
                                                                            END             AS ""院區"",                              
                                                                            b.div_full_name AS ""科別名稱"",                          
                                                                            c.doctor_name   AS ""醫師名稱""                           
                                                                  FROM      ipd.ptipd @hi a                                           
                                                                  LEFT JOIN mast.div b                                                
                                                                  ON        a.div_no = b.div_no                                       
                                                                  LEFT JOIN mast.doctor c                                             
                                                                  ON        a.vs_no = c.doctor_no                                     
                                                                  WHERE     1=1                                                       
                                                                  AND       a.discharge_date = '0'                                    
                                                                  AND       bed_no LIKE 'B%' )                                    
                                               GROUP BY ""院區"",                                                                     
                                                        ""科別名稱"",                                                                 
                                                        ""醫師名稱""                                                                  
                                               ORDER BY ""院區"" DESC
                                          ";

                oddap = new OleDbDataAdapter(odcmm);
                oddap.Fill(dt);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (oddap != null) { oddap.Dispose(); }
                if (odcmm != null) { odcmm.Dispose(); }
                if (odcnn != null) { odcnn.Dispose(); }
            }
            //計算小計
            try
            {
                int total = 0;
                string subtotal = "小計";
                foreach (DataRow row in dt.Rows)
                {
                    total += Convert.ToInt32(row["小計"]);
                }
                DataRow dataRow = dt.NewRow();
                dataRow["醫師名稱"] = subtotal;
                dataRow["小計"] = total;
                dt.Rows.Add(dataRow);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                ;
            }
            return dt;
        }

        public DataTable Nursestation1()
        {
            string yr2f = "ICU病房(YR)";
            string yr3f = "3F病房(YR)";
            string yr5f = "5F病房(YR)";
            string yr6f = "呼吸照護病房(YR)";
            string ys5f = "ICU病房(YS)";
            string ys6f = "6F病房(YS)";
            string ys7f = "7F病房(YS)";

            Database db = new Database();
            OleDbConnection odcnn = db.GetOleDbConnection("ORACLE_DB_HO");
            OleDbCommand odcmm = new OleDbCommand();
            OleDbDataAdapter oddap = null;
            DataTable dt = new DataTable();

            try
            {
                odcnn.Open();
                odcmm.Connection = odcnn;
                odcmm.CommandType = CommandType.Text;
                odcmm.CommandText =

                       @"
                            SELECT                                                                                       
                                      CASE SUBSTR(Y.ns_code,1,1)                                                         
                                        WHEN 'B' THEN '員生'                                                                
                                          ELSE '員榮'                                                                       
                                            END AS ""院區"",                                                             
                                      CASE Y.ns_code                                                                     
                                       WHEN 'BI05' THEN '5F病房(YS)'                                                     
                                       WHEN 'BI06' THEN '6F病房(YS)'                                                     
                                       WHEN 'BI07' THEN '7F病房(YS)'                                                     
                                       WHEN 'BI05' THEN '5F病房(YS)'                                                     
                                       WHEN 'BICU' THEN 'ICU病房(YS)'                                                    
                                       WHEN 'ICU' THEN 'ICU病房(YR)'                                                     
                                       WHEN 'I02' THEN '2F病房(YR)'                                                      
                                       WHEN 'I03' THEN '3F病房(YR)'                                                      
                                       WHEN 'I05' THEN '5F病房(YR)'                                                      
                                       WHEN 'I06' THEN '6F病房(YR)'                                                      
                                       WHEN 'I07' THEN '7F病房(YR)'                                                      
                                       WHEN 'RT' THEN '呼吸照護病房(YR)'                                                  
                                     END                                                                   AS            
                                     ""護理站"",                                                                          
                                     ""健保"",                                                                           
                                     ""套房"",                                                                           
                                     ""合計"",                                                                           
                                     ""實際可用床數"",                                                                    
                                     Concat(To_char(Round(""合計"" / ""實際可用床數"" * 100, 2)), '%') AS                 
                                     ""佔床率""                                                                           
                              FROM  (SELECT T.ns_code,                                                                   
                                            Count(Y.nh_paid_flag) AS ""健保"",                                           
                                            Count(U.nh_paid_flag) AS ""套房"",                                           
                                            Count(x.admit_no)     AS ""合計""                                            
                                     FROM   ipd.ptipd@hi x                                                               
                                            left join ipd.bed@hi T                                                       
                                                   ON X.bed_no = T.bed_no                                                
                                            left join(SELECT A.bed_no,                                                   
                                                             A.ns_code,                                                  
                                                             C.exclusive_ward_flag,                                      
                                                             B.grade_code,                                               
                                                             B.statistic_grade,                                          
                                                             B.description,                                              
                                                             B.nh_paid_flag                                              
                                                      FROM   ipd.bed@hi A                                                
                                                             left join ipd.bedgrade1@hi B                                
                                                                    ON A.grade_code = B.grade_code                       
                                                             --NH_PAID_FLAG Y: 健保保險床 N: 非健保保險床(差額床)          
                                                             left join ipd.bedgrade2@hi C                                
                                                                    ON B.grade_code = C.grade_code                       
                                                      --EXCLUSIVE_WARD_FLAG 0:一般 1:包房加收 2:隔離加收                   
                                                      --REVENUE_TYPE 1:自費 2:健保                                        
                                                      WHERE  A.effective_date = (SELECT Max(Z.effective_date)            
                                                                                 FROM   ipd.bed@hi Z                     
                                                                                 WHERE  A.bed_no = Z.bed_no)             
                                                             AND B.nh_paid_flag = 'Y'                                    
                                                      GROUP  BY A.bed_no,                                                
                                                                c.exclusive_ward_flag,                                   
                                                                B.grade_code,                                            
                                                                B.statistic_grade,                                       
                                                                B.description,                                           
                                                                B.nh_paid_flag,                                          
                                                                A.ns_code                                                
                                                      ORDER  BY A.bed_no) y                                              
                                                   ON x.bed_no = y.bed_no                                                
                                                      AND x.exclusive_ward_flag = y.exclusive_ward_flag                  
                                            left join(SELECT A.bed_no,                                                   
                                                             A.ns_code,                                                  
                                                             c.exclusive_ward_flag,                                      
                                                             B.grade_code,                                               
                                                             B.statistic_grade,                                          
                                                             B.description,                                              
                                                             B.nh_paid_flag                                              
                                                      --A.STATISTIC_FLAG  AS ""佔床率""                                   
                                                      FROM   ipd.bed @hi A                                               
                                                             left join ipd.bedgrade1 @hi B                               
                                                                    ON A.grade_code = B.grade_code                       
                                                             --NH_PAID_FLAG Y:健保保險床 N:非健保保險床(差額床)            
                                                             left join ipd.bedgrade2 @hi C                               
                                                                    ON B.grade_code = C.grade_code                       
                                                      --EXCLUSIVE_WARD_FLAG 0:一般 1:包房加收 2:隔離加收                   
                                                      --REVENUE_TYPE 1:自費 2:健保                                        
                                                      WHERE  A.effective_date = (SELECT Max(Z.effective_date)            
                                                                                 FROM   ipd.bed @hi Z                    
                                                                                 WHERE  A.bed_no = Z.bed_no)             
                                                             AND B.nh_paid_flag = 'N'                                    
                                                      GROUP  BY A.bed_no,                                                
                                                                c.exclusive_ward_flag,                                   
                                                                B.grade_code,                                            
                                                                B.statistic_grade,                                       
                                                                B.description,                                           
                                                                B.nh_paid_flag,                                          
                                                                A.ns_code                                                
                                                      ORDER  BY A.bed_no) U                                              
                                                   ON x.bed_no = U.bed_no                                                
                                                      AND x.exclusive_ward_flag = U.exclusive_ward_flag                  
                                     WHERE  x.discharge_date = '0'                                                       
                                            AND T.effective_date = (SELECT Max(S.effective_date)                         
                                                                    FROM   ipd.bed @hi S                                 
                                                                    WHERE  T.bed_no = S.bed_no)                          
                                            AND x.status = '1'                                                           
                                            AND x.bed_no LIKE 'B%'                                                   
                                     GROUP  BY T.ns_code,                                                                
                                               Substr(x.bed_no, 1, 1)) Y                                                 
                                    left join(SELECT a.ns_code,                                                          
                                                     a.bed_amt            AS ""衛福部登記開床數"",                        
                                                     a.nh_bed_amt         AS ""向健保署報備總床數"",                       
                                                     a.real_bed_amt       AS ""實際可用床數"",                            
                                                     a.real_empty_bed_amt AS ""實際可用空床數"",                          
                                                     a.nh_paid_bed_amt    AS                                             
                                                     ""向健保署報備的健保床數"",                                           
                                                     a.nh_diff_bed_amt    AS                                             
                                                     ""向健保署報備的差額床數""                                            
                                              FROM   ipd.nsdivs@hi A                                                     
                                              WHERE  a.effective_date = (SELECT Max(Z.effective_date)                    
                                                                         FROM   ipd.nsdivs@hi z                          
                                                                         WHERE  a.ns_code = Z.ns_code)                   
                                              ORDER  BY a.ns_code)X                                                      
                                           ON Y.ns_code = X.ns_code
                       ";

                oddap = new OleDbDataAdapter(odcmm);
                oddap.Fill(dt);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (oddap != null) { oddap.Dispose(); }
                if (odcmm != null) { odcmm.Dispose(); }
                if (odcnn != null) { odcnn.Dispose(); }
            }
            return dt;
        }



        public DataTable Nursestation2()
        {
            string bed = "一般急性病房";
            Double actual = 0;
            int health = 0;
            int suite = 0;
            Double total = 0;
            string rate;

            Database db = new Database();
            OleDbConnection odcnn = db.GetOleDbConnection("ORACLE_DB_HO");
            OleDbCommand odcmm = new OleDbCommand();
            OleDbDataAdapter oddap = null;
            DataTable dt = new DataTable();
            DataTable dt1 = new DataTable();

            try
            {
                odcnn.Open();
                odcmm.Connection = odcnn;
                odcmm.CommandType = CommandType.Text;
                odcmm.CommandText =

                       @"
                            SELECT                                                                                         
                                      CASE SUBSTR(Y.ns_code,1,1)                                                           
                                        WHEN 'B' THEN 'B'                                                                  
                                          ELSE 'A'                                                                         
                                            END AS ""院區"",                                                               
                                      CASE Y.ns_code                                                                       
                                       WHEN 'BI05' THEN '5F病房(YS)'                                                       
                                       WHEN 'BI06' THEN '6F病房(YS)'                                                       
                                       WHEN 'BI07' THEN '7F病房(YS)'                                                       
                                       WHEN 'BI05' THEN '5F病房(YS)'                                                       
                                       WHEN 'BICU' THEN 'ICU病房(YS)'                                                      
                                       WHEN 'ICU' THEN 'ICU病房(YR)'                                                       
                                       WHEN 'I02' THEN '2F病房(YR)'                                                        
                                       WHEN 'I03' THEN '3F病房(YR)'                                                        
                                       WHEN 'I05' THEN '5F病房(YR)'                                                        
                                       WHEN 'I06' THEN '6F病房(YR)'                                                        
                                       WHEN 'I07' THEN '7F病房(YR)'                                                        
                                       WHEN 'RT' THEN '呼吸照護病房(YR)'                                                   
                                     END                                                                   AS              
                                     ""護理站"",                                                                           
                                     ""健保"",                                                                             
                                     ""套房"",                                                                             
                                     ""合計"",                                                                             
                                     ""實際可用床數"",                                                                     
                                     Concat(To_char(Round(""合計"" / ""實際可用床數"" * 100, 2)), '%') AS                 
                                     ""佔床率""                                                                            
                              FROM  (SELECT T.ns_code,                                                                     
                                            Count(Y.nh_paid_flag) AS ""健保"",                                             
                                            Count(U.nh_paid_flag) AS ""套房"",                                             
                                            Count(x.admit_no)     AS ""合計""                                              
                                     FROM   ipd.ptipd@hi x                                                                 
                                            left join ipd.bed@hi T                                                         
                                                   ON X.bed_no = T.bed_no                                                  
                                            left join(SELECT A.bed_no,                                                     
                                                             A.ns_code,                                                    
                                                             C.exclusive_ward_flag,                                        
                                                             B.grade_code,                                                 
                                                             B.statistic_grade,                                            
                                                             B.description,                                                
                                                             B.nh_paid_flag                                                
                                                      FROM   ipd.bed@hi A                                                  
                                                             left join ipd.bedgrade1@hi B                                  
                                                                    ON A.grade_code = B.grade_code                         
                                                             --NH_PAID_FLAG Y: 健保保險床 N: 非健保保險床(差額床)            
                                                             left join ipd.bedgrade2@hi C                                  
                                                                    ON B.grade_code = C.grade_code                         
                                                      --EXCLUSIVE_WARD_FLAG 0:一般 1:包房加收 2:隔離加收                     
                                                      --REVENUE_TYPE 1:自費 2:健保                                         
                                                      WHERE  A.effective_date = (SELECT Max(Z.effective_date)              
                                                                                 FROM   ipd.bed@hi Z                       
                                                                                 WHERE  A.bed_no = Z.bed_no)               
                                                             AND B.nh_paid_flag = 'Y'                                      
                                                      GROUP  BY A.bed_no,                                                  
                                                                c.exclusive_ward_flag,                                     
                                                                B.grade_code,                                              
                                                                B.statistic_grade,                                         
                                                                B.description,                                             
                                                                B.nh_paid_flag,                                            
                                                                A.ns_code                                                  
                                                      ORDER  BY A.bed_no) y                                                
                                                   ON x.bed_no = y.bed_no                                                  
                                                      AND x.exclusive_ward_flag = y.exclusive_ward_flag                    
                                            left join(SELECT A.bed_no,                                                     
                                                             A.ns_code,                                                    
                                                             c.exclusive_ward_flag,                                        
                                                             B.grade_code,                                                 
                                                             B.statistic_grade,                                            
                                                             B.description,                                                
                                                             B.nh_paid_flag                                                
                                                      --A.STATISTIC_FLAG  AS ""佔床率""                                    
                                                      FROM   ipd.bed @hi A                                                 
                                                             left join ipd.bedgrade1 @hi B                                 
                                                                    ON A.grade_code = B.grade_code                         
                                                             --NH_PAID_FLAG Y:健保保險床 N:非健保保險床(差額床)             
                                                             left join ipd.bedgrade2 @hi C                                 
                                                                    ON B.grade_code = C.grade_code                         
                                                      --EXCLUSIVE_WARD_FLAG 0:一般 1:包房加收 2:隔離加收                    
                                                      --REVENUE_TYPE 1:自費 2:健保                                         
                                                      WHERE  A.effective_date = (SELECT Max(Z.effective_date)              
                                                                                 FROM   ipd.bed @hi Z                      
                                                                                 WHERE  A.bed_no = Z.bed_no)               
                                                             AND B.nh_paid_flag = 'N'                                      
                                                      GROUP  BY A.bed_no,                                                  
                                                                c.exclusive_ward_flag,                                     
                                                                B.grade_code,                                              
                                                                B.statistic_grade,                                         
                                                                B.description,                                             
                                                                B.nh_paid_flag,                                            
                                                                A.ns_code                                                  
                                                      ORDER  BY A.bed_no) U                                                
                                                   ON x.bed_no = U.bed_no                                                  
                                                      AND x.exclusive_ward_flag = U.exclusive_ward_flag                    
                                     WHERE  x.discharge_date = '0'                                                         
                                            AND T.effective_date = (SELECT Max(S.effective_date)                           
                                                                    FROM   ipd.bed @hi S                                   
                                                                    WHERE  T.bed_no = S.bed_no)                            
                                            AND x.status = '1'                                                             
                                            AND x.bed_no LIKE 'B%'                                                     
                                     GROUP  BY T.ns_code,                                                                  
                                               Substr(x.bed_no, 1, 1)) Y                                                   
                                    left join(SELECT a.ns_code,                                                            
                                                     a.bed_amt            AS ""衛福部登記開床數"",                          
                                                     a.nh_bed_amt         AS ""向健保署報備總床數"",                        
                                                     a.real_bed_amt       AS ""實際可用床數"",                             
                                                     a.real_empty_bed_amt AS ""實際可用空床數"",                           
                                                     a.nh_paid_bed_amt    AS                                               
                                                     ""向健保署報備的健保床數"",                                            
                                                     a.nh_diff_bed_amt    AS                                               
                                                     ""向健保署報備的差額床數""                                              
                                              FROM   ipd.nsdivs@hi A                                                       
                                              WHERE  a.effective_date = (SELECT Max(Z.effective_date)                      
                                                                         FROM   ipd.nsdivs@hi z                            
                                                                         WHERE  a.ns_code = Z.ns_code)                     
                                              ORDER  BY a.ns_code)X                                                        
                                           ON Y.ns_code = X.ns_code
                       ";
                oddap = new OleDbDataAdapter(odcmm);
                oddap.Fill(dt);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (oddap != null) { oddap.Dispose(); }
                if (odcmm != null) { odcmm.Dispose(); }
                if (odcnn != null) { odcnn.Dispose(); }
            }

            try
            {
                foreach (DataRow dr in dt.Rows)
                {
                    if (dr["護理站"].ToString() == "3F病房(YR)")
                    {
                        actual = actual + Convert.ToInt32(dr["實際可用床數"]);
                        health = health + Convert.ToInt32(dr["健保"]);
                        suite = suite + Convert.ToInt32(dr["套房"]);
                        total = health + suite;
                    }
                    else if (dr["護理站"].ToString() == "5F病房(YR)")
                    {
                        actual = actual + Convert.ToInt32(dr["實際可用床數"]);
                        health = health + Convert.ToInt32(dr["健保"]);
                        suite = suite + Convert.ToInt32(dr["套房"]);
                        total = health + suite;
                    }
                    else if (dr["護理站"].ToString() == "6F病房(YS)")
                    {
                        actual = actual + Convert.ToInt32(dr["實際可用床數"]);
                        health = health + Convert.ToInt32(dr["健保"]);
                        suite = suite + Convert.ToInt32(dr["套房"]);
                        total = health + suite;
                    }
                    else if (dr["護理站"].ToString() == "7F病房(YS)")
                    {
                        actual = actual + Convert.ToInt32(dr["實際可用床數"]);
                        health = health + Convert.ToInt32(dr["健保"]);
                        suite = suite + Convert.ToInt32(dr["套房"]);
                        total = health + suite;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
            finally
            {
                ;
            }

            try
            {
                dt1.Columns.Add("院區", typeof(string));
                dt1.Columns.Add("護理站", typeof(string));
                dt1.Columns.Add("健保", typeof(int));
                dt1.Columns.Add("套房", typeof(int));
                dt1.Columns.Add("合計", typeof(Double));
                dt1.Columns.Add("實際可用床數", typeof(Double));
                dt1.Columns.Add("佔床率", typeof(string));

                DataRow newrow = dt1.NewRow();

                newrow["院區"] = "員生";
                newrow["護理站"] = bed;
                newrow["健保"] = health;
                newrow["套房"] = suite;
                newrow["合計"] = total;
                newrow["實際可用床數"] = actual;
                newrow["佔床率"] = (Math.Round(total / actual * 100, 2)).ToString() + "%";
                dt1.Rows.Add(newrow);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                ;
            }
            return dt1;
        }


        public DataTable Nursestation3()
        {
            Database db = new Database();
            OleDbConnection odcnn = db.GetOleDbConnection("ORACLE_DB_HO");
            OleDbCommand odcmm = new OleDbCommand();
            OleDbDataAdapter oddap = null;
            DataTable dt = new DataTable();
            DataTable dt1 = new DataTable();

            string bed = "ICU加護病房";
            Double actual = 0;
            int health = 0;
            int suite = 0;
            Double total = 0;
            string rate;

            try
            {
                odcnn.Open();
                odcmm.Connection = odcnn;
                odcmm.CommandType = CommandType.Text;
                odcmm.CommandText =

                        @"	
                            SELECT                                                                                         
                                    CASE SUBSTR(Y.ns_code,1,1)                                                           
                                      WHEN 'B' THEN '員生'                                                                  
                                        ELSE '員榮'                                                                         
                                          END AS ""院區"",                                                               
                                    CASE Y.ns_code                                                                       
                                     WHEN 'BI05' THEN '5F病房(YS)'                                                       
                                     WHEN 'BI06' THEN '6F病房(YS)'                                                       
                                     WHEN 'BI07' THEN '7F病房(YS)'                                                       
                                     WHEN 'BI05' THEN '5F病房(YS)'                                                       
                                     WHEN 'BICU' THEN 'ICU病房(YS)'                                                      
                                     WHEN 'ICU' THEN 'ICU病房(YR)'                                                       
                                     WHEN 'I02' THEN '2F病房(YR)'                                                        
                                     WHEN 'I03' THEN '3F病房(YR)'                                                        
                                     WHEN 'I05' THEN '5F病房(YR)'                                                        
                                     WHEN 'I06' THEN '6F病房(YR)'                                                        
                                     WHEN 'I07' THEN '7F病房(YR)'                                                        
                                     WHEN 'RT' THEN '呼吸照護病房(YR)'                                                   
                                   END                                                                   AS              
                                   ""護理站"",                                                                           
                                   ""健保"",                                                                             
                                   ""套房"",                                                                             
                                   ""合計"",                                                                             
                                   ""實際可用床數"",                                                                     
                                   Concat(To_char(Round(""合計"" / ""實際可用床數"" * 100, 2)), '%') AS                  
                                   ""佔床率""                                                                            
                            FROM  (SELECT T.ns_code,                                                                     
                                          Count(Y.nh_paid_flag) AS ""健保"",                                             
                                          Count(U.nh_paid_flag) AS ""套房"",                                             
                                          Count(x.admit_no)     AS ""合計""                                              
                                   FROM   ipd.ptipd@hi x                                                                 
                                          left join ipd.bed@hi T                                                         
                                                 ON X.bed_no = T.bed_no                                                  
                                          left join(SELECT A.bed_no,                                                     
                                                           A.ns_code,                                                    
                                                           C.exclusive_ward_flag,                                        
                                                           B.grade_code,                                                 
                                                           B.statistic_grade,                                            
                                                           B.description,                                                
                                                           B.nh_paid_flag                                                
                                                    FROM   ipd.bed@hi A                                                  
                                                           left join ipd.bedgrade1@hi B                                  
                                                                  ON A.grade_code = B.grade_code                         
                                                           --NH_PAID_FLAG Y: 健保保險床 N: 非健保保險床(差額床)            
                                                           left join ipd.bedgrade2@hi C                                  
                                                                  ON B.grade_code = C.grade_code                         
                                                    --EXCLUSIVE_WARD_FLAG 0:一般 1:包房加收 2:隔離加收                     
                                                    --REVENUE_TYPE 1:自費 2:健保                                         
                                                    WHERE  A.effective_date = (SELECT Max(Z.effective_date)              
                                                                               FROM   ipd.bed@hi Z                       
                                                                               WHERE  A.bed_no = Z.bed_no)               
                                                           AND B.nh_paid_flag = 'Y'                                      
                                                    GROUP  BY A.bed_no,                                                  
                                                              c.exclusive_ward_flag,                                     
                                                              B.grade_code,                                              
                                                              B.statistic_grade,                                         
                                                              B.description,                                             
                                                              B.nh_paid_flag,                                            
                                                              A.ns_code                                                  
                                                    ORDER  BY A.bed_no) y                                                
                                                 ON x.bed_no = y.bed_no                                                  
                                                    AND x.exclusive_ward_flag = y.exclusive_ward_flag                    
                                          left join(SELECT A.bed_no,                                                     
                                                           A.ns_code,                                                    
                                                           c.exclusive_ward_flag,                                        
                                                           B.grade_code,                                                 
                                                           B.statistic_grade,                                            
                                                           B.description,                                                
                                                           B.nh_paid_flag                                                
                                                    --A.STATISTIC_FLAG  AS ""佔床率""                                    
                                                    FROM   ipd.bed @hi A                                                 
                                                           left join ipd.bedgrade1 @hi B                                 
                                                                  ON A.grade_code = B.grade_code                         
                                                           --NH_PAID_FLAG Y:健保保險床 N:非健保保險床(差額床)            
                                                           left join ipd.bedgrade2 @hi C                                 
                                                                  ON B.grade_code = C.grade_code                         
                                                    --EXCLUSIVE_WARD_FLAG 0:一般 1:包房加收 2:隔離加收                   
                                                    --REVENUE_TYPE 1:自費 2:健保                                         
                                                    WHERE  A.effective_date = (SELECT Max(Z.effective_date)              
                                                                               FROM   ipd.bed @hi Z                      
                                                                               WHERE  A.bed_no = Z.bed_no)               
                                                           AND B.nh_paid_flag = 'N'                                      
                                                    GROUP  BY A.bed_no,                                                  
                                                              c.exclusive_ward_flag,                                     
                                                              B.grade_code,                                              
                                                              B.statistic_grade,                                         
                                                              B.description,                                             
                                                              B.nh_paid_flag,                                            
                                                              A.ns_code                                                  
                                                    ORDER  BY A.bed_no) U                                                
                                                 ON x.bed_no = U.bed_no                                                  
                                                    AND x.exclusive_ward_flag = U.exclusive_ward_flag                    
                                   WHERE  x.discharge_date = '0'                                                         
                                          AND T.effective_date = (SELECT Max(S.effective_date)                           
                                                                  FROM   ipd.bed @hi S                                   
                                                                  WHERE  T.bed_no = S.bed_no)                            
                                          AND x.status = '1'                                                             
                                          AND x.bed_no LIKE 'B%'                                                     
                                   GROUP  BY T.ns_code,                                                                  
                                             Substr(x.bed_no, 1, 1)) Y                                                   
                                  left join(SELECT a.ns_code,                                                            
                                                   a.bed_amt            AS ""衛福部登記開床數"",                          
                                                   a.nh_bed_amt         AS ""向健保署報備總床數"",                        
                                                   a.real_bed_amt       AS ""實際可用床數"",                             
                                                   a.real_empty_bed_amt AS ""實際可用空床數"",                           
                                                   a.nh_paid_bed_amt    AS                                               
                                                   ""向健保署報備的健保床數"",                                            
                                                   a.nh_diff_bed_amt    AS                                               
                                                   ""向健保署報備的差額床數""                                              
                                            FROM   ipd.nsdivs@hi A                                                       
                                            WHERE  a.effective_date = (SELECT Max(Z.effective_date)                      
                                                                       FROM   ipd.nsdivs@hi z                            
                                                                       WHERE  a.ns_code = Z.ns_code)                     
                                            ORDER  BY a.ns_code)X                                                        
                                         ON Y.ns_code = X.ns_code                                                        
                        ";

                oddap = new OleDbDataAdapter(odcmm);
                oddap.Fill(dt);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (oddap != null) { oddap.Dispose(); }
                if (odcmm != null) { odcmm.Dispose(); }
                if (odcnn != null) { odcnn.Dispose(); }
            }



            try
            {
                foreach (DataRow dr in dt.Rows)
                {
                    if (dr["護理站"].ToString() == "ICU病房(YR)")
                    {
                        actual = actual + Convert.ToInt32(dr["實際可用床數"]);
                        health = health + Convert.ToInt32(dr["健保"]);
                        suite = suite + Convert.ToInt32(dr["套房"]);
                        total = health + suite;
                    }
                    else if (dr["護理站"].ToString() == "ICU病房(YS)")
                    {
                        actual = actual + Convert.ToInt32(dr["實際可用床數"]);
                        health = health + Convert.ToInt32(dr["健保"]);
                        suite = suite + Convert.ToInt32(dr["套房"]);
                        total = health + suite;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
            finally
            {
                ;
            }

            try
            {
                dt1.Columns.Add("院區", typeof(string));
                dt1.Columns.Add("護理站", typeof(string));
                dt1.Columns.Add("健保", typeof(int));
                dt1.Columns.Add("套房", typeof(int));
                dt1.Columns.Add("合計", typeof(Double));
                dt1.Columns.Add("實際可用床數", typeof(Double));
                dt1.Columns.Add("佔床率", typeof(string));

                DataRow newrow = dt1.NewRow();

                newrow["院區"] = "員生";
                newrow["護理站"] = bed;
                newrow["健保"] = health;
                newrow["套房"] = suite;
                newrow["合計"] = total;
                newrow["實際可用床數"] = actual;
                newrow["佔床率"] = (Math.Round(total / actual * 100, 2)).ToString() + "%";
                dt1.Rows.Add(newrow);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                ;
            }
            return dt1;
        }

        public DataTable Nursestation4()
        {
            Database db = new Database();
            OleDbConnection odcnn = db.GetOleDbConnection("ORACLE_DB_HO");
            OleDbCommand odcmm = new OleDbCommand();
            OleDbDataAdapter oddap = null;
            DataTable dt = new DataTable();
            DataTable dt1 = new DataTable();

            string bed = "呼吸照護病房";
            Double actual = 0;
            int health = 0;
            int suite = 0;
            Double total = 0;
            string rate;

            try
            {
                odcnn.Open();
                odcmm.Connection = odcnn;
                odcmm.CommandType = CommandType.Text;
                odcmm.CommandText =

                          @"
                        	SELECT                                                                                         
                                    CASE SUBSTR(Y.ns_code,1,1)                                                           
                                      WHEN 'B' THEN 'B'                                                                  
                                        ELSE 'A'                                                                         
                                          END AS ""院區"",                                                               
                                    CASE Y.ns_code                                                                       
                                     WHEN 'BI05' THEN '5F病房(YS)'                                                       
                                     WHEN 'BI06' THEN '6F病房(YS)'                                                       
                                     WHEN 'BI07' THEN '7F病房(YS)'                                                       
                                     WHEN 'BI05' THEN '5F病房(YS)'                                                       
                                     WHEN 'BICU' THEN 'ICU病房(YS)'                                                      
                                     WHEN 'ICU' THEN 'ICU病房(YR)'                                                       
                                     WHEN 'I02' THEN '2F病房(YR)'                                                        
                                     WHEN 'I03' THEN '3F病房(YR)'                                                        
                                     WHEN 'I05' THEN '5F病房(YR)'                                                        
                                     WHEN 'I06' THEN '6F病房(YR)'                                                        
                                     WHEN 'I07' THEN '7F病房(YR)'                                                        
                                     WHEN 'RT' THEN '呼吸照護病房(YR)'                                                   
                                   END                                                                   AS              
                                   ""護理站"",                                                                           
                                   ""健保"",                                                                             
                                   ""套房"",                                                                             
                                   ""合計"",                                                                             
                                   ""實際可用床數"",                                                                     
                                   Concat(To_char(Round(""合計"" / ""實際可用床數"" * 100, 2)), '%') AS                  
                                   ""佔床率""                                                                            
                            FROM  (SELECT T.ns_code,                                                                     
                                          Count(Y.nh_paid_flag) AS ""健保"",                                             
                                          Count(U.nh_paid_flag) AS ""套房"",                                             
                                          Count(x.admit_no)     AS ""合計""                                              
                                   FROM   ipd.ptipd@hi x                                                                 
                                          left join ipd.bed@hi T                                                         
                                                 ON X.bed_no = T.bed_no                                                  
                                          left join(SELECT A.bed_no,                                                     
                                                           A.ns_code,                                                    
                                                           C.exclusive_ward_flag,                                        
                                                           B.grade_code,                                                 
                                                           B.statistic_grade,                                            
                                                           B.description,                                                
                                                           B.nh_paid_flag                                                
                                                    FROM   ipd.bed@hi A                                                  
                                                           left join ipd.bedgrade1@hi B                                  
                                                                  ON A.grade_code = B.grade_code                         
                                                           --NH_PAID_FLAG Y: 健保保險床 N: 非健保保險床(差額床)            
                                                           left join ipd.bedgrade2@hi C                                  
                                                                  ON B.grade_code = C.grade_code                         
                                                    --EXCLUSIVE_WARD_FLAG 0:一般 1:包房加收 2:隔離加收                     
                                                    --REVENUE_TYPE 1:自費 2:健保                                         
                                                    WHERE  A.effective_date = (SELECT Max(Z.effective_date)              
                                                                               FROM   ipd.bed@hi Z                       
                                                                               WHERE  A.bed_no = Z.bed_no)               
                                                           AND B.nh_paid_flag = 'Y'                                      
                                                    GROUP  BY A.bed_no,                                                  
                                                              c.exclusive_ward_flag,                                     
                                                              B.grade_code,                                              
                                                              B.statistic_grade,                                         
                                                              B.description,                                             
                                                              B.nh_paid_flag,                                            
                                                              A.ns_code                                                  
                                                    ORDER  BY A.bed_no) y                                                
                                                 ON x.bed_no = y.bed_no                                                  
                                                    AND x.exclusive_ward_flag = y.exclusive_ward_flag                    
                                          left join(SELECT A.bed_no,                                                     
                                                           A.ns_code,                                                    
                                                           c.exclusive_ward_flag,                                        
                                                           B.grade_code,                                                 
                                                           B.statistic_grade,                                            
                                                           B.description,                                                
                                                           B.nh_paid_flag                                                
                                                    --A.STATISTIC_FLAG  AS ""佔床率""                                    
                                                    FROM   ipd.bed @hi A                                                 
                                                           left join ipd.bedgrade1 @hi B                                 
                                                                  ON A.grade_code = B.grade_code                         
                                                           --NH_PAID_FLAG Y:健保保險床 N:非健保保險床(差額床)               
                                                           left join ipd.bedgrade2 @hi C                                 
                                                                  ON B.grade_code = C.grade_code                         
                                                    --EXCLUSIVE_WARD_FLAG 0:一般 1:包房加收 2:隔離加收                     
                                                    --REVENUE_TYPE 1:自費 2:健保                                         
                                                    WHERE  A.effective_date = (SELECT Max(Z.effective_date)              
                                                                               FROM   ipd.bed @hi Z                      
                                                                               WHERE  A.bed_no = Z.bed_no)               
                                                           AND B.nh_paid_flag = 'N'                                      
                                                    GROUP  BY A.bed_no,                                                  
                                                              c.exclusive_ward_flag,                                     
                                                              B.grade_code,                                              
                                                              B.statistic_grade,                                         
                                                              B.description,                                             
                                                              B.nh_paid_flag,                                            
                                                              A.ns_code                                                  
                                                    ORDER  BY A.bed_no) U                                                
                                                 ON x.bed_no = U.bed_no                                                  
                                                    AND x.exclusive_ward_flag = U.exclusive_ward_flag                    
                                   WHERE  x.discharge_date = '0'                                                         
                                          AND T.effective_date = (SELECT Max(S.effective_date)                           
                                                                  FROM   ipd.bed @hi S                                   
                                                                  WHERE  T.bed_no = S.bed_no)                            
                                          AND x.status = '1'                                                             
                                          AND x.bed_no LIKE 'B%'                                                     
                                   GROUP  BY T.ns_code,                                                                  
                                             Substr(x.bed_no, 1, 1)) Y                                                   
                                  left join(SELECT a.ns_code,                                                            
                                                   a.bed_amt            AS ""衛福部登記開床數"",                          
                                                   a.nh_bed_amt         AS ""向健保署報備總床數"",                        
                                                   a.real_bed_amt       AS ""實際可用床數"",                             
                                                   a.real_empty_bed_amt AS ""實際可用空床數"",                           
                                                   a.nh_paid_bed_amt    AS                                               
                                                   ""向健保署報備的健保床數"",                                            
                                                   a.nh_diff_bed_amt    AS                                               
                                                   ""向健保署報備的差額床數""                                              
                                            FROM   ipd.nsdivs@hi A                                                       
                                            WHERE  a.effective_date = (SELECT Max(Z.effective_date)                      
                                                                       FROM   ipd.nsdivs@hi z                            
                                                                       WHERE  a.ns_code = Z.ns_code)                     
                                            ORDER  BY a.ns_code)X                                                        
                                         ON Y.ns_code = X.ns_code                                                        
                          ";
                oddap = new OleDbDataAdapter(odcmm);
                oddap.Fill(dt);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (oddap != null) { oddap.Dispose(); }
                if (odcmm != null) { odcmm.Dispose(); }
                if (odcnn != null) { odcnn.Dispose(); }
            }

            try
            {
                foreach (DataRow dr in dt.Rows)
                {
                    if (dr["護理站"].ToString() == "呼吸照護病房(YR)")
                    {
                        actual = actual + Convert.ToInt32(dr["實際可用床數"]);
                        health = health + Convert.ToInt32(dr["健保"]);
                        suite = suite + Convert.ToInt32(dr["套房"]);
                        total = health + suite;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
            finally
            {
                ;
            }

            try
            {
                dt1.Columns.Add("院區", typeof(string));
                dt1.Columns.Add("護理站", typeof(string));
                dt1.Columns.Add("健保", typeof(int));
                dt1.Columns.Add("套房", typeof(int));
                dt1.Columns.Add("合計", typeof(Double));
                dt1.Columns.Add("實際可用床數", typeof(Double));
                dt1.Columns.Add("佔床率", typeof(string));

                DataRow newrow = dt1.NewRow();

                newrow["院區"] = "員生";
                newrow["護理站"] = bed;
                newrow["健保"] = health;
                newrow["套房"] = suite;
                newrow["合計"] = total;
                newrow["實際可用床數"] = actual;
                newrow["佔床率"] = (Math.Round(total / actual * 100, 2)).ToString() + "%";
                dt1.Rows.Add(newrow);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                ;
            }
            return dt1;
        }


        public DataTable Nursestation5()
        {
            Database db = new Database();
            OleDbConnection odcnn = db.GetOleDbConnection("ORACLE_DB_HO");
            OleDbCommand odcmm = new OleDbCommand();
            OleDbDataAdapter oddap = null;
            DataTable dt = new DataTable();

            string nursestation;
            int healthcare;
            int suite;
            int total;
            int actual;


            try
            {
                odcnn.Open();
                odcmm.Connection = odcnn;
                odcmm.CommandType = CommandType.Text;
                odcmm.CommandText =
                                @"
                                    SELECT  '員生'                  AS ""院區"",
                                            '全部'                  AS ""護理站"",                                                                                     
                                           SUM(""實際可用床數"")     AS ""實際可用床數"",                                                                               
                                           SUM(""健保"")             AS ""健保"",                                                                                     
                                           SUM(""套房"")             AS ""套房"",                                                                                     
                                           SUM(""合計"")             AS ""合計"",                                                                                     
                                           Concat(To_char(Round(SUM(""合計"") / SUM(""實際可用床數"") * 100, 2)), '%') AS ""佔床率""                                   
                                    FROM                                                                                                                             
                                    (SELECT  ""健保"",                                                                                                                
                                             ""套房"",                                                                                                                
                                             ""合計"",                                                                                                                 
                                             ""實際可用床數"",                                                                                                          
                                             Concat(To_char(Round(""合計"" / ""實際可用床數"" * 100, 2)),'%') AS ""佔床率""                                             
                                             FROM  (SELECT T.ns_code,                                                                                                 
                                                           Count(Y.nh_paid_flag) AS ""健保"",                                                                          
                                                           Count(U.nh_paid_flag) AS ""套房"",                                                                         
                                                           Count(x.admit_no)     AS ""合計""                                                                          
                                                    FROM   ipd.ptipd@hi x                                                                                             
                                                           left join ipd.bed@hi T                                                                                      
                                                                  ON X.bed_no = T.bed_no                                                                              
                                                           left join(SELECT A.bed_no,                                                                                 
                                                                            A.ns_code,                                                                                
                                                                            C.exclusive_ward_flag,                                                                    
                                                                            B.grade_code,                                                                             
                                                                            B.statistic_grade,                                                                        
                                                                            B.description,                                                                            
                                                                            B.nh_paid_flag                                                                            
                                                                     FROM   ipd.bed@hi A                                                                              
                                                                            left join ipd.bedgrade1@hi B                                                              
                                                                                   ON A.grade_code = B.grade_code                                                     
                                                                            --NH_PAID_FLAG Y: 健保保險床 N: 非健保保險床(差額床)                                        
                                                                            left join ipd.bedgrade2@hi C                                                              
                                                                                   ON B.grade_code = C.grade_code                                                     
                                                                     --EXCLUSIVE_WARD_FLAG 0:一般 1:包房加收 2:隔離加收                                                 
                                                                     --REVENUE_TYPE 1:自費 2:健保                                                                     
                                                                     WHERE  A.effective_date =                                                                        
                                                                            (SELECT Max(Z.effective_date)                                                             
                                                                             FROM   ipd.bed@hi Z                                                                      
                                                                             WHERE  A.bed_no = Z.bed_no)                                                              
                                                                            AND B.nh_paid_flag = 'Y'                                                                  
                                                                     GROUP  BY A.bed_no,                                                                              
                                                                               c.exclusive_ward_flag,                                                                 
                                                                               B.grade_code,                                                                          
                                                                               B.statistic_grade,                                                                     
                                                                               B.description,                                                                         
                                                                               B.nh_paid_flag,                                                                        
                                                                               A.ns_code                                                                              
                                                                     ORDER  BY A.bed_no) y                                                                            
                                                                  ON x.bed_no = y.bed_no                                                                              
                                                                     AND x.exclusive_ward_flag =                                                                      
                                                                         y.exclusive_ward_flag                                                                        
                                                           left join(SELECT A.bed_no,                                                                                 
                                                                            A.ns_code,                                                                                
                                                                            c.exclusive_ward_flag,                                                                    
                                                                            B.grade_code,                                                                             
                                                                            B.statistic_grade,                                                                        
                                                                            B.description,                                                                            
                                                                            B.nh_paid_flag                                                                            
                                                                     --A.STATISTIC_FLAG  AS ""佔床率""                                                                 
                                                                     FROM   ipd.bed @hi A                                                                             
                                                                            left join ipd.bedgrade1 @hi B                                                              
                                                                                   ON A.grade_code = B.grade_code                                                     
                                                                            --NH_PAID_FLAG Y:健保保險床 N:非健保保險床(差額床)                                          
                                                                            left join ipd.bedgrade2 @hi C                                                             
                                                                                   ON B.grade_code = C.grade_code                                                     
                                                                     --EXCLUSIVE_WARD_FLAG 0:一般 1:包房加收 2:隔離加收                                               
                                                                     --REVENUE_TYPE 1:自費 2:健保                                                                     
                                                                     WHERE  A.effective_date =                                                                        
                                                                            (SELECT Max(Z.effective_date)                                                             
                                                                             FROM   ipd.bed @hi Z                                                                     
                                                                             WHERE  A.bed_no = Z.bed_no)                                                              
                                                                            AND B.nh_paid_flag = 'N'                                                                  
                                                                     GROUP  BY A.bed_no,                                                                              
                                                                               c.exclusive_ward_flag,                                                                 
                                                                               B.grade_code,                                                                          
                                                                               B.statistic_grade,                                                                     
                                                                               B.description,                                                                         
                                                                               B.nh_paid_flag,                                                                        
                                                                               A.ns_code                                                                              
                                                                     ORDER  BY A.bed_no) U                                                                            
                                                                  ON x.bed_no = U.bed_no                                                                              
                                                                     AND x.exclusive_ward_flag = U.exclusive_ward_flag                                                
                                                    WHERE  x.discharge_date = '0'                                                                                     
                                                           AND T.effective_date = (SELECT Max(S.effective_date)                                                       
                                                                                   FROM   ipd.bed @hi S                                                               
                                                                                   WHERE  T.bed_no = S.bed_no)                                                        
                                                    AND x.status = '1'                                                                                                
                                                    AND x.bed_no LIKE 'B%'                                                                                        
                                                    GROUP  BY T.ns_code,                                                                                              
                                                              Substr(x.bed_no, 1, 1)) Y                                                                               
                                                   left join(SELECT a.ns_code,                                                                                        
                                                                    a.bed_amt            AS ""衛福部登記開床數"",                                                       
                                                                    a.nh_bed_amt         AS ""向健保署報備總床數"",                                                     
                                                                    a.real_bed_amt       AS ""實際可用床數"",                                                           
                                                                    a.real_empty_bed_amt AS ""實際可用空床數"",                                                         
                                                                    a.nh_paid_bed_amt    AS ""向健保署報備的健保床數"",                                                 
                                                                    a.nh_diff_bed_amt    AS ""向健保署報備的差額床數""                                                  
                                                             FROM   ipd.nsdivs@hi A                                                                                   
                                                             WHERE  a.effective_date = (SELECT Max(Z.effective_date)                                                  
                                                                                        FROM   ipd.nsdivs@hi z                                                        
                                                                                        WHERE  a.ns_code =                                                            
                                                                                       Z.ns_code)                                                                     
                                                             ORDER  BY a.ns_code)X                                                                                    
                                                          ON Y.ns_code = X.ns_code)
                                 ";
                oddap = new OleDbDataAdapter(odcmm);
                oddap.Fill(dt);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (oddap != null) { oddap.Dispose(); }
                if (odcmm != null) { odcmm.Dispose(); }
            }
            return dt;
        }


        public DataTable Hotodayin()
        {
            Database db = new Database();
            OleDbConnection odcnn = db.GetOleDbConnection("ORACLE_DB_HO");
            OleDbCommand odcmm = new OleDbCommand();
            OleDbDataAdapter oddap = null;
            DataTable dt = new DataTable();

            try
            {
                odcnn.Open();
                odcmm.Connection = odcnn;
                odcmm.CommandType = CommandType.Text;
                odcmm.CommandText =
                                       @"
                                         SELECT  ""科別代碼與科別名稱"",                                                                                                        
                                                 ""診間號"",                                                                                                                   
                                                 ""醫師名稱"",                                                                                                                 
                                                 ""早診"",                                                                                                                     
                                                 ""早診初診人次"",                                                                                                             
                                                 ""午診"",                                                                                                                     
                                                 ""午診初診人次"",                                                                                                             
                                                 ""晚診"",                                                                                                                     
                                                 ""晚診初診人次"",                                                                                                             
                                                 COALESCE(""早診"", 0) + COALESCE(""午診"", 0)  + COALESCE(""晚診"", 0) AS ""合計""                                             
                                          FROM  (SELECT ""科別代碼與科別名稱"",                                                                                                 
                                                        ""診間號"",                                                                                                            
                                                        ""醫師名稱"",                                                                                                          
                                                        Sum(""早診"")       AS ""早診"",                                                                                       
                                                        Sum(""早診初診人次"") AS ""早診初診人次"",                                                                               
                                                        Sum(""午診"")       AS ""午診"",                                                                                       
                                                        Sum(""午診初診人次"") AS ""午診初診人次"",                                                                               
                                                        Sum(""晚診"")       AS ""晚診"",                                                                                       
                                                        Sum(""晚診初診人次"") AS ""晚診初診人次""                                                                                
                                                 FROM   (SELECT Concat(""科別代碼"", ""科別名稱"") AS                                                                           
                                                               ""科別代碼與科別名稱"",                                                                                          
                                                                ""診間號"",                                                                                                    
                                                                ""醫師名稱"",                                                                                                  
                                                                CASE ""午別""                                                                                                  
                                                                  WHEN '1' THEN ""已看診人數""                                                                                 
                                                                  ELSE 0                                                                                                      
                                                                END                                    AS ""早診"",                                                            
                                                                CASE ""午別""                                                                                                  
                                                                  WHEN '1' THEN ""已看初診人數""                                                                               
                                                                  ELSE 0                                                                                                      
                                                                END                                    AS ""早診初診人次"",                                                    
                                                                CASE ""午別""                                                                                                  
                                                                  WHEN '2' THEN ""已看診人數""                                                                                 
                                                                  ELSE 0                                                                                                      
                                                                END                                    AS ""午診"",                                                            
                                                                CASE ""午別""                                                                                                  
                                                                  WHEN '2' THEN ""已看初診人數""                                                                               
                                                                  ELSE 0                                                                                                      
                                                                END                                    AS ""午診初診人次"",                                                    
                                                                CASE ""午別""                                                                                                  
                                                                  WHEN '3' THEN ""已看診人數""                                                                                 
                                                                  ELSE 0                                                                                                      
                                                                END                                    AS ""晚診"",                                                            
                                                                CASE ""午別""                                                                                                  
                                                                  WHEN '3' THEN ""已看初診人數""                                                                               
                                                                  ELSE 0                                                                                                      
                                                                END                                    AS ""晚診初診人次"",                                                    
                                                                ""診別""                                                                                                       
                                                         FROM   (                                                                                                             
                                                                --門診無中醫--                                                                                                 
                                                                SELECT a.clinic_date                       AS ""就醫日期"",                                                    
                                                                       b.week                              AS ""星期"",                                                        
                                                                       a.clinic_apn                        AS ""午別"",                                                        
                                                                       a.clinic_no                         AS ""診間號"",                                                      
                                                                       b.div_no                            AS ""科別代碼"",                                                    
                                                                       b.clinic_name                       AS ""科別名稱"",                                                    
                                                                       d.doctor_name                       AS ""醫師名稱"",                                                    
                                                                       a.clinic_flag                       AS ""診別"",                                                        
                                                                       Count(*)                            AS ""已看診人數"",                                                  
                                                                       COALESCE(e.""已看初診人數"", 0)      AS ""已看初診人數""                                                  
                                                                FROM   opd.ptopd a                                                                                            
                                                                       LEFT JOIN opd.dclin b                                                                                  
                                                                              ON a.clinic_date = b.clinic_date                                                                
                                                                                 AND a.clinic_apn = b.clinic_apn                                                              
                                                                                 AND a.clinic_no = b.clinic_no                                                                
                                                                                 AND a.doctor_no = b.doctor_no                                                                
                                                                       LEFT JOIN mast.div c                                                                                   
                                                                              ON a.div_no = c.div_no                                                                          
                                                                       LEFT JOIN mast.doctor d                                                                                
                                                                              ON a.doctor_no = d.doctor_no                                                                    
                                                                       LEFT JOIN (SELECT                                                                                      
                                                                       a.clinic_date AS ""就醫日期"",                                                                          
                                                                       b.week        AS ""星期"",                                                                              
                                                                       a.clinic_apn  AS ""午別"",                                                                              
                                                                       a.clinic_no   AS ""診間號"",                                                                            
                                                                       b.div_no      AS ""科別代碼"",                                                                          
                                                                       b.clinic_name AS ""科別名稱"",                                                                          
                                                                       d.doctor_name AS ""醫師名稱"",                                                                          
                                                                       a.clinic_flag AS ""診別"",                                                                              
                                                                       Count(*)      AS ""已看初診人數""                                                                       
                                                                                  FROM   opd.ptopd a                                                                          
                                                                                         LEFT JOIN opd.dclin b                                                                
                                                                                                ON a.clinic_date =                                                            
                                                                                                   b.clinic_date                                                              
                                                                                                   AND a.clinic_apn =                                                         
                                                                                                       b.clinic_apn                                                           
                                                                                                   AND a.clinic_no =                                                          
                                                                                                       b.clinic_no                                                            
                                                                                                   AND a.doctor_no =                                                          
                                                                                                       b.doctor_no                                                            
                                                                                         LEFT JOIN mast.div c                                                                 
                                                                                                ON a.div_no = c.div_no                                                        
                                                                                         LEFT JOIN mast.doctor d                                                              
                                                                                                ON a.doctor_no =                                                              
                                                                                                   d.doctor_no                                                                
                                                                                  WHERE  To_date(To_char(To_number(a.clinic_date) + 19110000),'yyyy/MM/dd') = Trunc(sysdate)  
                                                                                         AND a.clinic_flag IN ( 'O', 'V' )                                                    
                                                                                         AND a.korder_flag IN ('Y', 'P', 'R' )                                                
                                                                                         AND a.reg_clerk NOT IN ('KOPDMISC', 'KOPDCHR','KOPDUREG' )                           
                                                                                         AND a.fv_rv_flag = '1'                                                               
                                                                                         AND a.clinic_no NOT IN ( 'DAS' )                                                     
                                                                                         AND a.building_no = 'B'                                                             
                                                                                  GROUP  BY a.clinic_date,                                                                    
                                                                                            b.week,                                                                           
                                                                                            a.clinic_apn,                                                                     
                                                                                            a.clinic_no,                                                                      
                                                                                            b.div_no,                                                                         
                                                                                            b.clinic_name,                                                                    
                                                                                            d.doctor_name,                                                                    
                                                                                            a.clinic_flag                                                                     
                                                                                  ORDER  BY a.clinic_flag,                                                                    
                                                                                            a.clinic_apn,                                                                     
                                                                                            a.clinic_no)e                                                                     
                                                                              ON a.clinic_date = e.""就醫日期""                                                                
                                                                                 AND b.week = e.""星期""                                                                       
                                                                                 AND a.clinic_apn = e.""午別""                                                                 
                                                                                 AND a.clinic_no = e.""診間號""                                                                
                                                                                 AND b.div_no = e.""科別代碼""                                                                 
                                                                                 AND b.clinic_name = e.""科別名稱""                                                            
                                                                                 AND d.doctor_name = e.""醫師名稱""                                                            
                                                                                 AND a.clinic_flag = e.""診別""                                                                
                                                                WHERE  To_date(To_char(To_number(a.clinic_date) + 19110000),'yyyy/MM/dd') = Trunc(sysdate)                    
                                                                       AND a.clinic_no != ' '                                                                                 
                                                                       AND A.div_no != '60'                                                                                   
                                                                       AND a.clinic_flag IN ( 'O', 'V' )                                                                      
                                                                       AND a.korder_flag IN ( 'Y', 'P', 'R' )                                                                 
                                                                       AND a.reg_clerk NOT IN ( 'KOPDMISC', 'KOPDCHR', 'KOPDREH','KOPDUREG' )                                 
                                                                       AND a.clinic_no NOT IN ( 'DAS' )                                                                       
                                                                       AND a.building_no = 'B'                                                                               
                                                                GROUP  BY a.clinic_date,                                                                                      
                                                                          b.week,                                                                                             
                                                                          a.clinic_apn,                                                                                       
                                                                          a.clinic_no,                                                                                        
                                                                          b.div_no,                                                                                           
                                                                          b.clinic_name,                                                                                      
                                                                          a.clinic_no,                                                                                        
                                                                          d.doctor_name,                                                                                      
                                                                          a.clinic_flag,                                                                                      
                                                                          e.""已看初診人數""                                                                                   
                                                                 UNION ALL                                                                                                    
                                                                 --門診中醫--                                                                                                  
                                                                 SELECT a.clinic_date                       AS ""就醫日期"",                                                   
                                                                        b.week                              AS ""星期"",                                                       
                                                                        a.clinic_apn                        AS ""午別"",                                                       
                                                                        a.clinic_no                         AS ""診間號"",                                                     
                                                                        b.div_no                            AS ""科別代碼"",                                                   
                                                                        b.clinic_name                       AS ""科別名稱"",                                                   
                                                                        d.doctor_name                       AS ""醫師名稱"",                                                   
                                                                        a.clinic_flag                       AS ""診別"",                                                       
                                                                        Count(*)                            AS ""已看診人數"",                                                 
                                                                 COALESCE(e.""已看初診人數"", 0) AS ""已看初診人數""                                                             
                                                                 FROM   opd.ptopd a                                                                                           
                                                                 LEFT JOIN opd.dclin b                                                                                        
                                                                        ON a.clinic_date = b.clinic_date                                                                      
                                                                           AND a.clinic_apn = b.clinic_apn                                                                    
                                                                           AND a.clinic_no = b.clinic_no                                                                      
                                                                           AND a.doctor_no = b.doctor_no                                                                      
                                                                 LEFT JOIN mast.div c                                                                                         
                                                                        ON a.div_no = c.div_no                                                                                
                                                                 LEFT JOIN mast.doctor d                                                                                      
                                                                        ON a.doctor_no = d.doctor_no                                                                          
                                                                 LEFT JOIN (SELECT a.clinic_date AS ""就醫日期"",                                                              
                                                                                   b.week        AS ""星期"",                                                                  
                                                                                   a.clinic_apn  AS ""午別"",                                                                  
                                                                                   a.clinic_no   AS ""診間號"",                                                                
                                                                                   b.div_no      AS ""科別代碼"",                                                              
                                                                                   b.clinic_name AS ""科別名稱"",                                                              
                                                                                   d.doctor_name AS ""醫師名稱"",                                                              
                                                                                   a.clinic_flag AS ""診別"",                                                                  
                                                                                   Count(*)      AS ""已看初診人數""                                                           
                                                                            FROM   opd.ptopd a                                                                                
                                                                                   LEFT JOIN opd.dclin b                                                                      
                                                                                          ON a.clinic_date = b.clinic_date                                                    
                                                                                             AND a.clinic_apn = b.clinic_apn                                                  
                                                                                             AND a.clinic_no = b.clinic_no                                                    
                                                                                             AND a.doctor_no = b.doctor_no                                                    
                                                                                   LEFT JOIN mast.div c                                                                       
                                                                                          ON a.div_no = c.div_no                                                              
                                                                                   LEFT JOIN mast.doctor d                                                                    
                                                                                          ON a.doctor_no = d.doctor_no                                                        
                                                                            WHERE  To_date(To_char(To_number(a.clinic_date) + 19110000), 'yyyy/MM/dd') = Trunc(sysdate)       
                                                                                   --AND a.course_seq IN ( '0', '1' )                                                         
                                                                                   AND A.div_no = '60'                                                                        
                                                                                   AND a.clinic_flag IN ( 'O', 'V' )                                                          
                                                                                   AND a.korder_flag IN ( 'Y', 'P', 'R' )                                                     
                                                                                   AND a.reg_clerk NOT IN ( 'KOPDMISC', 'KOPDCHR','KOPDREH','KOPDUREG' )                      
                                                                                   AND a.fv_rv_flag = '1'                                                                     
                                                                                   AND a.clinic_no NOT IN ( 'DAS' )                                                           
                                                                                   AND a.building_no = 'B'                                                                   
                                                                            GROUP  BY a.clinic_date,                                                                          
                                                                                      b.week,                                                                                 
                                                                                      a.clinic_apn,                                                                           
                                                                                      a.clinic_no,                                                                            
                                                                                      b.div_no,                                                                               
                                                                                      b.clinic_name,                                                                          
                                                                                      d.doctor_name,                                                                          
                                                                                      a.clinic_flag                                                                           
                                                                            ORDER  BY a.clinic_flag,                                                                          
                                                                                      a.clinic_apn,                                                                           
                                                                                      a.clinic_no)e                                                                           
                                                       ON a.clinic_date = e.""就醫日期""                                                                                       
                                                          AND b.week = e.""星期""                                                                                              
                                                          AND a.clinic_apn = e.""午別""                                                                                        
                                                          AND a.clinic_no = e.""診間號""                                                                                       
                                                          AND b.div_no = e.""科別代碼""                                                                                        
                                                          AND b.clinic_name = e.""科別名稱""                                                                                   
                                                          AND d.doctor_name = e.""醫師名稱""                                                                                   
                                                          AND a.clinic_flag = e.""診別""                                                                                       
                                                WHERE  To_date(To_char(To_number(a.clinic_date) + 19110000), 'yyyy/MM/dd') = Trunc(sysdate)                                   
                                                --AND a.course_seq IN ( '0', '1' )                                                                                            
                                                AND A.div_no = '60'                                                                                                           
                                                AND a.clinic_flag IN ( 'O', 'V' )                                                                                             
                                                AND a.korder_flag IN ( 'Y', 'P', 'R' )                                                                                        
                                                AND a.reg_clerk NOT IN ( 'KOPDMISC', 'KOPDCHR','KOPDREH','KOPDUREG' )                                                         
                                                AND a.clinic_no NOT IN ( 'DAS' )                                                                                              
                                                AND a.building_no = 'B'                                                                                                      
                                                GROUP  BY a.clinic_date,                                                                                                      
                                                   b.week,                                                                                                                    
                                                   a.clinic_apn,                                                                                                              
                                                   a.clinic_no,                                                                                                               
                                                   b.div_no,                                                                                                                  
                                                   b.clinic_name,                                                                                                             
                                                   a.clinic_no,                                                                                                               
                                                   d.doctor_name,                                                                                                             
                                                   a.clinic_flag,                                                                                                             
                                                   e.""已看初診人數""))                                                                                                        
                                                 GROUP  BY ""科別代碼與科別名稱"",                                                                                              
                                                           ""診間號"",                                                                                                         
                                                           ""醫師名稱"",                                                                                                       
                                                           ""診別""                                                                                                            
                                                 ORDER  BY ""科別代碼與科別名稱"",                                                                                              
                                                           ""診間號"")  
                                      
                                        ";

                oddap = new OleDbDataAdapter(odcmm);
                oddap.Fill(dt);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            finally
            {
                if (oddap != null) { oddap.Dispose(); }
                if (odcmm != null) { odcmm.Dispose(); }
                if (odcnn != null) { odcnn.Dispose(); }
            }
            return (dt);
        }


        public DataTable Emtoday()
        {
            Database db = new Database();
            OleDbConnection odcnn = db.GetOleDbConnection("ORACLE_DB_HO");
            OleDbCommand odcmm = new OleDbCommand();
            OleDbDataAdapter oddap = null;
            DataTable dt = new DataTable();

            try
            {
                odcnn.Open();
                odcmm.Connection = odcnn;
                odcmm.CommandType = CommandType.Text;
                odcmm.CommandText =
                                                @"
                                                     SELECT CASE A.building_no                                                               
                                                              WHEN 'A' THEN '員榮'                                                           
                                                              WHEN 'B' THEN '員生'                                                           
                                                            END             AS ""院區"",                                                     
                                                            A.clinic_date   AS ""看診日期"",                                                 
                                                            C.div_full_name AS ""科別全名"",                                                 
                                                            B.doctor_name   AS ""醫師姓名"",                                                 
                                                            Sum(CASE                                                                         
                                                                  WHEN A.reg_time BETWEEN 0000 AND 0800 THEN 1                               
                                                                  ELSE 0                                                                     
                                                                END)        AS ""大夜"",                                                     
                                                            Sum(CASE                                                                         
                                                                  WHEN A.reg_time BETWEEN 0801 AND 1600 THEN 1                               
                                                                  ELSE 0                                                                     
                                                               END)        AS ""早班"",                                                      
                                                            Sum(CASE                                                                         
                                                                  WHEN A.reg_time BETWEEN 1601 AND 2359 THEN 1                               
                                                                  ELSE 0                                                                     
                                                                END)        AS ""小夜""                                                      
                                                     FROM   opd.ptopd A                                                                      
                                                            LEFT JOIN mast.doctor B                                                          
                                                                   ON A.doctor_no = B.doctor_no                                              
                                                            LEFT JOIN mast.div C                                                             
                                                                   ON A.div_no = C.div_no                                                    
                                                     WHERE  1 = 1                                                                            
                                                            AND To_date(To_number(A.clinic_date) + 19110000, 'yyyymmdd') = Trunc(sysdate)    
                                                            AND A.clinic_flag IN ( 'E' )                                                     
                                                            AND A.korder_flag NOT IN ( 'X' )                                                 
                                                            AND A.course_seq <> '1'                                                          
                                                            AND A.chronic_med_seq NOT IN ( '2', '3' )                                        
                                                            AND A.building_no = 'B'                                                         
                                                     GROUP  BY A.building_no,                                                                
                                                               A.clinic_date,                                                                
                                                               C.div_full_name,                                                              
                                                               B.doctor_name   
                                                ";
                oddap = new OleDbDataAdapter(odcmm);
                oddap.Fill(dt);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            finally
            {
                if (oddap != null) { oddap.Dispose(); }
                if (odcmm != null) { odcmm.Dispose(); }
                if (odcnn != null) { odcnn.Dispose(); }
            }
            try
            {
                string cl0 = "";
                string cl1 = "";
                string cl2 = "小計";
                int cl3 = 0;
                int cl4 = 0;
                int cl5 = 0;
                foreach (DataRow dr in dt.Rows)
                {
                    cl3 += Convert.ToInt32(dr["大夜"]);
                    cl4 += Convert.ToInt32(dr["早班"]);
                    cl5 += Convert.ToInt32(dr["小夜"]);
                }
                DataRow row = dt.NewRow();
                row["看診日期"] = cl0;
                row["科別全名"] = cl1;
                row["醫師姓名"] = cl2;
                row["大夜"] = cl3;
                row["早班"] = cl4;
                row["小夜"] = cl5;

                dt.Rows.Add(row);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return (dt);
        }

        public DataTable Emtodayfirst()
        {
            Database db = new Database();
            OleDbConnection odcnn = db.GetOleDbConnection("ORACLE_DB_HO");
            OleDbCommand odcmm = new OleDbCommand();
            OleDbDataAdapter oddap = null;
            DataTable dt = new DataTable();

            try
            {
                odcnn.Open();
                odcmm.Connection = odcnn;
                odcmm.CommandType = CommandType.Text;
                odcmm.CommandText =
                                    @"
                                          SELECT A.clinic_date     AS ""看診日期"",                                                                      
                                                 Count(A.chart_no) AS ""初診人數""                                                                       
                                          FROM   opd.ptopd A                                                                                             
                                                LEFT JOIN mast.doctor B                                                                                  
                                                        ON A.doctor_no = B.doctor_no                                                                     
                                          WHERE  1 = 1                                                                                                   
                                                 AND A.fv_rv_flag = '1'                                                                                  
                                                 AND A.clinic_flag IN ( 'E' )                                                                            
                                                 AND A.merge_flag NOT IN( 'D', 'd' )                                                                     
                                                 AND A.reg_clerk NOT IN( 'KOPDMISC', 'KOPDCHR', 'KOPDREH', 'KOPDUREG' )                                  
                                                 AND A.chronic_med_seq NOT IN( '2', '3' )                                                                
                                                 AND A.pt_type NOT IN( '16', '18', '29', '36' )                                                          
                                                 AND To_date(( To_number(Replace(A.clinic_date, '.', ''))  + 19110000 ), 'yyyy-mm-dd') = Trunc(sysdate)  
                                                 --上行為晨恩提供的日期                                                                                   
                                                 --AND TO_DATE(TO_CHAR(TO_number(a.clinic_date)+19110000),'YYYY/MM/DD') = Trunc(sysdate)                 
                                                 --上行為我的日期                                                                                         
                                                 AND A.nh_case_type <> 'A3'                                                                              
                                                 AND A.nh_clinic_seq NOT IN( '31', '36', '91', '95', '85' )                                              
                                                 AND A.building_no = 'B'                                                                                
                                          GROUP  BY A.clinic_date 
                                    ";


                oddap = new OleDbDataAdapter(odcmm);
                oddap.Fill(dt);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            finally
            {
                if (oddap != null) { oddap.Dispose(); }
                if (odcmm != null) { odcmm.Dispose(); }
                if (odcnn != null) { odcnn.Dispose(); }
            }
            return (dt);
        }


        public DataTable Emtodaytransfer()
        {
            Database db = new Database();
            OleDbConnection odcnn = db.GetOleDbConnection("ORACLE_DB_HO");
            OleDbCommand odcmm = new OleDbCommand();
            OleDbDataAdapter oddap = null;
            DataTable dt = new DataTable();

            try
            {
                odcnn.Open();
                odcmm.Connection = odcnn;
                odcmm.CommandType = CommandType.Text;
                odcmm.CommandText =
                                                @"
                                                      SELECT A.clinic_date     AS ""看診日期"",                                                                          
                                                             Count(A.chart_no) AS ""轉住院人數""                                                                         
                                                      FROM   opd.ptopd A,                                                                                                
                                                             opd.pter D                                                                                                  
                                                      WHERE  1 = 1                                                                                                       
                                                             AND A.chart_no = D.chart_no                                                                                 
                                                             AND A.korder_flag NOT IN( 'X', ' ' )                                                                        
                                                             AND A.pt_type IN(SELECT pt_type                                                                             
                                                                              FROM   mast.pttype                                                                         
                                                                              WHERE  insurance_type = '2'                                                                
                                                                                     AND supple_flag <> 'Y')                                                             
                                                             AND A.nh_apply_flag IN( 'Y', 'U' )                                                                          
                                                             AND A.merge_flag NOT IN( 'D', 'd' )                                                                         
                                                             AND A.clinic_flag IN( 'E' )                                                                                 
                                                             AND D.er_acnt_close_flag IN( 'T0', 'T1', 'T2' )                                                             
                                                             AND A.reg_clerk NOT IN( 'KOPDREH' )                                                                         
                                                             AND A.chronic_med_seq NOT IN( '2', '3' )                                                                    
                                                             AND A.discnt_type NOT IN( '18', '29' )                                                                      
                                                             AND To_date(( To_number(Replace(A.clinic_date, '.', '')) + 19110000 ), 'yyyy-mm-dd') = Trunc(sysdate)       
                                                             --上行為晨恩提供的日期                                                                                       
                                                             --AND TO_DATE(TO_CHAR(TO_number(a.clinic_date)+19110000),'YYYY/MM/DD') = Trunc(sysdate)                     
                                                             --上行為我的日期                                                                                             
                                                             AND A.nh_case_type <> 'A3'                                                                                  
                                                             AND A.building_no = 'B'                                                                                    
                                                      GROUP  BY A.clinic_date
                                                ";

                oddap = new OleDbDataAdapter(odcmm);
                oddap.Fill(dt);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            finally
            {
                if (oddap != null) { oddap.Dispose(); }
                if (odcmm != null) { odcmm.Dispose(); }
                if (odcnn != null) { odcnn.Dispose(); }
            }
            return (dt);
        }

        public DataTable Emtotal()
        {
            Database db = new Database();
            OleDbConnection odcnn = db.GetOleDbConnection("ORACLE_DB_HO");
            OleDbCommand odcmm = new OleDbCommand();
            OleDbDataAdapter oddap = null;
            DataTable dt = new DataTable();

            try
            {
                odcnn.Open();
                odcmm.Connection = odcnn;
                odcmm.CommandType = CommandType.Text;
                odcmm.CommandText =
                                    @"
                                       SELECT CASE                                                                           
                                                WHEN building_no = 'A' THEN '員榮'                                           
                                                WHEN building_no = 'B' THEN '員生'                                           
                                              END                AS ""院區"",                                                
                                              A.clinic_date      AS ""看診日期"",                                            
                                              Count(building_no) AS ""小計""                                                 
                                       FROM   opd.ptopd A                                                                    
                                              LEFT JOIN mast.doctor B                                                        
                                                     ON A.doctor_no = B.doctor_no                                            
                                              LEFT JOIN mast.div C                                                           
                                                     ON A.div_no = C.div_no                                                  
                                       WHERE  1 = 1                                                                          
                                              AND To_date(To_number(A.clinic_date) + 19110000, 'yyyymmdd') = Trunc(sysdate)  
                                              AND A.clinic_flag IN ( 'E' )                                                   
                                              AND A.korder_flag NOT IN ( 'X' )                                               
                                              AND A.course_seq <> '1'                                                        
                                              AND A.chronic_med_seq NOT IN ( '2', '3' )                                      
                                              AND A.BUILDING_NO = 'B'                                                       
                                       GROUP  BY building_no,                                                                
                                                 clinic_date  
                                    ";
                oddap = new OleDbDataAdapter(odcmm);
                oddap.Fill(dt);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            finally
            {
                if (oddap != null) { oddap.Dispose(); }
                if (odcmm != null) { odcmm.Dispose(); }
                if (odcnn != null) { odcnn.Dispose(); }
            }
            return (dt);
        }



        public DataTable HotodayinO()
        {
            Database db = new Database();
            OleDbConnection odcnn = db.GetOleDbConnection("ORACLE_DB_HO");
            OleDbCommand odcmm = new OleDbCommand();
            OleDbDataAdapter oddap = null;
            DataTable dt = new DataTable();

            try
            {
                odcnn.Open();
                odcmm.Connection = odcnn;
                odcmm.CommandType = CommandType.Text;
                odcmm.CommandText =
                                        @"

                                                SELECT CASE ""院區""
                                                         WHEN 'A' THEN '員榮'
                                                         WHEN 'B' THEN '員生'
                                                       END                     AS ""院區"",
                                                       ""科別名稱"",
                                                       ""診間號"",
                                                       ""醫師名稱"",
                                                       ""早診"",
                                                       ""早診初診人次"",
                                                       ""午診"",
                                                       ""午診初診人次"",
                                                       ""晚診"",
                                                       ""晚診初診人次"",
                                                       COALESCE(""早診"", 0) + COALESCE(""午診"", 0)
                                                       + COALESCE(""晚診"", 0) AS ""合計""
                                                FROM   (SELECT ""院區"",
                                                               ""科別名稱"",
                                                               ""診間號"",
                                                               ""醫師名稱"",
                                                               Sum(""早診"")             AS ""早診"",
                                                               Sum(""早診初診人次"") AS ""早診初診人次"",
                                                               Sum(""午診"")             AS ""午診"",
                                                               Sum(""午診初診人次"") AS ""午診初診人次"",
                                                               Sum(""晚診"")             AS ""晚診"",
                                                               Sum(""晚診初診人次"") AS ""晚診初診人次""
                                                        FROM   (SELECT ""院區"",
                                                                       ""科別名稱"",
                                                                       ""診間號"",
                                                                       ""醫師名稱"",
                                                                       CASE ""午別""
                                                                         WHEN '1' THEN ""已看診人數""
                                                                         ELSE 0
                                                                       END AS ""早診"",
                                                                       CASE ""午別""
                                                                         WHEN '1' THEN ""已看初診人數""
                                                                         ELSE 0
                                                                       END AS ""早診初診人次"",
                                                                       CASE ""午別""
                                                                         WHEN '2' THEN ""已看診人數""
                                                                         ELSE 0
                                                                       END AS ""午診"",
                                                                       CASE ""午別""
                                                                         WHEN '2' THEN ""已看初診人數""
                                                                         ELSE 0
                                                                       END AS ""午診初診人次"",
                                                                       CASE ""午別""
                                                                         WHEN '3' THEN ""已看診人數""
                                                                         ELSE 0
                                                                       END AS ""晚診"",
                                                                       CASE ""午別""
                                                                         WHEN '3' THEN ""已看初診人數""
                                                                         ELSE 0
                                                                       END AS ""晚診初診人次"",
                                                                       ""診別""
                                                                FROM   (
                                                                       --門診無中醫--
                                                                       SELECT a.building_no                       AS ""院區"",
                                                                              a.clinic_date                       AS ""就醫日期"",
                                                                              b.week                              AS ""星期"",
                                                                              a.clinic_apn                        AS ""午別"",
                                                                              a.clinic_no                         AS ""診間號"",
                                                                              b.div_no                            AS ""科別代碼"",
                                                                              b.clinic_name                       AS ""科別名稱"",
                                                                              d.doctor_name                       AS ""醫師名稱"",
                                                                              a.clinic_flag                       AS ""診別"",
                                                                              Count(*)                            AS ""已看診人數"",
                                                                              COALESCE(e.""已看初診人數"", 0)       AS ""已看初診人數""
                                                                       FROM   opd.ptopd a
                                                                              LEFT JOIN opd.dclin b
                                                                                     ON a.clinic_date = b.clinic_date
                                                                                        AND a.clinic_apn = b.clinic_apn
                                                                                        AND a.clinic_no = b.clinic_no
                                                                                        AND a.doctor_no = b.doctor_no
                                                                              LEFT JOIN mast.div c
                                                                                     ON a.div_no = c.div_no
                                                                              LEFT JOIN mast.doctor d
                                                                                     ON a.doctor_no = d.doctor_no
                                                                              LEFT JOIN (SELECT
                                                                                                a.building_no AS ""院區"",
                                                                                                a.clinic_date AS ""就醫日期"",
                                                                                                b.week        AS ""星期"",
                                                                                                a.clinic_apn  AS ""午別"",
                                                                                                a.clinic_no   AS ""診間號"",
                                                                                                b.div_no      AS ""科別代碼"",
                                                                                                b.clinic_name AS ""科別名稱"",
                                                                                                d.doctor_name AS ""醫師名稱"",
                                                                                                a.clinic_flag AS ""診別"",
                                                                                                Count(*)      AS ""已看初診人數""
                                                                                         FROM   opd.ptopd a
                                                                                                LEFT JOIN opd.dclin b
                                                                                                       ON a.clinic_date =
                                                                                                          b.clinic_date
                                                                                                          AND a.clinic_apn =
                                                                                                              b.clinic_apn
                                                                                                          AND a.clinic_no =
                                                                                                              b.clinic_no
                                                                                                          AND a.doctor_no =
                                                                                                              b.doctor_no
                                                                                                LEFT JOIN mast.div c
                                                                                                       ON a.div_no = c.div_no
                                                                                                LEFT JOIN mast.doctor d
                                                                                                       ON a.doctor_no =
                                                                                                          d.doctor_no
                                                                                         WHERE  To_date(To_char(To_number(a.clinic_date) + 19110000), 'yyyy/MM/dd') = Trunc(sysdate)
                                                                                                AND a.clinic_flag IN ( 'O' )
                                                                                                --AND a.clinic_flag IN ( 'O', 'V' )
                                                                                                AND a.korder_flag IN ('Y', 'P', 'R' )
                                                                                                AND a.reg_clerk NOT IN ('KOPDMISC', 'KOPDCHR','KOPDREH','KOPDUREG')
                                                                                                AND a.fv_rv_flag = '1'
                                                                                                AND a.building_no <> 'B'
                                                                                         GROUP  BY a.building_no,
                                                                                                   a.clinic_date,
                                                                                                   b.week,
                                                                                                   a.clinic_apn,
                                                                                                   a.clinic_no,
                                                                                                   b.div_no,
                                                                                                   b.clinic_name,
                                                                                                   d.doctor_name,
                                                                                                   a.clinic_flag
                                                                                         ORDER  BY a.clinic_flag,
                                                                                                   a.clinic_apn,
                                                                                                   a.clinic_no)e
                                                                                     ON a.clinic_date = e.""就醫日期""
                                                                                        AND b.week = e.""星期""
                                                                                        AND a.clinic_apn = e.""午別""
                                                                                        AND a.clinic_no = e.""診間號""
                                                                                        AND b.div_no = e.""科別代碼""
                                                                                        AND b.clinic_name = e.""科別名稱""
                                                                                        AND d.doctor_name = e.""醫師名稱""
                                                                                        AND a.clinic_flag = e.""診別""
                                                                       WHERE  To_date(To_char(To_number(a.clinic_date) + 19110000), 'yyyy/MM/dd') = Trunc(sysdate)
                                                                              AND a.clinic_no != ' '
                                                                              AND A.div_no != '60'
                                                                              AND a.clinic_flag IN ( 'O' )
                                                                              --AND a.clinic_flag IN ( 'O', 'V' )
                                                                              AND a.korder_flag IN ( 'Y', 'P', 'R' )
                                                                              AND a.reg_clerk NOT IN ('KOPDMISC', 'KOPDCHR','KOPDREH','KOPDUREG')
                                                                              AND a.building_no = 'B'
                                                                       GROUP  BY a.building_no,
                                                                                 a.clinic_date,
                                                                                 b.week,
                                                                                 a.clinic_apn,
                                                                                 a.clinic_no,
                                                                                 b.div_no,
                                                                                 b.clinic_name,
                                                                                 a.clinic_no,
                                                                                 d.doctor_name,
                                                                                 a.clinic_flag,
                                                                                 e.""已看初診人數""
                                                                        UNION ALL
                                                                        --門診中醫--         
                                                                        SELECT a.building_no                       AS ""院區"",
                                                                               a.clinic_date                       AS ""就醫日期"",
                                                                               b.week                              AS ""星期"",
                                                                               a.clinic_apn                        AS ""午別"",
                                                                               a.clinic_no                         AS ""診間號"",
                                                                               b.div_no                            AS ""科別代碼"",
                                                                               b.clinic_name                       AS ""科別名稱"",
                                                                               d.doctor_name                       AS ""醫師名稱"",
                                                                               a.clinic_flag                       AS ""診別"",
                                                                               Count(*)                            AS ""已看診人數"",
                                                                               COALESCE(e.""已看初診人數"", 0)       AS ""已看初診人數""
                                                                        FROM   opd.ptopd a
                                                                        LEFT JOIN opd.dclin b
                                                                           ON a.clinic_date = b.clinic_date
                                                                            AND a.clinic_apn = b.clinic_apn
                                                                            AND a.clinic_no = b.clinic_no
                                                                            AND a.doctor_no = b.doctor_no
                                                                        LEFT JOIN mast.div c
                                                                           ON a.div_no = c.div_no
                                                                        LEFT JOIN mast.doctor d
                                                                           ON a.doctor_no = d.doctor_no
                                                                        LEFT JOIN (SELECT a.building_no AS ""院區"",
                                                                                a.clinic_date  AS ""就醫日期"",
                                                                                b.week         AS ""星期"",
                                                                                a.clinic_apn   AS ""午別"",
                                                                                a.clinic_no    AS ""診間號"",
                                                                                b.div_no       AS ""科別代碼"",
                                                                                b.clinic_name  AS ""科別名稱"",
                                                                                d.doctor_name  AS ""醫師名稱"",
                                                                                a.clinic_flag  AS ""診別"",
                                                                                Count(*)       AS ""已看初診人數""
                                                                             FROM   opd.ptopd a
                                                                                LEFT JOIN opd.dclin b
                                                                                   ON a.clinic_date = b.clinic_date
                                                                                      AND a.clinic_apn = b.clinic_apn
                                                                                      AND a.clinic_no = b.clinic_no
                                                                                      AND a.doctor_no = b.doctor_no
                                                                                LEFT JOIN mast.div c
                                                                                   ON a.div_no = c.div_no
                                                                                LEFT JOIN mast.doctor d
                                                                                   ON a.doctor_no = d.doctor_no
                                                                             WHERE  To_date(To_char(To_number(a.clinic_date) + 19110000), 'yyyy/MM/dd') = Trunc(sysdate)
                                                                                --AND a.course_seq IN ( '0', '1' )
                                                                                AND A.div_no = '60'
                                                                                AND a.clinic_flag IN ( 'O' )
                                                                                --AND a.clinic_flag IN ( 'O', 'V' )
                                                                                AND a.korder_flag IN ( 'Y', 'P', 'R' )
                                                                                AND a.reg_clerk NOT IN ('KOPDMISC', 'KOPDCHR','KOPDREH','KOPDUREG')
                                                                                AND a.fv_rv_flag = '1'
                                                                                AND a.building_no = 'B'
                                                                             GROUP  BY a.building_no,
                                                                                 a.clinic_date,
                                                                                 b.week,
                                                                                 a.clinic_apn,
                                                                                 a.clinic_no,
                                                                                 b.div_no,
                                                                                 b.clinic_name,
                                                                                 d.doctor_name,
                                                                                 a.clinic_flag
                                                                             ORDER  BY a.clinic_flag,
                                                                                 a.clinic_apn,
                                                                                 a.clinic_no)e
                                                                     ON a.clinic_date = e.""就醫日期""
                                                                        AND b.week = e.""星期""
                                                                        AND a.clinic_apn = e.""午別""
                                                                        AND a.clinic_no = e.""診間號""
                                                                        AND b.div_no = e.""科別代碼""
                                                                        AND b.clinic_name = e.""科別名稱""
                                                                        AND d.doctor_name = e.""醫師名稱""
                                                                        AND a.clinic_flag = e.""診別""
                                                              WHERE  To_date(To_char(To_number(a.clinic_date) + 19110000), 'yyyy/MM/dd') = Trunc(sysdate)
                                                              --AND a.course_seq IN ( '0', '1' )
                                                              AND A.div_no = '60'
                                                              AND a.clinic_flag IN ( 'O' )
                                                              --AND a.clinic_flag IN ( 'O', 'V' )
                                                              AND a.korder_flag IN ( 'Y', 'P', 'R' )
                                                              AND a.reg_clerk NOT IN ('KOPDMISC', 'KOPDCHR','KOPDREH','KOPDUREG')
                                                              AND a.building_no = 'B'
                                                              GROUP  BY a.building_no,
                                                                 a.clinic_date,
                                                                 b.week,
                                                                 a.clinic_apn,
                                                                 a.clinic_no,
                                                                 b.div_no,
                                                                 b.clinic_name,
                                                                 a.clinic_no,
                                                                 d.doctor_name,
                                                                 a.clinic_flag,
                                                                 e.""已看初診人數""))
                                                               GROUP  BY ""院區"",
                                                                         ""科別名稱"",
                                                                         ""診間號"",
                                                                         ""醫師名稱"",
                                                                         ""診別""
                                                               ORDER  BY ""院區"",
                                                                         ""科別名稱"",
                                                                         ""診間號"")

                                        ";
                oddap = new OleDbDataAdapter(odcmm);
                oddap.Fill(dt);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            finally
            {
                if (oddap != null) { oddap.Dispose(); }
                if (odcmm != null) { odcmm.Dispose(); }
                if (odcnn != null) { odcnn.Dispose(); }
            }

            try
            {

                string c0 = "";
                string c1 = "";
                string c2 = "合計";
                int c3 = 0;
                int c4 = 0;
                int c5 = 0;
                int c6 = 0;
                int c7 = 0;
                int c8 = 0;
                int c9 = 0;
                foreach (DataRow dr in dt.Rows)
                {
                    c3 += Convert.ToInt32(dr["早診"]);
                    c4 += Convert.ToInt32(dr["早診初診人次"]);
                    c5 += Convert.ToInt32(dr["午診"]);
                    c6 += Convert.ToInt32(dr["午診初診人次"]);
                    c7 += Convert.ToInt32(dr["晚診"]);
                    c8 += Convert.ToInt32(dr["晚診初診人次"]);
                    c9 += Convert.ToInt32(dr["合計"]);
                }
                DataRow newrow1 = dt.NewRow();
                newrow1["科別名稱"] = c0;
                newrow1["診間號"] = c1;
                newrow1["醫師名稱"] = c2;
                newrow1["早診"] = c3;
                newrow1["早診初診人次"] = c4;
                newrow1["午診"] = c5;
                newrow1["午診初診人次"] = c6;
                newrow1["晚診"] = c7;
                newrow1["晚診初診人次"] = c8;
                newrow1["合計"] = c9;
                dt.Rows.Add(newrow1);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            finally
            {
                ;
            }
            return (dt);
        }




        public DataTable HotodayinV()
        {
            Database db = new Database();
            OleDbConnection odcnn = db.GetOleDbConnection("ORACLE_DB_HO");
            OleDbCommand odcmm = new OleDbCommand();
            OleDbDataAdapter oddap = null;
            DataTable dt = new DataTable();

            try
            {
                odcnn.Open();
                odcmm.Connection = odcnn;
                odcmm.CommandType = CommandType.Text;
                odcmm.CommandText =

                    @"

                                                SELECT CASE ""院區""
                                                         WHEN 'A' THEN '員榮'
                                                         WHEN 'B' THEN '員生'
                                                       END                     AS ""院區"",
                                                       ""科別名稱"",
                                                       ""診間號"",
                                                       ""醫師名稱"",
                                                       ""早診"",
                                                       ""早診初診人次"",
                                                       ""午診"",
                                                       ""午診初診人次"",
                                                       ""晚診"",
                                                       ""晚診初診人次"",
                                                       COALESCE(""早診"", 0) + COALESCE(""午診"", 0)
                                                       + COALESCE(""晚診"", 0) AS ""合計""
                                                FROM   (SELECT ""院區"",
                                                               ""科別名稱"",
                                                               ""診間號"",
                                                               ""醫師名稱"",
                                                               Sum(""早診"")             AS ""早診"",
                                                               Sum(""早診初診人次"") AS ""早診初診人次"",
                                                               Sum(""午診"")             AS ""午診"",
                                                               Sum(""午診初診人次"") AS ""午診初診人次"",
                                                               Sum(""晚診"")             AS ""晚診"",
                                                               Sum(""晚診初診人次"") AS ""晚診初診人次""
                                                        FROM   (SELECT ""院區"",
                                                                       ""科別名稱"",
                                                                       ""診間號"",
                                                                       ""醫師名稱"",
                                                                       CASE ""午別""
                                                                         WHEN '1' THEN ""已看診人數""
                                                                         ELSE 0
                                                                       END AS ""早診"",
                                                                       CASE ""午別""
                                                                         WHEN '1' THEN ""已看初診人數""
                                                                         ELSE 0
                                                                       END AS ""早診初診人次"",
                                                                       CASE ""午別""
                                                                         WHEN '2' THEN ""已看診人數""
                                                                         ELSE 0
                                                                       END AS ""午診"",
                                                                       CASE ""午別""
                                                                         WHEN '2' THEN ""已看初診人數""
                                                                         ELSE 0
                                                                       END AS ""午診初診人次"",
                                                                       CASE ""午別""
                                                                         WHEN '3' THEN ""已看診人數""
                                                                         ELSE 0
                                                                       END AS ""晚診"",
                                                                       CASE ""午別""
                                                                         WHEN '3' THEN ""已看初診人數""
                                                                         ELSE 0
                                                                       END AS ""晚診初診人次"",
                                                                       ""診別""
                                                                FROM   (
                                                                       --門診無中醫--
                                                                       SELECT a.building_no                       AS ""院區"",
                                                                              a.clinic_date                       AS ""就醫日期"",
                                                                              b.week                              AS ""星期"",
                                                                              a.clinic_apn                        AS ""午別"",
                                                                              a.clinic_no                         AS ""診間號"",
                                                                              b.div_no                            AS ""科別代碼"",
                                                                              b.clinic_name                       AS ""科別名稱"",
                                                                              d.doctor_name                       AS ""醫師名稱"",
                                                                              a.clinic_flag                       AS ""診別"",
                                                                              Count(*)                            AS ""已看診人數"",
                                                                              COALESCE(e.""已看初診人數"", 0)       AS ""已看初診人數""
                                                                       FROM   opd.ptopd a
                                                                              LEFT JOIN opd.dclin b
                                                                                     ON a.clinic_date = b.clinic_date
                                                                                        AND a.clinic_apn = b.clinic_apn
                                                                                        AND a.clinic_no = b.clinic_no
                                                                                        AND a.doctor_no = b.doctor_no
                                                                              LEFT JOIN mast.div c
                                                                                     ON a.div_no = c.div_no
                                                                              LEFT JOIN mast.doctor d
                                                                                     ON a.doctor_no = d.doctor_no
                                                                              LEFT JOIN (SELECT
                                                                                                a.building_no AS ""院區"",
                                                                                                a.clinic_date AS ""就醫日期"",
                                                                                                b.week        AS ""星期"",
                                                                                                a.clinic_apn  AS ""午別"",
                                                                                                a.clinic_no   AS ""診間號"",
                                                                                                b.div_no      AS ""科別代碼"",
                                                                                                b.clinic_name AS ""科別名稱"",
                                                                                                d.doctor_name AS ""醫師名稱"",
                                                                                                a.clinic_flag AS ""診別"",
                                                                                                Count(*)      AS ""已看初診人數""
                                                                                         FROM   opd.ptopd a
                                                                                                LEFT JOIN opd.dclin b
                                                                                                       ON a.clinic_date =
                                                                                                          b.clinic_date
                                                                                                          AND a.clinic_apn =
                                                                                                              b.clinic_apn
                                                                                                          AND a.clinic_no =
                                                                                                              b.clinic_no
                                                                                                          AND a.doctor_no =
                                                                                                              b.doctor_no
                                                                                                LEFT JOIN mast.div c
                                                                                                       ON a.div_no = c.div_no
                                                                                                LEFT JOIN mast.doctor d
                                                                                                       ON a.doctor_no =
                                                                                                          d.doctor_no
                                                                                         WHERE  To_date(To_char(To_number(a.clinic_date) + 19110000), 'yyyy/MM/dd') = Trunc(sysdate)
                                                                                                AND a.clinic_flag IN ( 'V' )
                                                                                                --AND a.clinic_flag IN ( 'O', 'V' )
                                                                                                AND a.korder_flag IN ('Y', 'P', 'R' )
                                                                                                AND a.reg_clerk NOT IN ('KOPDMISC', 'KOPDCHR','KOPDREH','KOPDUREG')
                                                                                                AND a.fv_rv_flag = '1'
                                                                                                AND a.building_no = 'B'
                                                                                         GROUP  BY a.building_no,
                                                                                                   a.clinic_date,
                                                                                                   b.week,
                                                                                                   a.clinic_apn,
                                                                                                   a.clinic_no,
                                                                                                   b.div_no,
                                                                                                   b.clinic_name,
                                                                                                   d.doctor_name,
                                                                                                   a.clinic_flag
                                                                                         ORDER  BY a.clinic_flag,
                                                                                                   a.clinic_apn,
                                                                                                   a.clinic_no)e
                                                                                     ON a.clinic_date = e.""就醫日期""
                                                                                        AND b.week = e.""星期""
                                                                                        AND a.clinic_apn = e.""午別""
                                                                                        AND a.clinic_no = e.""診間號""
                                                                                        AND b.div_no = e.""科別代碼""
                                                                                        AND b.clinic_name = e.""科別名稱""
                                                                                        AND d.doctor_name = e.""醫師名稱""
                                                                                        AND a.clinic_flag = e.""診別""
                                                                       WHERE  To_date(To_char(To_number(a.clinic_date) + 19110000), 'yyyy/MM/dd') = Trunc(sysdate)
                                                                              AND a.clinic_no != ' '
                                                                              AND A.div_no != '60'
                                                                              AND a.clinic_flag IN ( 'V' )
                                                                              --AND a.clinic_flag IN ( 'O', 'V' )
                                                                              AND a.korder_flag IN ( 'Y', 'P', 'R' )
                                                                              AND a.reg_clerk NOT IN ('KOPDMISC', 'KOPDCHR','KOPDREH','KOPDUREG')
                                                                              AND a.building_no = 'B'
                                                                       GROUP  BY a.building_no,
                                                                                 a.clinic_date,
                                                                                 b.week,
                                                                                 a.clinic_apn,
                                                                                 a.clinic_no,
                                                                                 b.div_no,
                                                                                 b.clinic_name,
                                                                                 a.clinic_no,
                                                                                 d.doctor_name,
                                                                                 a.clinic_flag,
                                                                                 e.""已看初診人數""
                                                                        UNION ALL
                                                                        --門診中醫--         
                                                                        SELECT a.building_no                       AS ""院區"",
                                                                               a.clinic_date                       AS ""就醫日期"",
                                                                               b.week                              AS ""星期"",
                                                                               a.clinic_apn                        AS ""午別"",
                                                                               a.clinic_no                         AS ""診間號"",
                                                                               b.div_no                            AS ""科別代碼"",
                                                                               b.clinic_name                       AS ""科別名稱"",
                                                                               d.doctor_name                       AS ""醫師名稱"",
                                                                               a.clinic_flag                       AS ""診別"",
                                                                               Count(*)                            AS ""已看診人數"",
                                                                               COALESCE(e.""已看初診人數"", 0)       AS ""已看初診人數""
                                                                        FROM   opd.ptopd a
                                                                        LEFT JOIN opd.dclin b
                                                                           ON a.clinic_date = b.clinic_date
                                                                            AND a.clinic_apn = b.clinic_apn
                                                                            AND a.clinic_no = b.clinic_no
                                                                            AND a.doctor_no = b.doctor_no
                                                                        LEFT JOIN mast.div c
                                                                           ON a.div_no = c.div_no
                                                                        LEFT JOIN mast.doctor d
                                                                           ON a.doctor_no = d.doctor_no
                                                                        LEFT JOIN (SELECT a.building_no AS ""院區"",
                                                                                a.clinic_date  AS ""就醫日期"",
                                                                                b.week         AS ""星期"",
                                                                                a.clinic_apn   AS ""午別"",
                                                                                a.clinic_no    AS ""診間號"",
                                                                                b.div_no       AS ""科別代碼"",
                                                                                b.clinic_name  AS ""科別名稱"",
                                                                                d.doctor_name  AS ""醫師名稱"",
                                                                                a.clinic_flag  AS ""診別"",
                                                                                Count(*)       AS ""已看初診人數""
                                                                             FROM   opd.ptopd a
                                                                                LEFT JOIN opd.dclin b
                                                                                   ON a.clinic_date = b.clinic_date
                                                                                      AND a.clinic_apn = b.clinic_apn
                                                                                      AND a.clinic_no = b.clinic_no
                                                                                      AND a.doctor_no = b.doctor_no
                                                                                LEFT JOIN mast.div c
                                                                                   ON a.div_no = c.div_no
                                                                                LEFT JOIN mast.doctor d
                                                                                   ON a.doctor_no = d.doctor_no
                                                                             WHERE  To_date(To_char(To_number(a.clinic_date) + 19110000), 'yyyy/MM/dd') = Trunc(sysdate)
                                                                                --AND a.course_seq IN ( '0', '1' )
                                                                                AND A.div_no = '60'
                                                                                AND a.clinic_flag IN ( 'V' )
                                                                                --AND a.clinic_flag IN ( 'O', 'V' )
                                                                                AND a.korder_flag IN ( 'Y', 'P', 'R' )
                                                                                AND a.reg_clerk NOT IN ('KOPDMISC', 'KOPDCHR','KOPDREH','KOPDUREG')
                                                                                AND a.fv_rv_flag = '1'
                                                                                AND a.building_no = 'B'
                                                                             GROUP  BY a.building_no,
                                                                                 a.clinic_date,
                                                                                 b.week,
                                                                                 a.clinic_apn,
                                                                                 a.clinic_no,
                                                                                 b.div_no,
                                                                                 b.clinic_name,
                                                                                 d.doctor_name,
                                                                                 a.clinic_flag
                                                                             ORDER  BY a.clinic_flag,
                                                                                 a.clinic_apn,
                                                                                 a.clinic_no)e
                                                                     ON a.clinic_date = e.""就醫日期""
                                                                        AND b.week = e.""星期""
                                                                        AND a.clinic_apn = e.""午別""
                                                                        AND a.clinic_no = e.""診間號""
                                                                        AND b.div_no = e.""科別代碼""
                                                                        AND b.clinic_name = e.""科別名稱""
                                                                        AND d.doctor_name = e.""醫師名稱""
                                                                        AND a.clinic_flag = e.""診別""
                                                              WHERE  To_date(To_char(To_number(a.clinic_date) + 19110000), 'yyyy/MM/dd') = Trunc(sysdate)
                                                              --AND a.course_seq IN ( '0', '1' )
                                                              AND A.div_no = '60'
                                                              AND a.clinic_flag IN ( 'V' )
                                                              --AND a.clinic_flag IN ( 'O', 'V' )
                                                              AND a.korder_flag IN ( 'Y', 'P', 'R' )
                                                              AND a.reg_clerk NOT IN ('KOPDMISC', 'KOPDCHR','KOPDREH','KOPDUREG')
                                                              AND a.building_no = 'B'
                                                              GROUP  BY a.building_no,
                                                                 a.clinic_date,
                                                                 b.week,
                                                                 a.clinic_apn,
                                                                 a.clinic_no,
                                                                 b.div_no,
                                                                 b.clinic_name,
                                                                 a.clinic_no,
                                                                 d.doctor_name,
                                                                 a.clinic_flag,
                                                                 e.""已看初診人數""))
                                                               GROUP  BY ""院區"",
                                                                         ""科別名稱"",
                                                                         ""診間號"",
                                                                         ""醫師名稱"",
                                                                         ""診別""
                                                               ORDER  BY ""院區"",
                                                                         ""科別名稱"",
                                                                         ""診間號"")

                                        ";
                oddap = new OleDbDataAdapter(odcmm);
                oddap.Fill(dt);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            finally
            {
                if (oddap != null) { oddap.Dispose(); }
                if (odcmm != null) { odcmm.Dispose(); }
                if (odcnn != null) { odcnn.Dispose(); }
            }

            try
            {

                string c0 = "";
                string c1 = "";
                string c2 = "合計";
                int c3 = 0;
                int c4 = 0;
                int c5 = 0;
                int c6 = 0;
                int c7 = 0;
                int c8 = 0;
                int c9 = 0;
                foreach (DataRow dr in dt.Rows)
                {
                    c3 += Convert.ToInt32(dr["早診"]);
                    c4 += Convert.ToInt32(dr["早診初診人次"]);
                    c5 += Convert.ToInt32(dr["午診"]);
                    c6 += Convert.ToInt32(dr["午診初診人次"]);
                    c7 += Convert.ToInt32(dr["晚診"]);
                    c8 += Convert.ToInt32(dr["晚診初診人次"]);
                    c9 += Convert.ToInt32(dr["合計"]);
                }
                DataRow newrow1 = dt.NewRow();
                newrow1["科別名稱"] = c0;
                newrow1["診間號"] = c1;
                newrow1["醫師名稱"] = c2;
                newrow1["早診"] = c3;
                newrow1["早診初診人次"] = c4;
                newrow1["午診"] = c5;
                newrow1["午診初診人次"] = c6;
                newrow1["晚診"] = c7;
                newrow1["晚診初診人次"] = c8;
                newrow1["合計"] = c9;
                dt.Rows.Add(newrow1);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            finally
            {
                ;
            }
            return (dt);
        }



        public DataTable Icunow()
        {
            Database db = new Database();
            OleDbConnection odcnn = db.GetOleDbConnection("ORACLE_DB_HO");
            OleDbCommand odcmm = new OleDbCommand();
            OleDbDataAdapter oddap = null;
            DataTable dt = new DataTable();

            try
            {
                odcnn.Open();
                odcmm.Connection = odcnn;
                odcmm.CommandType = CommandType.Text;
                odcmm.CommandText =
                                    @"
                                    SELECT ""院區"",                                                   
                                           ""科別名稱"",                                               
                                           ""醫師名稱"",                                               
                                           Count(*)AS ""小計""                                         
                                    FROM   (SELECT CASE                                               
                                                     WHEN A.bed_no LIKE 'ICU%' THEN '員榮'            
                                                     WHEN A.bed_no LIKE 'BICU%' THEN '員生'           
                                                   END             AS ""院區"",                        
                                                   b.div_full_name AS ""科別名稱"",                    
                                                   c.doctor_name   AS ""醫師名稱""                     
                                            FROM   ipd.ptipd @hi a                                    
                                                   left join mast.div b                               
                                                          ON a.div_no = b.div_no                      
                                                   left join mast.doctor c                            
                                                          ON a.vs_no = c.doctor_no                    
                                            WHERE  1 = 1                                              
                                                   AND A.discharge_date = '0'                         
                                                   AND A.bed_no LIKE 'BICU%')                          
                                    GROUP  BY ""院區"",                                                
                                              ""科別名稱"",                                            
                                              ""醫師名稱""                                             
                                    ORDER  BY ""院區"" DESC                                            
                                    ";

                oddap = new OleDbDataAdapter(odcmm);
                oddap.Fill(dt);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (oddap != null) { oddap.Dispose(); }
                if (odcmm != null) { odcmm.Dispose(); }
                if (odcnn != null) { odcnn.Dispose(); }
            }
            //計算小計
            try
            {
                int total = 0;
                string subtotal = "小計";
                foreach (DataRow row in dt.Rows)
                {
                    total += Convert.ToInt32(row["小計"]);
                }
                DataRow dataRow = dt.NewRow();
                dataRow["醫師名稱"] = subtotal;
                dataRow["小計"] = total;
                dt.Rows.Add(dataRow);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                ;
            }
            return dt;
        }


        public DataTable Nursestation6()
        {
            Database db = new Database();
            OleDbConnection odcnn = db.GetOleDbConnection("ORACLE_DB_HO");
            OleDbCommand odcmm = new OleDbCommand();
            OleDbDataAdapter oddap = null;
            DataTable dt = new DataTable();

            string nursestation;
            int healthcare;
            int suite;
            int total;
            int actual;


            try
            {
                odcnn.Open();
                odcmm.Connection = odcnn;
                odcmm.CommandType = CommandType.Text;
                odcmm.CommandText =
                                         @"
                                            SELECT ""院區"",
                                                   '全護理站'            AS ""護理站"",
                                                   SUM(""實際可用床數"") AS ""實際可用床數"",                                                                
                                                   SUM(""健保"")             AS ""健保"",                                                                   
                                                   SUM(""套房"")             AS ""套房"",                                                                   
                                                   SUM(""合計"")             AS ""合計"",                                                                   
                                                   Concat(To_char(Round(SUM(""合計"") / SUM(""實際可用床數"") * 100, 2)), '%') AS ""佔床率""                 
                                            FROM   (SELECT CASE                                                                                           
                                                             WHEN Y.ns_code NOT LIKE 'B%' THEN '員榮'                                                     
                                                             WHEN Y.ns_code LIKE 'B%' THEN '員生'                                                         
                                                           END AS ""院區"",                                                                              
                                                           ""健保"",                                                                                       
                                                           ""套房"",                                                                                       
                                                           ""合計"",                                                                                       
                                                           ""實際可用床數""                                                                                 
                                                    FROM   (SELECT T.ns_code,                                                                             
                                                                   Count(Y.nh_paid_flag) AS ""健保"",                                                      
                                                                   Count(U.nh_paid_flag) AS ""套房"",                                                      
                                                                   Count(x.admit_no)     AS ""合計""                                                       
                                                            FROM   ipd.ptipd@hi x                                                                         
                                                                   left join ipd.bed@hi T                                                                 
                                                                          ON X.bed_no = T.bed_no                                                          
                                                                   left join(SELECT A.bed_no,                                                             
                                                                                    A.ns_code,                                                            
                                                                                    C.exclusive_ward_flag,                                                
                                                                                    B.grade_code,                                                         
                                                                                    B.statistic_grade,                                                    
                                                                                    B.description,                                                        
                                                                                    B.nh_paid_flag                                                        
                                                                             FROM   ipd.bed@hi A                                                          
                                                                                    left join ipd.bedgrade1@hi B                                          
                                                                                           ON A.grade_code = B.grade_code                                 
                                                                                    left join ipd.bedgrade2@hi C                                          
                                                                                           ON B.grade_code = C.grade_code                                 
                                                                             WHERE  A.effective_date = (SELECT Max(Z.effective_date)                      
                                                                                                        FROM ipd.bed@hi Z                                 
                                                                                                        WHERE A.bed_no = Z.bed_no)                        
                                                                                    AND B.nh_paid_flag = 'Y'                                              
                                                                             GROUP  BY A.bed_no,                                                          
                                                                                       c.exclusive_ward_flag,                                             
                                                                                       B.grade_code,                                                      
                                                                                       B.statistic_grade,                                                 
                                                                                       B.description,                                                     
                                                                                       B.nh_paid_flag,                                                    
                                                                                       A.ns_code                                                          
                                                                             ORDER  BY A.bed_no) y                                                        
                                                                          ON x.bed_no = y.bed_no                                                          
                                                                             AND x.exclusive_ward_flag = y.exclusive_ward_flag                            
                                                                   left join(SELECT A.bed_no,                                                             
                                                                                    A.ns_code,                                                            
                                                                                    c.exclusive_ward_flag,                                                
                                                                                    B.grade_code,                                                         
                                                                                    B.statistic_grade,                                                    
                                                                                    B.description,                                                        
                                                                                    B.nh_paid_flag                                                        
                                                                             --A.STATISTIC_FLAG  AS ""佔床率""                                             
                                                                             FROM   ipd.bed @hi A                                                         
                                                                                    left join ipd.bedgrade1 @hi B                                         
                                                                                           ON A.grade_code = B.grade_code                                 
                                                                                    left join ipd.bedgrade2 @hi C                                         
                                                                                           ON B.grade_code = C.grade_code                                 
                                                                             WHERE  A.effective_date = (SELECT Max(Z.effective_date)                      
                                                                                                        FROM ipd.bed @hi Z                                
                                                                                                        WHERE A.bed_no = Z.bed_no)                        
                                                                                    AND B.nh_paid_flag = 'N'                                              
                                                                             GROUP  BY A.bed_no,                                                          
                                                                                       c.exclusive_ward_flag,                                             
                                                                                       B.grade_code,                                                      
                                                                                       B.statistic_grade,                                                 
                                                                                       B.description,                                                     
                                                                                       B.nh_paid_flag,                                                    
                                                                                       A.ns_code                                                          
                                                                             ORDER  BY A.bed_no) U                                                        
                                                                          ON x.bed_no = U.bed_no                                                          
                                                                             AND x.exclusive_ward_flag = U.exclusive_ward_flag                            
                                                            WHERE  x.discharge_date = '0'                                                                 
                                                                   AND T.effective_date = (SELECT Max(S.effective_date)                                   
                                                                                           FROM ipd.bed @hi S                                             
                                                                                           WHERE  T.bed_no = S.bed_no)                                    
                                                                   AND x.status = '1'                                                                     
                                                            GROUP  BY T.ns_code,                                                                          
                                                                      Substr(x.bed_no, 1, 1)) Y                                                           
                                                           left join(SELECT a.ns_code,                                                                    
                                                                            a.bed_amt            AS ""衛福部登記開床數"",                                   
                                                                            a.nh_bed_amt         AS ""向健保署報備總床數"",                                 
                                                                            a.real_bed_amt       AS ""實際可用床數"",                                       
                                                                            a.real_empty_bed_amt AS ""實際可用空床數"",                                     
                                                                            a.nh_paid_bed_amt    AS ""向健保署報備的健保床數"",                             
                                                                            a.nh_diff_bed_amt    AS ""向健保署報備的差額床數""                              
                                                                     FROM   ipd.nsdivs@hi A                                                               
                                                                     WHERE  a.effective_date = (SELECT Max(Z.effective_date)                              
                                                                                                FROM ipd.nsdivs@hi z                                      
                                                                                                WHERE  a.ns_code = Z.ns_code)                             
                                                                     ORDER  BY a.ns_code)X                                                                
                                                                  ON Y.ns_code = X.ns_code                                                                
                                                                   WHERE Y.ns_code LIKE 'B%' )                                                        
                                            GROUP  BY ""院區""                                                                                           
                                            ORDER  BY ""院區"" DESC  
                                         ";

                oddap = new OleDbDataAdapter(odcmm);
                oddap.Fill(dt);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (oddap != null) { oddap.Dispose(); }
                if (odcmm != null) { odcmm.Dispose(); }
                if (odcnn != null) { odcnn.Dispose(); }
            }
            return dt;
        }

        public DataTable Hotodaysum()
        {

            Database db = new Database();
            OleDbConnection odcnn = db.GetOleDbConnection("ORACLE_DB_HO");
            OleDbCommand odcmm = new OleDbCommand();
            OleDbDataAdapter oddap = null;
            DataTable dt = new DataTable();
            try
            {
                odcnn.Open();
                odcmm.Connection = odcnn;
                odcmm.CommandType = CommandType.Text;
                odcmm.CommandText =
                 @"
                    SELECT CASE ""院區""                                                                                                                             
                             WHEN 'A' THEN '員榮'                                                                                                                    
                             WHEN 'B' THEN '員生'                                                                                                                    
                           END                 AS ""院區"",                                                                                                          
                           Sum(""已看診人數"")  AS ""小計""                                                                                                           
                    FROM   (                                                                                                                                         
                           SELECT a.building_no                   AS ""院區"",                                                                                       
                                  Count(*)                        AS ""已看診人數"",                                                                                 
                                  COALESCE(e.""已看初診人數"", 0) AS ""已看初診人數""                                                                                
                           FROM   opd.ptopd a                                                                                                                        
                                  LEFT JOIN opd.dclin b                                                                                                              
                                         ON a.clinic_date = b.clinic_date                                                                                            
                                            AND a.clinic_apn = b.clinic_apn                                                                                          
                                            AND a.clinic_no = b.clinic_no                                                                                            
                                            AND a.doctor_no = b.doctor_no                                                                                            
                                  LEFT JOIN mast.div c                                                                                                               
                                         ON a.div_no = c.div_no                                                                                                      
                                  LEFT JOIN mast.doctor d                                                                                                            
                                         ON a.doctor_no = d.doctor_no                                                                                                
                                  LEFT JOIN (SELECT a.building_no AS ""院區"",                                                                                       
                                                    a.clinic_date AS ""就醫日期"",                                                                                   
                                                    b.week        AS ""星期"",                                                                                       
                                                    a.clinic_apn  AS ""午別"",                                                                                       
                                                    a.clinic_no   AS ""診間號"",                                                                                     
                                                    b.div_no      AS ""科別代碼"",                                                                                   
                                                    b.clinic_name AS ""科別名稱"",                                                                                   
                                                    d.doctor_name AS ""醫師名稱"",                                                                                   
                                                    a.clinic_flag AS ""診別"",                                                                                       
                                                    Count(*)      AS ""已看初診人數""                                                                                
                                             FROM   opd.ptopd a                                                                                                      
                                                    LEFT JOIN opd.dclin b                                                                                            
                                                           ON a.clinic_date = b.clinic_date                                                                          
                                                              AND a.clinic_apn = b.clinic_apn                                                                        
                                                              AND a.clinic_no = b.clinic_no                                                                          
                                                              AND a.doctor_no = b.doctor_no                                                                          
                                                    LEFT JOIN mast.div c                                                                                             
                                                           ON a.div_no = c.div_no                                                                                    
                                                    LEFT JOIN mast.doctor d                                                                                          
                                                           ON a.doctor_no = d.doctor_no                                                                              
                                             WHERE  To_date(To_char(To_number(a.clinic_date) + 19110000),'yyyy/MM/dd') = Trunc(sysdate)                              
                                                    AND a.clinic_flag IN ( 'O', 'V' )                                                                                     
                                                    AND a.korder_flag IN ( 'Y', 'P', 'R' )                                                                           
                                                    AND a.reg_clerk NOT IN ( 'KOPDMISC', 'KOPDCHR', 'KOPDREH', 'KOPDUREG' )                                          
                                                    AND a.fv_rv_flag = '1'                                                                                           
                                             GROUP  BY a.building_no,                                                                                                
                                                       a.clinic_date,                                                                                                
                                                       b.week,                                                                                                       
                                                       a.clinic_apn,                                                                                                 
                                                       a.clinic_no,                                                                                                  
                                                       b.div_no,                                                                                                     
                                                       b.clinic_name,                                                                                                
                                                       d.doctor_name,                                                                                                
                                                       a.clinic_flag                                                                                                 
                                             ORDER  BY a.clinic_flag,                                                                                                
                                                       a.clinic_apn,                                                                                                 
                                                       a.clinic_no)e                                                                                                 
                                         ON a.clinic_date = e.""就醫日期""                                                                                           
                                            AND b.week = e.""星期""                                                                                                  
                                            AND a.clinic_apn = e.""午別""                                                                                            
                                            AND a.clinic_no = e.""診間號""                                                                                           
                                            AND b.div_no = e.""科別代碼""                                                                                            
                                            AND b.clinic_name = e.""科別名稱""                                                                                       
                                            AND d.doctor_name = e.""醫師名稱""                                                                                       
                                            AND a.clinic_flag = e.""診別""                                                                                           
                           WHERE  To_date(To_char(To_number(a.clinic_date) + 19110000), 'yyyy/MM/dd') = Trunc(sysdate)                                               
                                  AND a.clinic_no != ' '                                                                                                             
                                  AND A.div_no != '60'                                                                                                               
                                  AND a.clinic_flag IN ( 'O', 'V' )                                                                                                       
                                  AND a.korder_flag IN ( 'Y', 'P', 'R' )                                                                                             
                                  AND a.reg_clerk NOT IN ( 'KOPDMISC', 'KOPDCHR', 'KOPDREH', 'KOPDUREG' )                                                            
                                  AND a.building_no = 'B'                                                                                                           
                           GROUP  BY a.building_no,                                                                                                                  
                                     a.clinic_date,                                                                                                                  
                                     b.week,                                                                                                                         
                                     a.clinic_apn,                                                                                                                   
                                     a.clinic_no,                                                                                                                    
                                     b.div_no,                                                                                                                       
                                     b.clinic_name,                                                                                                                  
                                     a.clinic_no,                                                                                                                    
                                     d.doctor_name,                                                                                                                  
                                     a.clinic_flag,                                                                                                                  
                                     e.""已看初診人數""                                                                                                              
                                                                                                                                                                     
                            UNION ALL                                                                                                                                
                                                                                                                                                                     
                            SELECT a.building_no                   AS ""院區"",                                                                                      
                                   Count(*)                        AS ""已看診人數"",                                                                                
                                   COALESCE(e.""已看初診人數"", 0) AS ""已看初診人數""                                                                               
                            FROM   opd.ptopd a                                                                                                                       
                                   LEFT JOIN opd.dclin b                                                                                                             
                                          ON a.clinic_date = b.clinic_date                                                                                           
                                             AND a.clinic_apn = b.clinic_apn                                                                                         
                                             AND a.clinic_no = b.clinic_no                                                                                           
                                             AND a.doctor_no = b.doctor_no                                                                                           
                                   LEFT JOIN mast.div c                                                                                                              
                                          ON a.div_no = c.div_no                                                                                                     
                                   LEFT JOIN mast.doctor d                                                                                                           
                                          ON a.doctor_no = d.doctor_no                                                                                               
                                   LEFT JOIN (SELECT a.building_no AS ""院區"",                                                                                      
                                                     a.clinic_date AS ""就醫日期"",                                                                                  
                                                     b.week        AS ""星期"",                                                                                      
                                                     a.clinic_apn  AS ""午別"",                                                                                      
                                                     a.clinic_no   AS ""診間號"",                                                                                    
                                                     b.div_no      AS ""科別代碼"",                                                                                  
                                                     b.clinic_name AS ""科別名稱"",                                                                                  
                                                     d.doctor_name AS ""醫師名稱"",                                                                                  
                                                     a.clinic_flag AS ""診別"",                                                                                      
                                                     Count(*)      AS ""已看初診人數""                                                                               
                                              FROM   opd.ptopd a                                                                                                     
                                                     LEFT JOIN opd.dclin b                                                                                           
                                                            ON a.clinic_date = b.clinic_date                                                                         
                                                               AND a.clinic_apn = b.clinic_apn                                                                       
                                                               AND a.clinic_no = b.clinic_no                                                                         
                                                               AND a.doctor_no = b.doctor_no                                                                         
                                                     LEFT JOIN mast.div c                                                                                            
                                                            ON a.div_no = c.div_no                                                                                   
                                                     LEFT JOIN mast.doctor d                                                                                         
                                                            ON a.doctor_no = d.doctor_no                                                                             
                                              WHERE  To_date(To_char(To_number(a.clinic_date) + 19110000), 'yyyy/MM/dd') = Trunc(sysdate)                            
                                                     AND A.div_no = '60'                                                                                             
                                                     AND a.clinic_flag IN ( 'O', 'V' )                                                                                    
                                                     AND a.korder_flag IN ( 'Y', 'P', 'R' )                                                                          
                                                     AND a.reg_clerk NOT IN ( 'KOPDMISC', 'KOPDCHR','KOPDREH','KOPDUREG' )                                           
                                                     AND a.fv_rv_flag = '1'                                                                                          
                                              GROUP  BY a.building_no,                                                                                               
                                                        a.clinic_date,                                                                                               
                                                        b.week,                                                                                                      
                                                        a.clinic_apn,                                                                                                
                                                        a.clinic_no,                                                                                                 
                                                        b.div_no,                                                                                                    
                                                        b.clinic_name,                                                                                               
                                                        d.doctor_name,                                                                                               
                                                        a.clinic_flag                                                                                                
                                              ORDER  BY a.clinic_flag,                                                                                               
                                                        a.clinic_apn,                                                                                                
                                                        a.clinic_no)e                                                                                                
                                          ON a.clinic_date = e.""就醫日期""                                                                                          
                                             AND b.week = e.""星期""                                                                                                 
                                             AND a.clinic_apn = e.""午別""                                                                                           
                                             AND a.clinic_no = e.""診間號""                                                                                          
                                             AND b.div_no = e.""科別代碼""                                                                                           
                                             AND b.clinic_name = e.""科別名稱""                                                                                      
                                             AND d.doctor_name = e.""醫師名稱""                                                                                      
                                             AND a.clinic_flag = e.""診別""                                                                                          
                            WHERE  To_date(To_char(To_number(a.clinic_date) + 19110000),'yyyy/MM/dd') = Trunc(sysdate)                                               
                                   AND A.div_no = '60'                                                                                                               
                                   AND a.clinic_flag IN ( 'O', 'V' )                                                                                                      
                                   AND a.korder_flag IN ( 'Y', 'P', 'R' )                                                                                            
                                   AND a.reg_clerk NOT IN ( 'KOPDMISC', 'KOPDCHR', 'KOPDREH', 'KOPDUREG' )                                                           
                                   AND a.building_no = 'B'                                                                                                          
                            GROUP  BY a.building_no,                                                                                                                 
                                      a.clinic_date,                                                                                                                 
                                      b.week,                                                                                                                        
                                      a.clinic_apn,                                                                                                                  
                                      a.clinic_no,                                                                                                                   
                                      b.div_no,                                                                                                                      
                                      b.clinic_name,                                                                                                                 
                                      a.clinic_no,                                                                                                                   
                                      d.doctor_name,                                                                                                                 
                                      a.clinic_flag,                                                                                                                 
                                      e.""已看初診人數"")                                                                                                            
                    GROUP  BY ""院區""  
                 ";

                oddap = new OleDbDataAdapter(odcmm);
                oddap.Fill(dt);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (oddap != null) { oddap.Dispose(); }
                if (odcmm != null) { odcmm.Dispose(); }
                if (odcnn != null) { odcnn.Dispose(); }
            }

            return dt;
        }

        public DataTable Emtodaysum()
        {
            Database db = new Database();
            OleDbConnection odcnn = db.GetOleDbConnection("ORACLE_DB_HO");
            OleDbCommand odcmm = new OleDbCommand();
            OleDbDataAdapter oddap = null;
            DataTable dt = new DataTable();

            try
            {
                odcnn.Open();
                odcmm.Connection = odcnn;
                odcmm.CommandType = CommandType.Text;
                odcmm.CommandText =
                                    @"
                                          SELECT CASE                                                                                    
                                                   WHEN building_no = 'A' THEN '員榮'                                                    
                                                   WHEN building_no = 'B' THEN '員生'                                                    
                                                 END                AS ""院區"",                                                         
                                                 A.clinic_date      AS ""看診日期"",                                                     
                                                 Count(building_no) AS ""小計""                                                          
                                          FROM   opd.ptopd A                                                                             
                                                 LEFT JOIN mast.doctor B                                                                 
                                                        ON A.doctor_no = B.doctor_no                                                     
                                                 LEFT JOIN mast.div C                                                                    
                                                        ON A.div_no = C.div_no                                                           
                                          WHERE  1 = 1                                                                                   
                                                 AND To_date(To_number(A.clinic_date) + 19110000, 'yyyymmdd') = Trunc(sysdate)           
                                                 AND A.clinic_flag IN ( 'E' )                                                            
                                                 AND A.korder_flag NOT IN ( 'X' )                                                        
                                                 AND A.course_seq <> '1'                                                                 
                                                 AND A.chronic_med_seq NOT IN ( '2', '3' )                                               
                                                 AND A.building_no = 'B'                                                                
                                          GROUP  BY building_no,                                                                         
                                                    clinic_date 
                                    ";
                oddap = new OleDbDataAdapter(odcmm);
                oddap.Fill(dt);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            finally
            {
                if (oddap != null) { oddap.Dispose(); }
                if (odcmm != null) { odcmm.Dispose(); }
                if (odcnn != null) { odcnn.Dispose(); }
            }
            return (dt);
        }
    }
}
