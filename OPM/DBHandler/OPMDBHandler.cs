using System;
using System.Data.SqlClient;

namespace OPM.DBHandler
{

    class OPMDBHandler
    {

        public static int fInsertData(string strSqlCommand)
        {
            String strconnection = @"Data Source=DOANTD; Initial Catalog = OpmDB; User ID = sa; Password=Pa$$w0rd";
            SqlConnection con = new SqlConnection(strconnection);
            try
            {
                con.Open();
                using (SqlCommand insertCommand = con.CreateCommand())
                {
                    insertCommand.CommandText = strSqlCommand;
                    var row = insertCommand.ExecuteNonQuery();
                }
                con.Close();
                con.Dispose();
                con = null;
                return 1;
            }
            catch(Exception e)
            {
                con.Close();
                con.Dispose();
                con = null;
                return 0;
            }

        }
        public static int fQuerryData(string strQuerry)
        {
            String strconnection = @"Data Source=DOANTD; Initial Catalog = OpmDB; User ID = sa; Password=Pa$$w0rd";
            SqlConnection con = new SqlConnection(strconnection);
            try
            {
                con.Open();
                using (SqlCommand sCommand = con.CreateCommand())
                {
                    sCommand.CommandText = strQuerry;
                    using (SqlDataReader reader = sCommand.ExecuteReader())
                    {
                        if(reader.HasRows)
                        {
                            while(reader.Read())
                            {
                                //Do Sothing
                            }    
                        }
                        else
                        {
                            //Do Something
                        }    

                    }    
                }
                con.Close();
                con.Dispose();
                con = null;
                return 1;

            }
            catch(Exception e)
            {
                con.Close();
                con.Dispose();
                con = null;
                return 0;
            }
        }
    }
}
