using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace RecipeIPlacementCount
{
    class db
    {
        //Data Source=migsqlclu4\smt;Initial Catalog=Explorer_ASM;Persist Security Info=True;User ID=kati ; password= $Flex2020"
        private static string GetConnectionString()
        {
            //Data Source=172.20.20.2;Initial Catalog=SiplacePro;Persist Security Info=True;User ID=aoi
            string conString = @"Data Source=172.20.20.2;Initial Catalog=SiplacePro;Persist Security Info=True;User ID=aoi; Password =$Flex2016";
            return conString;
        }
        public static SqlConnection con = new SqlConnection();
        public static SqlCommand cmd = new SqlCommand("", con);//INSERT, UPDATE, DELETE, SELECT
        public static SqlDataReader rd;// SqlDataReader;// SqlDataReader
        public static DataSet ds;
        public static SqlDataAdapter da;
        public static BindingSource bs;

        //SELECT, INSERT, UPDATE, DELETE
        public static string sql;
        //Open Database
        public static void openConnection()
        {
            if (con.State == ConnectionState.Closed)
            {
                con.ConnectionString = GetConnectionString();
                con.Open();
                //MessageBox.Show("The connection is "+ con.State.ToString());
            }
        }
        //Close DataBase
        public static void closeConnection()
        {
            if (con.State == ConnectionState.Open)
            {
                con.Close();
                //MessageBox.Show("The connection is " + con.State.ToString());
            }
        }
    }
}
