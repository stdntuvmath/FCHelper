using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;
using System.Windows.Forms;

namespace FCHelper_v001
{
    class GetGroupID
    {
        private string windowsUserName = System.Environment.UserName;//gives windows username
        private SqlConnection sqlConnection;
        private SqlCommand sqlCommand;

        public void GetGroupIDMethod(string ERID)
        {

           
            PrivateGetGroupIDMethod(ERID);
            
        }

        private void PrivateGetGroupIDMethod(string erid)
        {

            string connectionString = @"Server = phx-edidb-01\edi,49565; Database = DataServices; User Id = PA155965; Password = Yrthsa12``; Trusted_Connection=True;";

            sqlConnection = new SqlConnection(connectionString);
            

            //try
            //{
            //    sqlConnection.ConnectionString = @"Server = phx-edidb-01\edi,49565; Database = DataServices; User Id = A155965; Password = Yrthsa12``; Trusted_Connection=True;";

            //}
            //catch (ArgumentException ex)
            //{
            //    MessageBox.Show(ex.ToString());
            //}


            string sqlPullString = "using DataServices \r\n SELECT * FROM dbo.CBAS_Employers WHERE Employer_ID = "+erid+";";
            SqlDataAdapter sda = new SqlDataAdapter(sqlPullString, sqlConnection); 
            

            
            try
            {
                sqlConnection.Open();
                //sda.Fill();
                MessageBox.Show("privateTable after sda.Fill(): " );
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.ToString());
            }
           
           
            sqlConnection.Dispose();
        }    
    }
}
