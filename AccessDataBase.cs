using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//using Microsoft.Office.Interop.Access;
using System.Data;
using System.Data.OleDb;
using System.ComponentModel;
using System.Windows.Forms;


namespace FCHelper_v001
{
    class AccessDataBase
    {
        private string windowsUserName = System.Environment.UserName;//gives windows username
        private OleDbConnection connection = new OleDbConnection();

        public void AccessDataBaseMethod(string ername,  string erid,    string region,
                                         string segment, string effDate, string curProd,
                                         string addProd, string newImp,  string AM_IM,
                                         string impDdline,
                                         string exConName,  string exConPhone,
                                         string exConEmail, string exConType, string fileType
                                         )//inputs to method
        {
            string accessDBFilePath = "C:\\Users\\PA155965\\Documents\\ImpList.txt";
            connection.ConnectionString = @"PROVIDER = Microsoft.ACE.OLEDB.12.0; "+
                                            "DATA SOURCE = "+ accessDBFilePath + "; " +
                                            "PERSIST SECURITY INFO = False;";
            try
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand();
                command.Connection = connection;
                command.CommandText = "INSERT INTO Implementation List (ERname, ERID, Region, Segment," +
                    "                                                   EffDate, CurrentProduct, AddingProduct," +
                    "                                                   NewImplementation, AM_IM, ImplementationDeadline," +
                    "                                                   ExternalContactName, ExternalContactPhone," +
                    "                                                   ExternalContactEmail, ExternalContactType," +
                    "                                                   FileType) " +
                                      "VALUES ('" + ername + "','" + erid + "','" + region +
                                              "','" + segment + "','" + effDate + "','" + curProd + "" +
                                              "','" + addProd + "','" + newImp + "','" + AM_IM + "" +
                                              "','" + impDdline + "','" + exConName + "','" + exConPhone + "" +
                                              "','" + exConEmail + "','" + exConType + "','" + exConType + "','" + fileType + "')";

                command.ExecuteNonQuery();
                MessageBox.Show("Implementation saved.", "Data pushed", MessageBoxButtons.OK, MessageBoxIcon.Information);

               // connection.Close();
            }
            catch(Exception ex)
            {
                MessageBox.Show("Something prevented the data from pushing to the Access database.\r\r"+ex,"Error",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
            
            

            
        }
    }
}
