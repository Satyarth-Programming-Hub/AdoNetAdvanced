using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BulkInsertUsingSqlBulkCopy
{
    class Program
    {
        static void Main(string[] args)
        {

            try
            {
                string constr = @"Data Source=(local); Initial Catalog=AdoNetDb;Integrated Security=SSPI;";
                DataTable empDataTable = new DataTable();
                empDataTable.Columns.Add("Id");
                empDataTable.Columns.Add("FirstName");
                empDataTable.Columns.Add("LastName");
                empDataTable.Columns.Add("Gender");
                empDataTable.Columns.Add("City");
                empDataTable.Columns.Add("IsActive");

                //Rows to be updated
                empDataTable.Rows.Add(1, "Arun", "Singh", "Male", "Kanpur", "True");
                empDataTable.Rows.Add(2, "Priya", "Rawat", "Female", "Jamnagar", "False");
                empDataTable.Rows.Add(3, "Preety", "Shina", "Female", "Jaipur", "True");

                //Rows to be inserted
                empDataTable.Rows.Add(6, "Rohit", "Verma", "Male", "Kanpur", "True");
                empDataTable.Rows.Add(7, "Rehan", "Khan", "Male", "Jaipur", "True");
                empDataTable.Rows.Add(8, "Seem", "Shina", "Female", "Jamnagar", "True");

                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand("BulkInsertUpdate_OnSQLServerVersion2005OrAbove", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.AddWithValue("@Employees", empDataTable);

                        con.Open();

                        cmd.ExecuteNonQuery();

                        Console.WriteLine("Process executed successfully");
                    }
                }
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);
            }
            

            Console.ReadLine();
        }
    }
}
