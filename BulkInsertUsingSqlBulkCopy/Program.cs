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
            #region SQLBulkCopy with in-memory DataTable
            ////SQLBulkCopy with in-memory DataTable
            //DataTable dtEmployees = new DataTable();

            //dtEmployees.Columns.Add("Id");
            //dtEmployees.Columns.Add("FirstName");
            //dtEmployees.Columns.Add("LastName");
            //dtEmployees.Columns.Add("Gender");
            //dtEmployees.Columns.Add("City");
            //dtEmployees.Columns.Add("IsActive");

            //dtEmployees.Rows.Add(1, "Ebenezer", "McGruar", "Male", "Sanshan", true);
            //dtEmployees.Rows.Add(2, "Yanaton", "Lennon", "Male", "Eshowe", false);
            //dtEmployees.Rows.Add(3, "Etienne", "Rowlatt", "Male", "Yueyang", true);
            //dtEmployees.Rows.Add(4, "Ketty", "Guerri", "Female", "Shatou", true);
            //dtEmployees.Rows.Add(5, "Tabbie", "Auten", "Genderfluid", "Prado", false);


            //string dbConnectionStr = @"Data Source=(local); Initial Catalog=AdoNetDb; Integrated Security=SSPI;";

            //using (SqlConnection con = new SqlConnection(dbConnectionStr))
            //{
            //    using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
            //    {

            //        sqlBulkCopy.DestinationTableName = "[dbo].[Employees]";

            //        //Optional
            //        //sqlBulkCopy.ColumnMappings.Add("Id", "Id");
            //        //sqlBulkCopy.ColumnMappings.Add("FirstName", "FirstName");
            //        //sqlBulkCopy.ColumnMappings.Add("LastName", "LastName");
            //        //sqlBulkCopy.ColumnMappings.Add("Gender", "Gender");
            //        //sqlBulkCopy.ColumnMappings.Add("City", "City");
            //        //sqlBulkCopy.ColumnMappings.Add("IsActive", "IsActive");

            //        con.Open();
            //        sqlBulkCopy.WriteToServer(dtEmployees);
            //        Console.WriteLine("Bulk Insert Successful");

            //    }
            //}
            #endregion SQLBulkCopy with in-memory DataTable

            #region SQL Bulk Copy with CSV File
            //string csvFilePath = @"D:\FilesToLoad\AdoNet\Employees.csv";

            //DataTable dtEmployees = new DataTable();

            //dtEmployees.Columns.Add("Id");
            //dtEmployees.Columns.Add("FirstName");
            //dtEmployees.Columns.Add("LastName");
            //dtEmployees.Columns.Add("Gender");
            //dtEmployees.Columns.Add("City");
            //dtEmployees.Columns.Add("IsActive");

            ////Read data from csv and fill it in data table (dtEmployees)

            //string csvData = File.ReadAllText(csvFilePath);
            //foreach (string row in csvData.Split('\n'))
            //{

            //    if (!string.IsNullOrEmpty(row))
            //    {
            //        dtEmployees.Rows.Add();
            //        int i = 0;
            //        foreach (string cell in row.Split(','))
            //        {
            //            dtEmployees.Rows[dtEmployees.Rows.Count - 1][i] = cell;
            //            i++;
            //        }
            //    }
            //}

            //string dbConnectionStr = @"Data Source=(local); Initial Catalog=AdoNetDb; Integrated Security=SSPI;";

            //using (SqlConnection con = new SqlConnection(dbConnectionStr))
            //{
            //    using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
            //    {

            //        sqlBulkCopy.DestinationTableName = "[dbo].[Employees]";

            //        //Optional
            //        //sqlBulkCopy.ColumnMappings.Add("Id", "Id");
            //        //sqlBulkCopy.ColumnMappings.Add("FirstName", "FirstName");
            //        //sqlBulkCopy.ColumnMappings.Add("LastName", "LastName");
            //        //sqlBulkCopy.ColumnMappings.Add("Gender", "Gender");
            //        //sqlBulkCopy.ColumnMappings.Add("City", "City");
            //        //sqlBulkCopy.ColumnMappings.Add("IsActive", "IsActive");

            //        con.Open();
            //        sqlBulkCopy.WriteToServer(dtEmployees);
            //        Console.WriteLine("Bulk Insert Successful");

            //    }
            //}
            #endregion #region SQL Bulk Copy with CSV File

            #region SQL Bulk Cpoy on Excel File Data
            string excelFilePath = @"D:\FilesToLoad\AdoNet\Employees.xlsx";
            string excelConStr = string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0 Xml;HDR=YES;'", excelFilePath);

            using (OleDbConnection xlCon = new OleDbConnection(excelConStr))
            {
                xlCon.Open();
                string sheetName = xlCon.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null).Rows[0]["TABLE_NAME"].ToString();
                DataTable dtXlTable = new DataTable();
                dtXlTable.Columns.AddRange(new DataColumn[6] {
                new DataColumn("Id", typeof(int)),
                new DataColumn("FirstName", typeof(string)),
                new DataColumn("LastName", typeof(string)),
                new DataColumn("Gender", typeof(string)),
                new DataColumn("City", typeof(string)),
                new DataColumn("IsActive", typeof(bool))
                });

                using (OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM ["+ sheetName + "]",xlCon))
                {
                    oda.Fill(dtXlTable);
                }

                string dbConnectionStr = @"Data Source=(local); Initial Catalog=AdoNetDb; Integrated Security=SSPI;";

                using (SqlConnection con = new SqlConnection(dbConnectionStr))
                {
                    using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                    {

                        sqlBulkCopy.DestinationTableName = "[dbo].[Employees]";

                        //Optional
                        //sqlBulkCopy.ColumnMappings.Add("Id", "Id");
                        //sqlBulkCopy.ColumnMappings.Add("FirstName", "FirstName");
                        //sqlBulkCopy.ColumnMappings.Add("LastName", "LastName");
                        //sqlBulkCopy.ColumnMappings.Add("Gender", "Gender");
                        //sqlBulkCopy.ColumnMappings.Add("City", "City");
                        //sqlBulkCopy.ColumnMappings.Add("IsActive", "IsActive");

                        con.Open();
                        sqlBulkCopy.WriteToServer(dtXlTable);
                        Console.WriteLine("Bulk Insert Successful");

                    }
                }
            }


           

            #endregion SQL Bulk Cpoy on Excel File Data



            Console.ReadLine();
        }
    }
}
