using OfficeOpenXml;
using VT1.Models;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Text;

namespace VT1.Services
{
    public class DbService
    {
        private Table table;
        private const string _connectionString = "Server=.;initial catalog={0};Integrated Security=SSPI";

        public DbService(Table table)
        {
            this.table = table;
        }


        public void CreateTable()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine($"IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='{table.name}' and xtype='U')");
            sb.AppendLine("BEGIN");
            sb.AppendLine($"CREATE TABLE {table.name} (");

            for (int colIdx = 0; colIdx < table.columnCount; colIdx++)
            {
                string prefix = " ";
                if (colIdx != 0) prefix = ",";
                sb.AppendLine($"{prefix}[{table.columns[colIdx]}] NVARCHAR(255) NULL");
            }
            sb.AppendLine(")");
            sb.AppendLine("END");
            ExecCommand(sb.ToString(), GetConnectionString("Test"));
            
        }

        public void TableInsert()
        {
            DataTable tbl = new DataTable();

            for (int colIdx = 0; colIdx < table.columnCount; colIdx++)
            {
                tbl.Columns.Add(new DataColumn(table.columns[colIdx], typeof(string)));

            }

            for (int i = 0; i < table.rowCount; i++)
            {
                DataRow dr = tbl.NewRow();

                for (int j = 0; j < table.columnCount; j++)
                {
                    dr[table.columns[j]] = table.values[i, j];

                }
                tbl.Rows.Add(dr);
            }

            string connection = GetConnectionString("Test");
            using (SqlConnection con = new SqlConnection(connection))
            {
                con.Open();
                using (SqlTransaction transaction = con.BeginTransaction())
                {
                    SqlBulkCopy objbulk = new SqlBulkCopy(con, SqlBulkCopyOptions.KeepIdentity, transaction);
                    objbulk.DestinationTableName = table.name;                    
                    
                    try
                    {
                        objbulk.WriteToServer(tbl);
                        transaction.Commit();
                    }

                    catch
                    {
                        transaction.Rollback();
                        throw;

                    }
                }
            }
        }

        // 0 = DB name, 1 = DB path
        private const string createDbCmd = @"
IF NOT EXISTS(SELECT * FROM sys.databases WHERE name = '{0}')
BEGIN
	CREATE DATABASE [{0}]
END
";

        public void CreateDb(string DbName)
        {            
            var cmd = string.Format(createDbCmd, DbName);
            ExecCommand(cmd, GetConnectionString("master"));
        }

        private static void ExecCommand(string queryString, string connectionString) 
        { 
            using (SqlConnection connection = new SqlConnection(connectionString))
            { 
                SqlCommand command = new SqlCommand(queryString, connection);
                command.Connection.Open();
                command.ExecuteNonQuery();
            } 
        }

        public string GetConnectionString(string DbName)
        {
            return string.Format(_connectionString, DbName);            
        }
    }
}

