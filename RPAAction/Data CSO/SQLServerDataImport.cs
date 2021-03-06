using RPAAction.Base;
using System;
using System.Data.Common;
using System.Data.SqlClient;

namespace RPAAction.Data_CSO
{
    public class SQLServerDataImport : RPADataImport
    {
        public SQLServerDataImport(string DataSource, string DataBase, string user, string pwd, string table)
        {
            connStr = $@"Data Source={DataSource};Initial Catalog={DataBase};User ID={user};Pwd={pwd};";
            conn = new SqlConnection(connStr);
            conn.Open();
            tableName = table;
        }

        public SQLServerDataImport(string connStr, string table)
        {
            this.connStr = connStr;
            conn = new SqlConnection(connStr);
            conn.Open();
            tableName = table;
        }

        public SQLServerDataImport(SqlConnection conn, string table)
        {
            this.conn = conn;
            tableName = table;
        }

        public override void ImportFrom(DbDataReader reader)
        {
            try
            {
                CreateTable(reader);
            }
            catch (Exception e)
            {
                if (e is ActionException)
                    throw e;
            }

            using (SqlBulkCopy bulkCopy = new SqlBulkCopy(conn))
            {
                bulkCopy.DestinationTableName = tableName;
                bulkCopy.WriteToServer(reader);
            }
        }

        protected override void Close()
        {
            if (connStr == null)
            {
                conn.Dispose();
            }
        }

        protected override void SetValue(string field, object value)
        {
            throw new NotImplementedException();
        }

        protected override void UpdataRow()
        {
            throw new NotImplementedException();
        }

        protected override void CreateTable(DbDataReader r)
        {
            string sql = GetCreateTableString(r, "text");
            var cmd = new SqlCommand(sql.ToString(), conn);
            cmd.ExecuteNonQuery();
        }

        private readonly string connStr = null;
        private readonly SqlConnection conn;
    }
}
