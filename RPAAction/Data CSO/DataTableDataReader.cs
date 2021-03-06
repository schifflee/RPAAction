using System.Data;

namespace RPAAction.Data_CSO
{
    public class DataTableDataReader : RPADataReader
    {
        public DataTableDataReader(DataTable table)
        {
            this.table = table;
        }

        public override int FieldCount => table.Columns.Count;

        public override bool IsClosed => true;

        public override void Close()
        {

        }

        public override string GetName(int ordinal)
        {
            return table.Columns[ordinal].ColumnName;
        }

        public override object GetValue(int ordinal)
        {
            return table.Rows[readingRowIndex][ordinal];
        }

        public override bool Read()
        {
            return ++readingRowIndex < table.Rows.Count;
        }

        private int readingRowIndex = -1;
        private readonly DataTable table;
    }
}
