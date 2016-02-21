using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.Linq;
using CTT.ScheduleMaker.Model;

namespace CTT.ScheduleMaker.Importers
{
    internal static class XlsImport
    {
        internal static DataSet ReadExcel(string path)
        {
            var ds = new DataSet();

            var connectionString =
                $"Provider=Microsoft.ACE.OLEDB.12.0;Data source={path};Extended Properties=Excel 12.0;";
            using (var conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                var cmd = new OleDbCommand { Connection = conn };

                var dtSheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                if (dtSheet == null)
                    return ds;
                foreach (DataRow dr in dtSheet.Rows)
                {
                    string sheetName = dr["TABLE_NAME"].ToString();
                    if (!sheetName.EndsWith("$"))
                        continue;

                    // Get all rows from the Sheet
                    cmd.CommandText = "SELECT * FROM [" + sheetName + "]";

                    var dt = new DataTable { TableName = sheetName };
                    var da = new OleDbDataAdapter(cmd);
                    da.Fill(dt);

                    ds.Tables.Add(dt);
                }
            }

            return ds;
        }



        internal static IEnumerable<ScheduleItem> ReadFromDataset(DataSet ds, string sheet)
        {
            var list = new List<ScheduleItem>();

            var rowCollection = ds.Tables[sheet + "$"].Rows;
            var rows = rowCollection.Cast<DataRow>().ToList();
            
            foreach (var row in rows)
            {
                var isEmptyRow = row[0] is DBNull && row[1] is DBNull ;
                if (isEmptyRow)
                    break;

                list.AddRange(ParsePriceRow(row, list));
            }
            return list;
        }


        private static IEnumerable<ScheduleItem> ParsePriceRow(DataRow row, ICollection<ScheduleItem> list)
        {
            var unity = GetValue(row[0]).Trim();
            var @group = GetValue(row[1]).Trim();
            var professor = GetValue(row[2]).Trim();
            var auditory = GetValue(row[3])?.Trim();
            for (var i = 0; i != 7; ++i)
            {
                var dayOfWeek = (DayOfWeek) ((i + 1)%7);
                var timesString = GetValue(row[4+i]);
                if (string.IsNullOrWhiteSpace(timesString))
                    continue;

                timesString = timesString.Replace(" -", "-");
                timesString = timesString.Replace("- ", "-");
                timesString = timesString.Replace("--", "-");

                var times = timesString.Split(' ');

                foreach (var time in times)
                {
                    if (string.IsNullOrWhiteSpace(time))
                        continue;

                    var pair = time.Split('-');
                    
                    yield return new ScheduleItem
                    {
                        Unity = unity,
                        Group = @group,
                        Professor = professor,
                        Auditory = auditory,
                        DayOfWeek = dayOfWeek,
                        TimeBegin = pair[0],
                        TimeEnd = pair[1],
                    };
                }
            }
        }


        private static string GetValue(object value)
        {
            if (value is DBNull)
                return null;
            if (value is string)
                return value as string;
            if (value is double)
                return ((double) value).ToString(CultureInfo.CurrentCulture);
            throw new ArgumentOutOfRangeException(nameof(value));
        }
    }
}