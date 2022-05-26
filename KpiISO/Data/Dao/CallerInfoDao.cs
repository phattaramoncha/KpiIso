using MySqlConnector;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;

namespace _3CXInfo.Data.Dao
{
    public class CallerInfoDao : BaseDao
    {
        protected readonly String DB_CONNECTION = ConfigurationManager.ConnectionStrings["DB_3CX"].ConnectionString;
        public string getCallerInfo(string numberPhone, string dtSrt, string dtEnd)
        {
            using (MySqlConnection conn = new MySqlConnection(DB_CONNECTION))
            {
                conn.Open();
                using (MySqlCommand command = conn.CreateCommand())
                {
                    command.CommandText = "SELECT from_dn, from_dispname, to_dn, DATE_FORMAT(duration, '%T') as duration, " +
                        "DATE_FORMAT(CONVERT_TZ(time_start, 'UTC', 'Asia/Bangkok'), '%Y-%m-%d %T') as time_start, " +
                        "DATE_FORMAT(CONVERT_TZ(time_end, 'UTC', 'Asia/Bangkok'), '%Y-%m-%d %T')   as time_end, " +
                        "to_type " +
                        "FROM cdr WHERE to_dispname <> 'IP Phone Test' AND to_dn = '" + numberPhone + "' " +
                        "AND DATE_FORMAT(time_start, '%Y-%m-%d') >= '" + dtSrt + "' AND DATE_FORMAT(time_end, '%Y-%m-%d') <= '" + dtEnd + "' " +
                        "ORDER BY time_start DESC;";

                    using (MySqlDataReader reader = command.ExecuteReader())
                    {
                        //return SQLDataMapper.MapToCollection<string>(reader);
                        var dataTable = new DataTable();
                        dataTable.Load(reader);

                        System.Web.Script.Serialization.JavaScriptSerializer serializer = new System.Web.Script.Serialization.JavaScriptSerializer();
                        List<Dictionary<string, object>> rows = new List<Dictionary<string, object>>();
                        Dictionary<string, object> row;
                        foreach (DataRow dr in dataTable.Rows)
                        {
                            row = new Dictionary<string, object>();
                            foreach (DataColumn col in dataTable.Columns)
                            {
                                row.Add(col.ColumnName, dr[col]);
                            }
                            rows.Add(row);
                        }
                        return serializer.Serialize(rows);
                    };
                }
            }
        }

        public string getCallerInfoTODAY()
        {
            using (MySqlConnection conn = new MySqlConnection(DB_CONNECTION))
            {
                conn.Open();
                using (MySqlCommand command = conn.CreateCommand())
                {
                    command.CommandText = "SELECT from_dn, from_dispname, to_dn, DATE_FORMAT(duration, '%T') as duration, " +
                        "DATE_FORMAT(CONVERT_TZ(time_start, 'UTC', 'Asia/Bangkok'), '%Y-%m-%d %T') as time_start, " +
                        "DATE_FORMAT(CONVERT_TZ(time_end, 'UTC', 'Asia/Bangkok'), '%Y-%m-%d %T')   as time_end, " +
                        "to_type " +
                        "FROM cdr WHERE to_dispname <> 'IP Phone Test' AND DATE_FORMAT(time_start, '%Y-%m-%d') = DATE_FORMAT(NOW(), '%Y-%m-%d') " +
                        "ORDER BY time_start DESC;";

                    using (MySqlDataReader reader = command.ExecuteReader())
                    {
                        //return SQLDataMapper.MapToCollection<string>(reader);
                        var dataTable = new DataTable();
                        dataTable.Load(reader);

                        System.Web.Script.Serialization.JavaScriptSerializer serializer = new System.Web.Script.Serialization.JavaScriptSerializer();
                        List<Dictionary<string, object>> rows = new List<Dictionary<string, object>>();
                        Dictionary<string, object> row;
                        foreach (DataRow dr in dataTable.Rows)
                        {
                            row = new Dictionary<string, object>();
                            foreach (DataColumn col in dataTable.Columns)
                            {
                                row.Add(col.ColumnName, dr[col]);
                            }
                            rows.Add(row);
                        }
                        return serializer.Serialize(rows);
                    };
                }
            }
        }
    }
}