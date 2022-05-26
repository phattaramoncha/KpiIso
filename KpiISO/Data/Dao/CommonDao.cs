using Npgsql;
using KpiISO.App_Helpers;
using KpiISO.Data.Model;
using System;
using System.Collections.Generic;
using System.Data;

namespace KpiISO.Data.Dao
{
    public class CommonDao : BaseDao
    {

        public List<Line_Bus> getlineBus ()
        {
            try
            {
                using (var conn = new NpgsqlConnection(DB_CONNECTION))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand("spl_get_kpi_lob", conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;

                        using (var reader = cmd.ExecuteReader())
                        {
                            return SQLDataMapper.MapToCollection<Line_Bus>(reader);
                            //var dataTable = new DataTable();
                            //dataTable.Load(reader);

                            //System.Web.Script.Serialization.JavaScriptSerializer serializer = new System.Web.Script.Serialization.JavaScriptSerializer();
                            //List<Dictionary<string, object>> rows = new List<Dictionary<string, object>>();
                            //Dictionary<string, object> row;
                            //foreach (DataRow dr in dataTable.Rows)
                            //{
                            //    row = new Dictionary<string, object>();
                            //    foreach (DataColumn col in dataTable.Columns)
                            //    {
                            //        row.Add(col.ColumnName, dr[col]);
                            //    }
                            //    rows.Add(row);
                            //}
                            //return serializer.Serialize(rows);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        public List<Project> getProj(string lob_id, string proj_type)
        {
            try
            {
                using (var conn = new NpgsqlConnection(DB_CONNECTION))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand("spl_get_kpi_proj", conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("in_lobid", NpgsqlTypes.NpgsqlDbType.Uuid, string.IsNullOrEmpty(lob_id) ? (object)DBNull.Value : lob_id);
                        cmd.Parameters.AddWithValue("in_proj_type", NpgsqlTypes.NpgsqlDbType.Text, string.IsNullOrEmpty(proj_type) ? (object)DBNull.Value : proj_type);

                        using (var reader = cmd.ExecuteReader())
                        {
                            return SQLDataMapper.MapToCollection<Project>(reader);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
    
       
    }
}