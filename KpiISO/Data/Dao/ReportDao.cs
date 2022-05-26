using KpiISO.App_Helpers;
using KpiISO.Data.Model;

using System;
using System.Collections.Generic;
using System.Data;
using Npgsql;

namespace KpiISO.Data.Dao
{
    public class ReportDao : BaseDao
    {
        public List<DataKpiHO> GetKpiHO(PrmGetRpt prm)
        {
            try
            {
                using (var conn = new NpgsqlConnection(DB_CONNECTION))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand("spl_get_data_temp_kpi_iso_ho", conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("in_lobid", NpgsqlTypes.NpgsqlDbType.Uuid, string.IsNullOrEmpty(prm.in_lob) ? (object)DBNull.Value : prm.in_lob); //XXXX,XXXX,XXXX
                        cmd.Parameters.AddWithValue("in_projid", NpgsqlTypes.NpgsqlDbType.Text, string.IsNullOrEmpty(prm.in_projid) ? (object)DBNull.Value : prm.in_projid); //XXXX,XXXX,XXXX
                        cmd.Parameters.AddWithValue("in_period", NpgsqlTypes.NpgsqlDbType.Text, string.IsNullOrEmpty(prm.in_period) ? (object)DBNull.Value : prm.in_period); //XXXX,XXXX,XXXX
                        cmd.Parameters.AddWithValue("in_projtype", NpgsqlTypes.NpgsqlDbType.Text, string.IsNullOrEmpty(prm.in_projtype) ? (object)DBNull.Value : prm.in_projtype); //XXXX,XXXX,XXXX

                        using (var reader = cmd.ExecuteReader())
                        {
                            return SQLDataMapper.MapToCollection<DataKpiHO>(reader);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public List<DataKpiFix> GetKpiFix(PrmGetRpt prm)
        {
            try
            {
                using (var conn = new NpgsqlConnection(DB_CONNECTION))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand("spl_get_data_temp_kpi_iso_fix", conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("in_lobid", NpgsqlTypes.NpgsqlDbType.Uuid, string.IsNullOrEmpty(prm.in_lob) ? (object)DBNull.Value : prm.in_lob); //XXXX,XXXX,XXXX
                        cmd.Parameters.AddWithValue("in_projid", NpgsqlTypes.NpgsqlDbType.Text, string.IsNullOrEmpty(prm.in_projid) ? (object)DBNull.Value : prm.in_projid); //XXXX,XXXX,XXXX
                        cmd.Parameters.AddWithValue("in_period", NpgsqlTypes.NpgsqlDbType.Text, string.IsNullOrEmpty(prm.in_period) ? (object)DBNull.Value : prm.in_period); //XXXX,XXXX,XXXX
                        cmd.Parameters.AddWithValue("in_projtype", NpgsqlTypes.NpgsqlDbType.Text, string.IsNullOrEmpty(prm.in_projtype) ? (object)DBNull.Value : prm.in_projtype); //XXXX,XXXX,XXXX

                        using (var reader = cmd.ExecuteReader())
                        {
                            return SQLDataMapper.MapToCollection<DataKpiFix>(reader);
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