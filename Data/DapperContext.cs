using System.Data;
using Npgsql;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;

namespace ShowPms.Data
{
    public class DapperContext
    {
        private readonly IConfiguration _configuration;
        private readonly string _mssqlConnectionString;
        private readonly string _npgsqlConnectionString;

        public DapperContext(IConfiguration configuration) 
        {
            _configuration = configuration;
            _mssqlConnectionString = _configuration.GetConnectionString("MssqlConnection");
            _npgsqlConnectionString = _configuration.GetConnectionString("NpgsqlConnection");
        }

        public IDbConnection CreateMssqlConnection() => new SqlConnection(_mssqlConnectionString);

        public IDbConnection CreateNpgsqlConnection() => new NpgsqlConnection(_npgsqlConnectionString);
    }
}
