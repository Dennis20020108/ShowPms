using Dapper;
using ShowPms.Data;
using System.Data;
using ShowPms.DTOs;

namespace ShowPms.Repositories
{
    public class OldPriceRepository
    {
        private readonly DapperContext _context;

        public OldPriceRepository(DapperContext context)
        {
            _context = context;
        }

        public async Task<IEnumerable<OldPriceItemDto>> GetOldPriceItemsAsync(int sourceId, int vendorId)
        {
            var query = @"
                SELECT 
                    i.Id,
                    i.Name,
                    i.Unit,
                    i.Quantity,
                    i.UnitPrice,
                    i.Amount,
                    i.ContractUnitPrice,
                    i.ContractAmount,
                    i.Note,
                    mi.Name AS MinorCategoryName,
                    m.Name AS MiddleCategoryName,
                    ma.Name AS MajorCategoryName
                FROM OldPriceItem i
                INNER JOIN MinorCategory mi ON i.MinorCategoryId = mi.Id
                INNER JOIN MiddleCategory m ON mi.MiddleCategoryId = m.Id
                INNER JOIN MajorCategory ma ON m.MajorCategoryId = ma.Id
                WHERE i.SourceId = @SourceId AND i.VendorId = @VendorId
                ORDER BY ma.Id, m.Id, mi.Id, i.Id";

            using var connection = _context.CreateMssqlConnection();
            var result = await connection.QueryAsync<OldPriceItemDto>(query, new { SourceId = sourceId, VendorId = vendorId });

            return result;
        }
    }

}
