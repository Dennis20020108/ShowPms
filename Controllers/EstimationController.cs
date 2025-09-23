using ShowPms.DTOs;
using Microsoft.AspNetCore.Mvc;

namespace ShowPms.Controllers
{
    public class EstimationController : Controller
    {
        private readonly EstimationServices _estimationServices;

        public EstimationController(EstimationServices estimationServices)
        {
            _estimationServices = estimationServices;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public IActionResult ExportExcel([FromBody] FlatExportRequest flatRequest)
        {
            try
            {
                if (flatRequest == null || string.IsNullOrWhiteSpace(flatRequest.ProjectName))
                    return BadRequest("請輸入專案名稱");

                if (flatRequest.Items == null || !flatRequest.Items.Any())
                    return BadRequest("請選擇至少一個項目");

                // 將平面 items 分組轉階層
                var majorGroups = flatRequest.Items
                    .GroupBy(i => i.MajorCategory ?? "")
                    .Select(g => new MajorCategoryDto
                    {
                        Name = g.Key,
                        MiddleItems = g
                            .GroupBy(i => i.MiddleCategory ?? "")
                            .Select(mg => new MiddleCategoryDto
                            {
                                Name = mg.Key,
                                Items = mg.Select(x => new ExportRequestItem
                                {
                                    Id = x.Id,
                                    Code = x.Code,
                                    Vender = x.Vender,
                                    Name = x.Name,
                                    Spec = x.Spec,
                                    Unit = x.Unit,
                                    Quantity = x.Quantity,
                                    UnitPrice = x.UnitPrice,
                                    ContractUnitPrice = x.ContractUnitPrice,
                                    Note = x.Note
                                }).ToList()
                            }).ToList()
                    }).ToList();

                var request = new ExportRequest
                {
                    ProjectName = flatRequest.ProjectName,
                    MajorItems = majorGroups
                };

                var bytes = _estimationServices.ExportExcel(request);
                string fileName = $"{flatRequest.ProjectName}_估價單.xlsx";
                return File(bytes,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    fileName);
            }
            catch (Exception ex)
            {
                // log ex
                return BadRequest("匯出失敗：" + ex.Message + "\n" + ex.StackTrace);
            }
        }



        public IActionResult List()
        {
            return View();
        }
    }
}
