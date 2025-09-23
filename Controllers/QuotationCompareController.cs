using Microsoft.AspNetCore.Mvc;

namespace ShowPms.Controllers
{
    public class QuotationCompareController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Import()
        {
            return View();
        }
    }
}
