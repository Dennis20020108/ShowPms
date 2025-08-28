using Microsoft.AspNetCore.Mvc;

namespace ShowPms.Controllers
{
    public class DashboardController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }
    }
}
