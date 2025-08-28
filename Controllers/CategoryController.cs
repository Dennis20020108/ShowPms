using Microsoft.AspNetCore.Mvc;

namespace ShowPms.Controllers
{
    public class CategoryController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }
    }
}
