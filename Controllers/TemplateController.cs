using Microsoft.AspNetCore.Mvc;

namespace ShowPms.Controllers
{
    public class TemplateController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }
    }
}
