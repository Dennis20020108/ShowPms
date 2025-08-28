using Microsoft.AspNetCore.Mvc;

namespace ShowPms.Controllers
{
    public class UserManageController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }
    }
}
