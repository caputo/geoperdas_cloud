using Microsoft.AspNetCore.Mvc;

namespace GeoPerdasAPI.Controllers
{
    public class FrontApplicationController : Controller
    {
        public IActionResult Index()
        {
            return File("~/frontapp/index.html", "text/html"); // Altere o caminho do arquivo se necessário
        }
    }
}
