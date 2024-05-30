using APIRandomUserGenerator.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;

namespace APIRandomUserGenerator.Controllers
{
    public class HomeController : Controller
    {
        //Chama index da pasta Views/Home
        public ActionResult Index()
        {
            return View();
        }
    }
}
