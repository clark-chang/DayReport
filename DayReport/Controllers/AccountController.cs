using DayReport.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Authentication;
using System.Diagnostics;

namespace DayReport.Controllers
{
    public class AccountController : Controller
    {
        //// GET: AccountController
        //public ActionResult Index()
        //{
        //    return View();
        //}

        //// GET: AccountController/Details/5
        //public ActionResult Details(int id)
        //{
        //    return View();
        //}

        //// GET: AccountController/Create
        //public ActionResult Create()
        //{
        //    return View();
        //}

        //// POST: AccountController/Create
        //[HttpPost]
        //[ValidateAntiForgeryToken]
        //public ActionResult Create(IFormCollection collection)
        //{
        //    try
        //    {
        //        return RedirectToAction(nameof(Index));
        //    }
        //    catch
        //    {
        //        return View();
        //    }
        //}

        //// GET: AccountController/Edit/5
        //public ActionResult Edit(int id)
        //{
        //    return View();
        //}

        //// POST: AccountController/Edit/5
        //[HttpPost]
        //[ValidateAntiForgeryToken]
        //public ActionResult Edit(int id, IFormCollection collection)
        //{
        //    try
        //    {
        //        return RedirectToAction(nameof(Index));
        //    }
        //    catch
        //    {
        //        return View();
        //    }
        //}

        //// GET: AccountController/Delete/5
        //public ActionResult Delete(int id)
        //{
        //    return View();
        //}

        //// POST: AccountController/Delete/5
        //[HttpPost]
        //[ValidateAntiForgeryToken]
        //public ActionResult Delete(int id, IFormCollection collection)
        //{
        //    try
        //    {
        //        return RedirectToAction(nameof(Index));
        //    }
        //    catch
        //    {
        //        return View();
        //    }
        //}

        
        
        // GET: AccountController/Login
        public IActionResult Login(string ReturnUrl="/")
        {
            LoginModel loginmodel = new LoginModel();
            loginmodel.ReturnUrl = ReturnUrl;
            return View(loginmodel);
        }

        // POST: AccountController/Login
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Login(LoginModel loginmodel)
        {
            if (ModelState.IsValid)
            {
                List<string> passgroup = new List<string> {"3732", "2886"};
                var check = passgroup.FirstOrDefault(x => x == loginmodel.Username);
                if(check != null)
                {
                    Login login = new Login();
                    int result = login.Validate(loginmodel.Username, loginmodel.Password);
                    
                    if(result == 1)
                    {
                        //1.Claims
                        var claims = new List<Claim>()
                        {
                            new Claim(ClaimTypes.NameIdentifier, loginmodel.Username),
                            new Claim(ClaimTypes.Name, loginmodel.Username),
                            new Claim(ClaimTypes.Role, "User")
                        };
                        //2.ClaimsIdentity
                        var identity = new ClaimsIdentity(claims, CookieAuthenticationDefaults.AuthenticationScheme);
                        //3.ClaimsPrincipal
                        var principal = new ClaimsPrincipal(identity);

                        await HttpContext.SignInAsync(CookieAuthenticationDefaults.AuthenticationScheme, principal);

                        //return RedirectToAction("Index", "Home");
                        return LocalRedirect("/DayReports");
                    }
                }
                else
                {
                    ViewBag.ErrorMessage = "請重新驗證";
                    return View(loginmodel);
                }
                //return RedirectToAction("Index", "Home");
            }
            ViewBag.ErrorMessage = "請重新驗證";
            return View(loginmodel);
        }

        public async Task<IActionResult> LogOut()
        {
            await HttpContext.SignOutAsync(CookieAuthenticationDefaults.AuthenticationScheme);
            return LocalRedirect("/DayReports");
        }
    }
}
