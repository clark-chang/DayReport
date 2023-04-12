using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Data;
using DayReport.Models;
using Microsoft.AspNetCore.Authorization;

namespace DayReport.Controllers
{
   [Authorize]
    public class YshController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Yshho()
        {
            string datetime = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
            int hour = Convert.ToInt32(datetime.Substring(11, 2));
            Ysh ysh = new Ysh();
            DataSet ds = new DataSet();
            DataTable todayinhospital = ysh.Hotodayin();
            DataTable todayinhospitalo = ysh.HotodayinO();
            DataTable todayinhospitalv = ysh.HotodayinV();
            DataTable Gettodayinhospitalsum = ysh.Hotodaysum();
            ds.Tables.Add(todayinhospital);
            ds.Tables.Add(todayinhospitalo);
            ds.Tables.Add(todayinhospitalv);
            ds.Tables.Add(Gettodayinhospitalsum);
            return View(ds);
        }

        public IActionResult Yshhi()
        {
            Ysh ysh = new Ysh();
            DataTable todayin = ysh.Hitodayin();
            DataTable todayout = ysh.Hitodayout();
            DataTable now = ysh.Hinow();
            DataTable peopleinursestation1 = ysh.Nursestation1();
            DataTable peopleinursestation2 = ysh.Nursestation2();
            DataTable peopleinursestation3 = ysh.Nursestation3();
            DataTable peopleinursestation4 = ysh.Nursestation4();
            DataTable peopleinursestation5 = ysh.Nursestation5();
            DataTable todayicu = ysh.Icunow();
            DataTable peopleinursestation6 = ysh.Nursestation6();


            DataSet ds = new DataSet();
            ds.Tables.Add(todayin);
            ds.Tables.Add(todayout);
            ds.Tables.Add(now);
            ds.Tables.Add(peopleinursestation1);
            ds.Tables.Add(peopleinursestation2);
            ds.Tables.Add(peopleinursestation3);
            ds.Tables.Add(peopleinursestation4);
            ds.Tables.Add(peopleinursestation5);
            ds.Tables.Add(todayicu);
            ds.Tables.Add(peopleinursestation6);
            return View(ds);
        }

        public IActionResult Yshem()
        {
            Ysh ysh = new Ysh();
            DataSet ds = new DataSet();
            DataTable emtoday = ysh.Emtoday();
            DataTable emtodayfirst = ysh.Emtodayfirst();
            DataTable emtodaytransfer = ysh.Emtodaytransfer();
            DataTable emtotal = ysh.Emtodaysum();
            ds.Tables.Add(emtoday);
            ds.Tables.Add(emtodayfirst);
            ds.Tables.Add(emtodaytransfer);
            ds.Tables.Add(emtotal);
            return View(ds);
        }
    }
}
