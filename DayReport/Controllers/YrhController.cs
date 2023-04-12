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
    public class YrhController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Yrhho()
        {
            string datetime = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
            int hour = Convert.ToInt32(datetime.Substring(11, 2));
            Yrh yrh = new Yrh();
            DataSet ds = new DataSet();
            DataTable todayinhospital = yrh.Hotodayin();
            DataTable todayinhospitalo = yrh.HotodayinO();
            DataTable todayinhospitalv = yrh.HotodayinV();
            DataTable Gettodayinhospitalsum = yrh.Hotodaysum();
            ds.Tables.Add(todayinhospital);
            ds.Tables.Add(todayinhospitalo);
            ds.Tables.Add(todayinhospitalv);
            ds.Tables.Add(Gettodayinhospitalsum);
            return View(ds);
        }

        public IActionResult Yrhhi()
        {
            Yrh yrh = new Yrh();
            DataTable todayin = yrh.Hitodayin();
            DataTable todayout = yrh.Hitodayout();
            DataTable now = yrh.Hinow();
            DataTable peopleinursestation1 = yrh.Nursestation1();
            DataTable peopleinursestation2 = yrh.Nursestation2();
            DataTable peopleinursestation3 = yrh.Nursestation3();
            DataTable peopleinursestation4 = yrh.Nursestation4();
            DataTable peopleinursestation5 = yrh.Nursestation5();
            DataTable todayicu = yrh.Icunow();
            DataTable peopleinursestation6 = yrh.Nursestation6();


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

        public IActionResult Yrhem()
        {
            Yrh yrh = new Yrh();
            DataSet ds = new DataSet();
            DataTable emtoday = yrh.Emtoday();
            DataTable emtodayfirst = yrh.Emtodayfirst();
            DataTable emtodaytransfer = yrh.Emtodaytransfer();
            DataTable emtotal = yrh.Emtodaysum();
            ds.Tables.Add(emtoday);
            ds.Tables.Add(emtodayfirst);
            ds.Tables.Add(emtodaytransfer);
            ds.Tables.Add(emtotal);
            return View(ds);
        }
    }
}
