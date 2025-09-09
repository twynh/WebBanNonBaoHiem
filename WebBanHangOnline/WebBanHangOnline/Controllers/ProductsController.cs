using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using WebBanHangOnline.Models;
using PagedList;
using PagedList.Mvc;
using WebBanHangOnline.Models.EF;

namespace WebBanHangOnline.Controllers
{
    public class ProductsController : Controller
    {
        private ApplicationDbContext db = new ApplicationDbContext();
        // GET: Products
        public ActionResult Index(int? page, string _name)
        {
           
            IEnumerable<Products> items = db.Products.OrderByDescending(x => x.Id);
            var pageSize = 8;
            if (page == null)
            {
                page = 1;
            }
            var pageIndex = page.HasValue ? Convert.ToInt32(page) : 1;
            items = items.ToPagedList(pageIndex, pageSize);
            ViewBag.pageSize = pageSize;
            ViewBag.Page = page;
            
            return View(items);

           
        }

        public ActionResult ProductCategory(int? page, string alias, int id)
        {

            IEnumerable<Products> itemss = db.Products.ToList();
            if (id > 0)
            {
                itemss = itemss.Where(x => x.ProductCategoryID == id).ToList();
            }
            var cate = db.ProductCategories.Find(id);
            if (cate != null)
            {
                ViewBag.CateName = cate.Title;
            }
            ViewBag.CateId = id;

            var pageSizee = 8;
            var pageIndexx = page.HasValue ? Convert.ToInt32(page) : 1;
            itemss = itemss.ToPagedList(pageIndexx, pageSizee);
            ViewBag.pageSize = pageSizee;
            ViewBag.Page = page;

            return View(itemss);
          
        }

        public ActionResult Partial_ItemsByCateId()
        {
            var items = db.Products.Where(x => x.isHome && x.isActive).ToList();
            return PartialView(items);
        }

        public ActionResult Detail(string alias, int id)
        {
            var items = db.Products.Find(id);
            return View(items);
        }

        public ActionResult Partial_BestSellerProduct()
        {
            var items = db.Products.Where(x => x.isHot && x.isActive).ToList();
            return PartialView(items);
        }
    }
}