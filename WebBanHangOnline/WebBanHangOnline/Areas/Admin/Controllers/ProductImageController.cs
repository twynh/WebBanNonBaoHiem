using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using WebBanHangOnline.Models;
using WebBanHangOnline.Models.EF;

namespace WebBanHangOnline.Areas.Admin.Controllers
{
    public class ProductImageController : Controller
    {
        private ApplicationDbContext db = new ApplicationDbContext();
        // GET: Admin/ProductImage
        public ActionResult Index( int id)
        {
            ViewBag.ProductId = id;
            var items = db.ProductImgs.Where(x => x.ProductId == id).ToList();
            return View(items);
        }

        [HttpPost]
        public ActionResult AddImage(int productId , string url)
        {
            db.ProductImgs.Add(new ProductImgs
            {
                ProductId = productId,
                Image = url,
                isDefault = true
            });
            db.SaveChanges();
            return Json(new { Success = true });
        }

        [HttpPost]
        public ActionResult Delete(int id)
        {
            var item = db.ProductImgs.Find(id);
            db.ProductImgs.Remove(item);
            db.SaveChanges();
            return Json(new { success = true});
        }
    }
}