using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace WebBanHangOnline.Models.EF
{
    [Table("tb_Product")]
    public class Products: CommonAbstract
    {
        public Products()
        {
            this.ProductImage = new HashSet<ProductImgs>();
            this.OrderDetails = new HashSet<OrderDetail>();
        }
        [Key]
        [DatabaseGeneratedAttribute(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }
        [Required]
        public string Title { get; set; }
        public string Description { get; set; }
        public string Alias { get; set; }

        public string ProductCode { get; set; }
        public int ProductCategoryID { get; set; }
        public string SeoTitle { get; set; }
        public string SeoDescription { get; set; }
        public string SeoKeywords { get; set; }
        [AllowHtml]
        public string Detail { get; set; }
        public string Image { get; set; }
        public double OriginalPrice { get; set; }
        public double Price { get; set; }
    
       
        public bool isActive { get; set; }
        public double? PriceSale { get; set; }
        public bool isHome { get; set; }
        public bool isFeature { get; set; }
        public bool isHot { get; set; }
        public bool isSale { get; set; }
        public int Quantity { get; set; }
        public virtual ProductCategory ProductCategory { get; set; }
        public virtual ICollection<ProductImgs> ProductImage { get; set; }
        public virtual ICollection<OrderDetail> OrderDetails { get; set; }

    }
}