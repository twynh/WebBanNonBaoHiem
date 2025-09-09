using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace WebBanHangOnline.Models.EF
{
    [Table("tb_ProductCategory")]
    public class ProductCategory :CommonAbstract
    {
        
        public ProductCategory()
        {
            this.Products = new HashSet<Products>();
        }
        [Key]
        [DatabaseGeneratedAttribute(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }
        [Required]
        public string Title { get; set; }
        public string Alias { get; set; }
        public string SeoTitle { get; set; }
        public string SeoDescription { get; set; }
        public string SeoKeywords { get; set; }
        
        public string Description { get; set; }
        
        public string Icon { get; set; }
        public ICollection<Products> Products { get; set; }

    }
}