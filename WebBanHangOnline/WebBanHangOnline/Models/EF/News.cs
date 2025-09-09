using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace WebBanHangOnline.Models.EF
{
    [Table("tb_News")]
    public class News: CommonAbstract
    {
        [Key]
        [DatabaseGeneratedAttribute(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }
        [Required(ErrorMessage ="Bạn không được để trống mục này")]
        public string Title { get; set; }
        public string Description { get; set; }
        public string Alias { get; set; }

        public int CategoryID { get; set; }
        public string SeoTitle { get; set; }
        public string SeoDescription { get; set; }
        public string SeoKeywords { get; set; }
        [AllowHtml]
        public string Detail { get; set; }
        [AllowHtml]
        public string Image { get; set; }
        public bool isActive { get; set; }
        public virtual Category Category { get; set; }
       

    }
}