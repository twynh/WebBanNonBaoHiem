using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace WebBanHangOnline.Models.EF
{
    [Table("tb_Category")]
    public class Category : CommonAbstract
    {
        public Category()
        {
            this.News = new HashSet<News>(); 
        }
        [Key]
        [DatabaseGeneratedAttribute(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        [Required(ErrorMessage ="Tên danh mục không được để trống")]
        public string Title { get; set; }
        public string Alias { get; set; }
        //public string TypeCode { get; set; }
        //public string Link { get; set; }


        public string Description { get; set; } 
        public int Position { get; set; }
        public string SeoTitle { get; set; }
        public bool isActive { get; set; }
        public string SeoDescription { get; set; }
        public string SeoKeywords { get; set; }
        public ICollection<News> News { get; set; }
        public ICollection<Post> Posts { get; set; }
       
       
    }
}