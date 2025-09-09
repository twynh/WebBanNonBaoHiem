using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace WebBanHangOnline.Models
{
    public class OrderViewModel
    {
        [Required(ErrorMessage = "Không được để trống")]
        public string CustumerName { get; set; }
        [Required(ErrorMessage = "Không được để trống")]
        public string Phone { get; set; }
        [Required(ErrorMessage = "Không được để trống")]
        public string Address { get; set; }
       
        public string Email { get; set; }
        public int TypePayment { get; set; }
    }
}