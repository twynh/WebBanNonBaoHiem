﻿using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace WebBanHangOnline.Models.EF
{
    [Table("tb_Order")]
    public class Order: CommonAbstract
    {
        public Order()
        {
            this.OrderDetails = new HashSet<OrderDetail>();
        }
        [Key]
        [DatabaseGeneratedAttribute(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }
        [Required]
        public string Code { get; set; }
        [Required(ErrorMessage ="Không được để trống")]
        public string CustumerName { get; set; }
        [Required(ErrorMessage = "Không được để trống")]
        public string Phone { get; set; }
        [Required(ErrorMessage = "Không được để trống")]
        public string Address { get; set; }
        
        public string Email { get; set; }
        public double TotalAmount { get; set; }
        public int Quantity { get; set; }
        public int TypePayment { get; set; }
        public virtual ICollection<OrderDetail> OrderDetails { get; set; }
    }
}