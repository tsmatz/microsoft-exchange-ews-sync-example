using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace ExchangeSyncSample.Models
{
    public class O365AccountModel
    {
        [Required]
        [DisplayName("メール アドレス")]
        public string MailAddress { get; set; }

        [Required]
        [DataType(DataType.Password)]
        [DisplayName("パスワード")]
        public string Password { get; set; }

        [Required]
        [DisplayName("URL")]
        public string Url { get; set; }
    }
}