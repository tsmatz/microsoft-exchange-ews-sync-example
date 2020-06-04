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
        [DisplayName("OAuth Token")]
        public string OAuthToken { get; set; }
    }
}