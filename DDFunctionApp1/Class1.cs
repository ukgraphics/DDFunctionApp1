using System;
using System.Collections.Generic;
using System.Text;

namespace DDFunctionApp1
{

    public class Data
    {
        public Publisher publisher { get; set; }
        public Customer customer { get; set; }
    }

    //発行元情報
    public class Publisher
    {
        public string companyname { get; set; }
        public string postalcode { get; set; }
        public string address1 { get; set; }
        public string address2 { get; set; }
        public string tel { get; set; }
        public string bankname { get; set; }
        public string bankblanch { get; set; }
        public string account { get; set; }
        public string representative { get; set; }
    }

    //顧客情報
    public class Customer
    {
        public string companyname { get; set; }
        public string name { get; set; }
        public string postalcode { get; set; }
        public string address1 { get; set; }
        public string address2 { get; set; }
        public string tel { get; set; }
        public Detail[] detail { get; set; }
    }

    //明細
    public class Detail
    {
        public string sku { get; set; }
        public string name { get; set; }
        public int price { get; set; }
        public int unit { get; set; }
        public string remark { get; set; }
    }

}
