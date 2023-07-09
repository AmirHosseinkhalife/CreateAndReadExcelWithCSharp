using System.ComponentModel.DataAnnotations;

namespace ReadExcelDemo
{
    public class DataModel
    {
        [Display(Name ="نام")]
        public string? FirstName { get; set; }
        public string? LastName { get; set; }
        public string? UserName { get; set; }
        [Display(Name ="کد ملی")]
        public string? NationalCode { get; set; }
        public string? PhoneNumber { get; set; }   
     
    }
}
