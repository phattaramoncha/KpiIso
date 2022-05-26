using System;
using System.ComponentModel.DataAnnotations;

namespace KpiISO.Data.Model
{
    public class Login
    {
        [Required(ErrorMessage = "กรุณาระบุ ชื่อผู้ใช้งาน !")]
        [Display(Name = "ชื่อผู้ใช้")]
        public string UserName { get; set; }

        [Required(ErrorMessage = "กรุณาระบุ รหัสผ่าน !")]
        [Display(Name = "รหัสผ่าน")]
        public string Password { get; set; }
    }
    public class LoginConfirm
    {
        [Required(ErrorMessage = "กรุณาระบุ ชื่อผู้ใช้งาน !")]
        [Display(Name = "ชื่อผู้ใช้")]
        //[Remote("IsValidUserLogin", "Common", ErrorMessage = "ชื่อผู้ใช้ หรือ รหัสผ่าน ไม่ถูกต้อง กรุณากรอกใหม่ !", AdditionalFields = "Password")]
        public string UserName { get; set; }

        [Required(ErrorMessage = "กรุณาระบุ รหัสผ่าน !")]
        [Display(Name = "รหัสผ่าน")]
        public string Password { get; set; }
        //[DefaultValue(false)]
        public bool? IsValid { get; set; }
    }
    public class UserInfo
    {
        public string UserId { get; set; }
        public string UserName { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Email { get; set; }
        public string ProjectId { get; set; }
        public string ProjectName { get; set; }
        public bool IsProjectSelected { get; set; }
        public Int32 IsManager { get; set; }
        public Int32 IsEmp { get; set; }
        public Int32 IsAdmin { get; set; }
        public Int32 IsAds { get; set; }
        public string DisplayName { get; set; }
    }
}