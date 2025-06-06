using System;
using System.ComponentModel.DataAnnotations.Schema;

namespace Model
{
    [Table("DM_SanPham")]
    public class DMSanPham
    {
        public int Id { get; set; }
        public int? CongTyId { get; set; }
        public int? ParentId { get; set; }
        public string MaSPKH { get; set; }
        public string MaSPNB { get; set; }
        public string TenSanPham { get; set; }
        public string GhiChu { get; set; }
        public int HoatDong { get; set; }
        public int? NhomSanPhamId { get; set; }
        public int FlagDel { get; set; }
        public DateTime? CreatedDate { get; set; }
        public DateTime? UpdatedDate { get; set; }
    }
}
