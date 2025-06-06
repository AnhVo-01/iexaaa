using System;
using System.ComponentModel.DataAnnotations.Schema;

namespace Model
{
    [Table("DM_ThuocTinhSP")]
    public class DMThuocTinh
    {
        public int Id { get; set; }
        public string TenThuocTinh { get; set; }
        public string MaThuocTinh { get; set; }
        public int? CongTyId { get; set; }
        public string DanhSachLuaChon { get; set; }
        public int BatBuoc { get; set; }
        public string MoTa { get; set; }
        public int? TypeDataId { get; set; }
        public int FlagDel { get; set; }
        public int TrangThai { get; set; }
        public DateTime? CreatedDate { get; set; }
        public DateTime? UpdatedDate { get; set; }
    }
}
