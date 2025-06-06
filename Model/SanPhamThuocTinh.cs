using System;
using System.ComponentModel.DataAnnotations.Schema;

namespace Model
{
    [Table("SanPham_ThuocTinh")]
    public class SanPhamThuocTinh
    {
        public int Id { get; set; }
        public int SanPhamId { get; set; }
        public int ThuocTinhSPId { get; set; }
        public string NoiDung { get; set; }
        public int FlagDel { get; set; }
        public DateTime CreatedDate { get; set; }
        public DateTime UpdatedDate { get; set; }
    }
}
