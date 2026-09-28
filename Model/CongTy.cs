using System.ComponentModel.DataAnnotations.Schema;

namespace Model
{
    [Table("CongTy")]
    public class CongTy
    {
        public int Id { get; set; }
        public string MaCongTy { get; set; }
        public string TenCongTy { get; set; }
        public int? ParentId { get; set; }
        public int FlagDel { get; set; }
    }
}
