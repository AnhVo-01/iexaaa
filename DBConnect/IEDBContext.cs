using Model;
using System.Data.Entity;

namespace DBConnect
{
    public class IEDBContext : DbContext
    {
        public DbSet<DMSanPham> DMSanPham { get; set; }
        public DbSet<DMThuocTinh> DMThuocTinh { get; set; }
        public DbSet<SanPhamThuocTinh> SanPhamThuocTinh { get; set; }

        public IEDBContext() : base("Sales")
        {
            string a = "";
        }
    }
}
