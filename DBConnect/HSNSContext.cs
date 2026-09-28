using Model;
using System.Data.Entity;

namespace DBConnect
{
    public class HSNSContext : DbContext
    {
        public DbSet<CongTy> CongTy { get; set; }

        public HSNSContext() : base("HoSoNhanSu")
        {
        }
    }
}
