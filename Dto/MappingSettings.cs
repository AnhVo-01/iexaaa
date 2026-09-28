using System.Collections.Generic;

namespace Dto
{
    public class MappingSettings
    {
        public int? CongTyId { get; set; }
        public List<ColumnMapping> Mappings { get; set; } = new List<ColumnMapping>();
    }
}
