using System.Collections.Generic;

namespace ExcelTrainingMonitor.Models
{
    internal sealed class ProcessRecordMetadata
    {
        public string SupplierFarmName { get; set; } = "";
        public int BirdsProcessed { get; set; }
        public Dictionary<string, List<string>> DropdownLists { get; set; } =
            new Dictionary<string, List<string>>(System.StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, string> CellDropdownAssignments { get; set; } =
            new Dictionary<string, string>();
        public Dictionary<string, string> DailyProductionCellDropdownAssignments { get; set; } =
            new Dictionary<string, string>();
    }
}
