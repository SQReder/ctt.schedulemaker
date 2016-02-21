using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using CTT.ScheduleMaker.Model;
using MoreLinq;
using Newtonsoft.Json;

namespace CTT.ScheduleMaker
{
    internal static class JsonExport
    {
        public static async Task ExportAsync(TextWriter writer, IEnumerable<ScheduleItem> items)
        {
            foreach (var s in items.Select(JsonConvert.SerializeObject))
            {
                await writer.WriteLineAsync(s);
            }
        }
    }
}