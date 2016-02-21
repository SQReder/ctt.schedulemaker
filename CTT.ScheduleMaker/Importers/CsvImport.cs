using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using CTT.ScheduleMaker.Model;

namespace CTT.ScheduleMaker.Importers
{
    internal class CsvImport
    {
        public async Task<ICollection<ScheduleItem>> Read(TextReader reader)
        {
            var line = await reader.ReadLineAsync();
            var tokenz = line.Split(';').Select(x => x.Trim()).ToList();
            return null;

        }
    }
}