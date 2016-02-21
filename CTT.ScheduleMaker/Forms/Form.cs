using System;
using System.IO;
using System.Windows.Forms;
using CTT.ScheduleMaker.Importers;
using CTT.ScheduleMaker.Model;
using MoreLinq;
using Newtonsoft.Json;

namespace CTT.ScheduleMaker.Forms
{
    public partial class Form : System.Windows.Forms.Form
    {
        public Form()
        {
            InitializeComponent();
        }

        private async void button1_Click(object sender, System.EventArgs e)
        {
            var dataSet = XlsImport.ReadExcel("./app_data/schedule.xls");
            var collection = XlsImport.ReadFromDataset(dataSet, "расписание");
            collection.ForEach(x =>
            {
                Assert(() => !string.IsNullOrWhiteSpace(x.TimeBegin));
                Assert(() => !string.IsNullOrWhiteSpace(x.TimeEnd));
            });
            using (var writer = new StreamWriter("./App_Data/output.json"))
            {
                await writer.WriteLineAsync(JsonConvert.SerializeObject(collection, Formatting.Indented));
            }
            button1.Text = @"Done";
        }

        private void Assert(Func<bool> func, string message = null)
        {
            if (!func())
                throw new Exception(message);
        }
    }
}
