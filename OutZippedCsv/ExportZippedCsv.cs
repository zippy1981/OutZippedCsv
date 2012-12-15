using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Management.Automation;

using Ionic.Zip;
using Kent.Boogaart.KBCsv;

namespace OutZippedCsv
{
    [Cmdlet(VerbsData.Export, "ZippedCsv")]
    public class ExportZippedCsv : Cmdlet
    {
        [Parameter(Mandatory = true)]
        public string FileName { get; set; }

        [Parameter(ValueFromPipeline = true, Mandatory = true)]
        public DataSet Data { get; set; }

        protected override void ProcessRecord()
        {
            FileName = Path.GetFullPath(FileName);
            var basePath = Path.GetDirectoryName(FileName);
            var csvName = string.Format("{0}.csv", Path.GetFileNameWithoutExtension(FileName));

            if (File.Exists(FileName)) { throw new ArgumentException(string.Format("File {0} exists!", FileName));}
            if (!Directory.Exists(basePath)) { throw new DirectoryNotFoundException(string.Format("Directory {0} does not exist.", basePath));}
            
            using (var zipStream = new ZipFile(basePath))
            using (var csvRaw = new MemoryStream())
            using (var csvWriter = new CsvWriter(csvRaw))
            {
                csvWriter.WriteAll(Data);
                csvRaw.Seek(0, SeekOrigin.Begin);
                zipStream.AddEntry(csvName, csvRaw);
            }
        }
    }
}
