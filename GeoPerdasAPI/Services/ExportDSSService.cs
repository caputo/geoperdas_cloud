using GeoPerdasCloud.ProgGeoPerdas.Legacy.Config;
using GeoPerdasCloud.ProgGeoPerdas.Legacy.LegacyCode;
using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;

namespace GeoPerdasAPI.Services
{
    public class ExportDSSService
    {
        public static MemoryStream ExportDss(FormConfigControls config) {
            var legacy = new ProgGeoperdasForm(config);
            legacy.bConnection_Click();
            var content = legacy.bExportDB_Click();

            using MemoryStream memoryStream = new MemoryStream();
            using (ZipArchive arquivoZip = new ZipArchive(memoryStream, ZipArchiveMode.Create, true))
            {

                foreach (var fileName in content.Keys)
                {
                    ZipArchiveEntry entrada = arquivoZip.CreateEntry(fileName.Replace("\\",""), CompressionLevel.Optimal);
                    using (StreamWriter writer = new StreamWriter(entrada.Open()))
                    {
                        foreach (var line in content[fileName])
                        {
                            writer.WriteLine(line);
                        }
                    }

                }
            }            
            memoryStream.Seek(0, SeekOrigin.Begin);
            return memoryStream;
        }
    }
}
