using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace ExcelReader
{
    public class Workbook
    {
        public static SharedStringRoot SharedStrings;

        /// <summary>
        /// All worksheets in the Excel workbook deserialized
        /// </summary>
        /// <param name="ExcelFileName">Full path and filename of the Excel xlsx-file</param>
        /// <returns></returns>
        public static IEnumerable<Worksheet> Worksheets(string ExcelFileName)
        {
            Worksheet ws;
            FileStream fs = new FileStream(ExcelFileName, FileMode.OpenOrCreate, FileAccess.Read);

            if (fs.CanRead)
            {

                using (ZipArchive zipArchive = ZipFile.Open(ExcelFileName, ZipArchiveMode.Read))
                {
                    SharedStrings =
                        DeserializedZipEntry<SharedStringRoot>(GetZipArchiveEntry(zipArchive, @"xl/sharedStrings.xml"));
                    foreach (var worksheetEntry in (WorkSheetFileNames(zipArchive)).OrderBy(x => x.FullName))
                    {
                        ws = DeserializedZipEntry<Worksheet>(worksheetEntry);
                        ws.NumberOfColumns = Worksheet.MaxColumnIndex + 1;
                        ws.ExpandRows();
                        yield return ws;
                    }
                }
            }
            else
            {
                throw new FileLoadException("Access denied!");
            }
        }

        /// <summary>
        /// Method converting an Excel cell value to a date
        /// </summary>
        /// <param name="ExcelCellValue"></param>
        /// <returns></returns>
        public static DateTime DateFromExcelFormat(string ExcelCellValue)
        {
            return DateTime.FromOADate(Convert.ToDouble(ExcelCellValue));
        }

        private static ZipArchiveEntry GetZipArchiveEntry(ZipArchive ZipArchive, string ZipEntryName)
        {
            return ZipArchive.Entries.First<ZipArchiveEntry>(n => n.FullName.Equals(ZipEntryName));
        }
        private static IEnumerable<ZipArchiveEntry> WorkSheetFileNames(ZipArchive ZipArchive)
        {
            foreach (var zipEntry in ZipArchive.Entries)
                if (zipEntry.FullName.StartsWith("xl/worksheets/sheet"))
                    yield return zipEntry;
        }
        private static T DeserializedZipEntry<T>(ZipArchiveEntry ZipArchiveEntry)
        {
            using (Stream stream = ZipArchiveEntry.Open())
                return (T)new XmlSerializer(typeof(T)).Deserialize(XmlReader.Create(stream));
        }
    }
}
