using System;
using System.IO;
using System.Text;

namespace UIK_writer_calc
{
    public class ExportAndLog : IDisposable
    {
        private StreamWriter writerCsv;
        public ExportAndLog(string fileName)
        {
            writerCsv = new StreamWriter(Path.ChangeExtension(fileName, "csv"));
            writerCsv.WriteLine(Extract.GetHeader());
        }

        public void Dispose()
        {
            writerCsv.Flush();
            writerCsv.Close();
            writerCsv.Dispose();
        }

        public void LogOffice(Extract ext)
        {
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("++  " + ext.test_office);
            Console.WriteLine(">>  " + ext.UikToString());

            Console.ForegroundColor = ext.FullAddressOffice ? ConsoleColor.Yellow : ConsoleColor.Red;
            Console.WriteLine(">>  " + ext.OfficeAddrToString());
            Console.ForegroundColor = ConsoleColor.Gray;

            Console.WriteLine(">>  " + ext.OfficePlaceToString());
        }

        public void LogVisit(Extract ext)
        {
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("++  " + ext.test_visit);

            Console.ForegroundColor = ext.FullAddressVisit ? ConsoleColor.Yellow : ConsoleColor.Red;
            Console.WriteLine(">>  " + ext.VisitAddrToString());
            Console.ForegroundColor = ConsoleColor.Gray;

            Console.WriteLine(">>  " + ext.VisitPlaceToString());
        }

        public void Export(Extract ext)
        {
            writerCsv.WriteLine(ext.GetRow());
        }
    }
}
