using System;
using System.Text;
using AODL.Document.TextDocuments;
using AODL.Document.Content.Tables;
using AODL.Document.Content;
using System.IO;
using System.Xml;

namespace UIK_writer_calc
{
    class Program
    {
        static void Main(string[] args)
        {
            //ExtractFromTable_odt(@"fixed_addr\barysh_2706.odt", 6, 3, 1, 4, 5);
            //ExtractFromTable_odt(@"fixed_addr\bazarnyy_2_version.odt", 6, 3, 1, 4, 5);
            //ExtractFromTable_odt(@"fixed_addr\veshkayma.odt", 6, 2, 1, 4, 5);
            //ExtractFromTable_odt(@"fixed_addr\inza0107.odt", 6, 2, 1, 4, 5);
            //ExtractFromTable_odt(@"karsun.odt", 6, 2, 1, 4, 5);
            //ExtractFromTable_odt(@"kuz_uchastki28.odt", 6, 2, 1, 4, 5);
            ExtractFromTable_odt(@"mayna_1.odt", 6, 2, 1, 4, 5);

            Console.WriteLine("Complite. Pres any key");
            Console.ReadKey();
        }
        
        static void ExtractFromTable_odt(string fileName, int tableColumns, int skipFirstRows, int columnUIK_id, int columnAddresOffice, int columnAddresVisit)
        {
            TextDocument dt = new TextDocument();
            dt.Load(fileName);
            Console.WriteLine();

            StreamWriter writerCsv = new StreamWriter(Path.ChangeExtension(fileName, "csv"));
            writerCsv.WriteLine(Extract.GetHeader());

            Table table = null;
            foreach (IContent content in dt.Content)
            {
                table = content as Table;
                if (table != null && table.ColumnCollection.Count == tableColumns) break;
            }

            if (table == null) return;

            int printRow = 0;

            foreach (Row row in table.Rows)
            {
                if (skipFirstRows > 0) { skipFirstRows--; continue; }
                if (row.CellSpanCollection.Count > 0) continue; // объединённые ячейки

                Extract ext = new Extract();

                ext.FillId(row.Cells[columnUIK_id].Node.InnerText);
                                
                // офис
                string test_extract = NodeToString(row.Cells[columnAddresOffice].Node);

                ext.FillAddress(test_extract, Extract.Place.office);

                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("++  " + test_extract);
                Console.WriteLine(">>  " + ext.UikToString());

                Console.ForegroundColor = ext.FullAddressOffice ? ConsoleColor.Yellow : ConsoleColor.Red;
                Console.WriteLine(">>  " + ext.OfficeAddrToString());                
                Console.ForegroundColor = ConsoleColor.Gray;

                Console.WriteLine(">>  " + ext.OfficePlaceToString());

                // место голосования

                test_extract = NodeToString(row.Cells[columnAddresVisit].Node);
                ext.FillAddress(test_extract, Extract.Place.visit);

                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("++  " + test_extract);

                Console.ForegroundColor = ext.FullAddressVisit ? ConsoleColor.Yellow : ConsoleColor.Red;
                Console.WriteLine(">>  " + ext.VisitAddrToString());
                Console.ForegroundColor = ConsoleColor.Gray;

                Console.WriteLine(">>  " + ext.VisitPlaceToString());
                
                Console.WriteLine();

                writerCsv.WriteLine(ext.GetRow());

                printRow++;
                if (printRow % 16 == 0) Console.ReadKey();
            }

            dt.Dispose();
            writerCsv.Flush();
            writerCsv.Close();            
        }

        static string NodeToString(XmlNode nodeCell)
        {
            string test_extract = string.Empty;
            // разделение текста через новые строчки, клеим их в одну строку
            if (nodeCell.ChildNodes.Count > 1)
            {
                StringBuilder builder_extract = new StringBuilder(512);
                foreach (XmlNode node in nodeCell.ChildNodes)
                {
                    builder_extract.Append(node.InnerText);
                    while (builder_extract[builder_extract.Length - 1] == ' '
                            || builder_extract[builder_extract.Length - 1] == ','
                            || builder_extract[builder_extract.Length - 1] == '.') builder_extract.Length--;
                    builder_extract.Append(",");
                }
                builder_extract.Length--;
                test_extract = builder_extract.ToString();
            } else
            {
                test_extract = nodeCell.InnerText;
            }
            return test_extract;
        }
    }
}
