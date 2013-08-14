using System;
using System.Text;
using AODL.Document.TextDocuments;
using AODL.Document.Content.Tables;
using AODL.Document.Content;
using System.IO;
using System.Xml;
using AODL.Document.Content.Text;

namespace UIK_writer_calc
{
    class Program
    {
        static void Main(string[] args)
        {
            //ExtractFromTable_odt(@"barysh_2706.odt", 6, 3, 1, 4, 5);
            //ExtractFromTable_odt(@"bazarnyy_2_version.odt", 6, 2, 1, 4, 5);
            //ExtractFromTable_odt(@"veshkayma.odt", 6, 2, 1, 4, 5);
            //ExtractFromTable_odt(@"inza0107.odt", 6, 2, 1, 4, 5);
            //ExtractFromTable_odt(@"karsun.odt", 6, 1, 1, 4, 5);
            //ExtractFromTable_odt(@"kuz_uchastki28.odt", 6, 2, 1, 4, 5);
            //ExtractFromTable_odt(@"mayna_1.odt", 6, 0, 1, 4, 5);
            //ExtractFromTable_odt(@"nikolaevka_2_version.odt", 6, 2, 1, 4, 5);
            //ExtractFromTable_odt(@"novomal.odt", 6, 2, 1, 4, 5);
            //ExtractFromTable_odt(@"pavlovka_0107.odt", 6, 1, 1, 4, 5);
            //ExtractFromTable_odt(@"novospasskiy.odt", 6, 2, 1, 4, 5);
            //ExtractFromTable_odt(@"radischevo.odt", 6, 2, 1, 4, 5);
            //ExtractFromTable_odt(@"sengileevskiy.odt", 6, 1, 1, 4, 5);
            //ExtractFromTable_odt(@"kulatka.odt", 6, 2, 1, 4, 5);
            //ExtractFromTable_odt(@"st_uch28.odt", 6, 1, 1, 4, 5);
            //ExtractFromTable_odt(@"surskiy.odt", 6, 1, 1, 4, 5);
            //ExtractFromTable_odt(@"terenga.odt", 6, 1, 1, 4, 5);
            //ExtractFromTable_odt(@"ulbyanovskiy_rayon.odt", 6, 1, 1, 4, 5);
            //ExtractFromTable_odt(@"cylna.odt", 8, 2, 1, 6, 7);
            //ExtractFromTable_odt(@"tcherdaklinskiy_1008_3_version.odt", 6, 1, 1, 4, 5);
            //ExtractFromText_odt(@"1_zd.odt", "город Ульяновск");    // именно городской округ, т.к. адреса идут по улица в деревнях округа
            //ExtractFromTable_odt(@"dimitrov_11_07.odt", 6, 1, 1, 4, 5);     // есть уникальный код в NodeToString (адрес в последней строчке)
            //ExtractFromTable_odt(@"novoulbyanovsk.odt", 6, 2, 1, 4, 5);     // есть уникальный код в NodeToString и ExtractFromTable_odt (адрес в предпоследней строчке)
            //ExtractFromText_odt(@"4_len.odt", "город Ульяновск");    // именно городской округ, т.к. адреса идут по улица в деревнях округа
            //ExtractFromText_odt(@"2_zv.odt", "город Ульяновск");    // именно городской округ, т.к. адреса идут по улица в деревнях округа
            //ExtractFromText_odt(@"3_zsv.odt", "город Ульяновск");    // именно городской округ, т.к. адреса идут по улица в деревнях округа
            ExtractFromTable_odt(@"melekess.odt", 6, 2, 1, 4, 5);

            Console.WriteLine("Complite. Pres any key");
            Console.ReadKey();
        }

        static string foundUIK = "Избирательный участок №";
        static int foundUIK_len = foundUIK.Length;

        static string found_addr_o = "Место нахождения УИК ";
        static int found_addr_o_len = found_addr_o.Length;

        static string found_place_v = "Помещение для голосования ";
        static int found_place_v_len = found_place_v.Length;

        static void ExtractFromText_odt(string fileName, string addr_town)
        {
            TextDocument dt = new TextDocument();
            dt.Load(fileName);
            Console.WriteLine();

            ExportAndLog export_log = new ExportAndLog(Path.ChangeExtension(fileName, "csv"));

            Paragraph paragraph = null;
            Extract ext = null;
            foreach (IContent content in dt.Content)
            {
                paragraph = content as Paragraph;
                if (paragraph != null)
                {
                    string test_text = paragraph.Node.InnerText;

                    if (test_text.StartsWith(foundUIK))
                    {
                        if (ext != null)
                        {
                            export_log.LogOffice(ext);
                            export_log.LogVisit(ext);
                            export_log.Export(ext);
                            Console.WriteLine();
                        }

                        ext = new Extract();

                        int foundNum = foundUIK_len + 1;
                        while (test_text[foundNum] == ' ') { foundNum++; }
                        ext.FillId(test_text.Substring(foundNum));
                        continue;
                    }

                    if (ext != null && ext.HaveUikId && test_text.StartsWith(found_addr_o))
                    {
                        int foundAddr = found_addr_o_len + 1;
                        while (test_text[foundAddr] == ' ') { foundAddr++; }
                        test_text = test_text.Substring(foundAddr);

                        test_text = ExtractTextInbrackets(test_text);

                        ext.FillAddress(test_text, Extract.Place.office);
                        if (ext.SomethingAddressOffice && !ext.ExistsAddressPlaceOffice) ext.SetAddrPlace(addr_town, Extract.Place.office);
                    }

                    if (ext != null && ext.HaveUikId && test_text.StartsWith(found_place_v))
                    {
                        int foundPlace = found_place_v_len + 1;
                        while (test_text[foundPlace] == ' ') { foundPlace++; }
                        test_text = test_text.Substring(foundPlace);

                        test_text = ExtractTextInbrackets(test_text);

                        ext.FillAddress(test_text, Extract.Place.visit);
                        if (ext.SomethingAddressVisit && !ext.ExistsAddressPlaceVisit) ext.SetAddrPlace(addr_town, Extract.Place.visit);
                    }

                }
            }

            if (ext != null)
            {
                export_log.LogOffice(ext);
                export_log.LogVisit(ext);
                export_log.Export(ext);
                Console.WriteLine();
            }
            dt.Dispose();
            export_log.Dispose();
        }
        
        static void ExtractFromTable_odt(string fileName, int tableColumns, int skipFirstRows, int columnUIK_id, int columnAddresOffice, int columnAddresVisit)
        {
            TextDocument dt = new TextDocument();
            dt.Load(fileName);
            Console.WriteLine();

            ExportAndLog export_log = new ExportAndLog(Path.ChangeExtension(fileName, "csv"));

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
                //if (!ext.FullAddressOffice && ext.SomethingAddressOffice) ext.SetAddrPlace("г. Новоульяновк", Extract.Place.office);    // новоульяновск
                export_log.LogOffice(ext);

                // место голосования

                test_extract = NodeToString(row.Cells[columnAddresVisit].Node);
                ext.FillAddress(test_extract, Extract.Place.visit);
                //if (!ext.FullAddressVisit && ext.SomethingAddressVisit) ext.SetAddrPlace("г. Новоульяновк", Extract.Place.visit);       // новоульяновск
                Console.WriteLine();

                export_log.Export(ext);

                printRow++;
                if (printRow % 8 == 0) Console.ReadKey();
            }

            dt.Dispose();
            export_log.Dispose();
        }

        static StringBuilder builder_extract = new StringBuilder(512);
        static string NodeToString(XmlNode nodeCell)
        {
            string test_extract = string.Empty;
            //int dimitrovgrad_last_pi = 0;                               // димитровград
            //int novul_last_pi = 0;                                              // новоульяновск
            //int novul_prelast_pi = 0;                                           // новоульяновск
            // разделение текста через новые строчки, клеим их в одну строку
            if (nodeCell.ChildNodes.Count > 1)
            {
                builder_extract.Clear();
                foreach (XmlNode node in nodeCell.ChildNodes)
                {
                    //if (node.InnerText.Length > 0) dimitrovgrad_last_pi = builder_extract.Length;      // димитровград
                    //if (node.InnerText.Length > 0 && novul_last_pi != 0) novul_prelast_pi = novul_last_pi;              // новоульяновск
                    //if (node.InnerText.Length > 0) novul_last_pi = builder_extract.Length;                              // новоульяновск

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

            // для Димитровград: добавить город в начало, адрес лежит в последней строчке ячейки
            /*int len = builder_extract.Length - dimitrovgrad_last_pi;
            builder_extract.Insert(0, builder_extract.ToString().Substring(dimitrovgrad_last_pi));
            builder_extract.Insert(len, ',');
            builder_extract.Length -= len +1;
            builder_extract.Insert(0, "г. Димитровград,");
            test_extract = builder_extract.ToString();*/
            // конец Димитровград

            // для Новоульяновк: добавить город в начало, адрес лежит в предпоследней строчке ячейки
            /*int len = novul_last_pi - novul_prelast_pi;
            builder_extract.Insert(0, builder_extract.ToString().Substring(novul_prelast_pi, len));
            builder_extract.Remove(novul_prelast_pi + len, len);
            test_extract = builder_extract.ToString();*/
            // конец Новоульяновк

            return test_extract;
        }

        static string ExtractTextInbrackets(string test_text)
        {
            int open = test_text.LastIndexOf('(');
            int close = test_text.LastIndexOf(')');
            if (open == -1 || close == -1) return test_text;

            builder_extract.Clear();
            builder_extract.Append(test_text.Substring(open + 1, close - open - 1)).Append(',');
            builder_extract.Append(test_text.Substring(0, open - 1));
            builder_extract.Append(test_text.Substring(close + 1, test_text.Length - close - 1));
            return builder_extract.ToString();
        }
    }
}
