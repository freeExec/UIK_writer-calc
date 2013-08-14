using System;
using System.Text;

namespace UIK_writer_calc
{
    public class Extract
    {
        public enum Place
        {
            visit,
            office
        };

        private int uik_id;

        private string place_addr_v;
        private string street_addr_v;
        private string building_addr_v;

        private string place_addr_o;
        private string street_addr_o;
        private string building_addr_o;

        private string place_place_v;
        private string place_place_o;

        private string phone_v;
        private string phone_o;

        public string test_office;
        public string test_visit;

        private static string[] place_prefix = { "г.", "п.", "р.п.", "с.", "пос.", "д.", "р-д", "станция", "село", "ст.", "город" };
        private static string[] street_prefix = { "ул.", "пл.", "пер.", "переулок", "площадь", "улица", "пр-т", "пр." };
        private static string[] building_prefix = { "д.", "дом" };
        private static string[] phone_prefix = { "т.", "тел.:", "тел.", "телефон:", "тел:" };

        private StringBuilder resultToString;

        private static string csv_header = "uik;addr_v;place_v;phone_v;addr_o;place_o;phone_o;comment;g_status";

        public Extract()
        {
            resultToString = new StringBuilder(255);
        }

        public bool FullAddressOffice
        {
            get { return place_addr_o != null && street_addr_o != null && building_addr_o != null; }
        }

        public bool FullAddressVisit
        {
            get { return place_addr_v != null && street_addr_v != null && building_addr_v != null; }
        }

        public bool SomethingAddressOffice
        {
            get { return place_addr_o != null || street_addr_o != null || building_addr_o != null; }
        }

        public bool SomethingAddressVisit
        {
            get { return place_addr_v != null || street_addr_v != null || building_addr_v != null; }
        }

        public bool HaveUikId
        {
            get { return uik_id != 0; }
        }

        public void FillAddress(string text, Place place)
        {
            if (place == Place.visit) test_visit = text; else test_office = text;

            string[] splits = text.Split(',');

            bool processed_split = false;
            int lastRecogSplit = -1;
            for (int spl = 0; spl < splits.Length; spl++)
            {
                string test_split = splits[spl].TrimStart(' ');
                processed_split = false;

                // населенные пункт
                if ((place_addr_o == null && place == Place.office || place_addr_v == null && place == Place.visit)
                    //&& test_split.Length > 5)   // небольшая защита от дома (д.) в том случае если указана только улица и номер дома
                    && (street_addr_o == null && place == Place.office || street_addr_v == null && place == Place.visit))
                {
                    for (int pref = 0; pref < place_prefix.Length; pref++)
                    {
                        if (test_split.StartsWith(place_prefix[pref]))
                        {
                            if (place == Place.visit) place_addr_v = test_split; else place_addr_o = test_split;
                            processed_split = true;
                            lastRecogSplit = spl;
                            break;
                        }
                    }
                    if (processed_split) continue;
                }

                // улица
                if (street_addr_o == null && place == Place.office || street_addr_v == null && place == Place.visit)
                {
                    for (int pref = 0; pref < street_prefix.Length; pref++)
                    {
                        if (test_split.StartsWith(street_prefix[pref]))
                        {
                            if (place == Place.visit) street_addr_v = test_split; else street_addr_o = test_split;
                            processed_split = true;
                            lastRecogSplit = spl;
                            break;
                        }
                    }

                    // если между улицой и дома нет разделителя, не работает если буква в адресе
                    if (processed_split && test_split.Length > 4)
                    {
                        int last_char = test_split.Length - 1;
                        while (test_split[last_char] >= '0' && test_split[last_char] <= '9') { last_char--; }
                        if (last_char != test_split.Length - 1)
                        {
                            string build_addr = test_split.Substring(last_char).TrimStart(' ');
                            if (place == Place.visit)
                            {
                                building_addr_v = build_addr;
                                street_addr_v = street_addr_v.Remove(last_char);
                            } else
                            {
                                building_addr_o = build_addr;
                                street_addr_o = street_addr_o.Remove(last_char);
                            }

                        }
                    }
                    if (processed_split) continue;
                }

                // строение
                if (building_addr_o == null && place == Place.office || building_addr_v == null && place == Place.visit)
                {
                    for (int pref = 0; pref < building_prefix.Length; pref++)
                    {
                        if (test_split.StartsWith(building_prefix[pref]))
                        {
                            if (place == Place.visit) building_addr_v = test_split; else building_addr_o = test_split;
                            processed_split = true;
                            lastRecogSplit = spl;
                            break;
                        }
                    }
                    // попытка пофиксить багу отсутствия префикса дом.
                    if (test_split.Length > 0
                        //&& (building_addr_o == null && place == Place.office || building_addr_v == null && place == Place.visit)
                        && lastRecogSplit != -1     // было найдено или НП или улица
                        && test_split[0] >= '0' && test_split[0] <= '9')
                    {
                        if (place == Place.visit) building_addr_v = test_split; else building_addr_o = test_split;
                        processed_split = true;
                        lastRecogSplit = spl;
                    }
                    if (processed_split) continue;
                }
                break;
            }

            resultToString.Clear();
            for (int lastSplits = lastRecogSplit + 1; lastSplits < splits.Length; lastSplits++)
            {
                processed_split = false;
                string test_split = splits[lastSplits].TrimStart(' ');
                // телефон
                for (int pref = 0; pref < phone_prefix.Length; pref++)
                {
                    if (test_split.StartsWith(phone_prefix[pref]))
                    {
                        if (place == Place.visit) phone_v = test_split; else phone_o = test_split;
                        processed_split = true;
                        break;
                    }
                }

                if (!processed_split) resultToString.Append(test_split).Append(", ");
            }
            if (resultToString.Length > 0) resultToString.Remove(resultToString.Length - 2, 2);
            if (place == Place.visit) place_place_v = resultToString.ToString(); else place_place_o = resultToString.ToString();
        }

        public void FillId(string text)
        {
            int test_id = -1;
            if (Int32.TryParse(text, out test_id)) uik_id = test_id;
        }

        public string OfficeAddrToString()
        {
            resultToString.Clear();
            //resultToString.Append(place_addr_o).Append(", ").Append(street_addr_o).Append(", ").Append(building_addr_o);
            if (place_addr_o != null) resultToString.Append(place_addr_o).Append(", ");
            if (street_addr_o != null) resultToString.Append(street_addr_o).Append(", ");
            if (building_addr_o != null) resultToString.Append(building_addr_o).Append(", ");
            if (resultToString.Length > 1) resultToString.Length -= 2;
            return resultToString.ToString();
        }

        public string OfficePlaceToString()
        {
            return place_place_o;
        }

        public string VisitPlaceToString()
        {
            return place_place_v;
        }

        public string VisitAddrToString()
        {
            resultToString.Clear();
            //resultToString.Append(place_addr_v).Append(", ").Append(street_addr_v).Append(", ").Append(building_addr_v);
            if (place_addr_v != null) resultToString.Append(place_addr_v).Append(", ");
            if (street_addr_v != null) resultToString.Append(street_addr_v).Append(", ");
            if (building_addr_v != null) resultToString.Append(building_addr_v).Append(", ");
            if (resultToString.Length > 1) resultToString.Length -= 2;
            return resultToString.ToString();
        }

        public string VisitPhoneToString()
        {
            return phone_v;
        }

        public string OfficePhoneToString()
        {
            return phone_o;
        }

        public string UikToString()
        {
            return uik_id.ToString();
        }

        // "uik;addr_v;place_v;phone_v;addr_o;place_o;phone_o;comment;g_status";
        public static string GetHeader()
        {
            return csv_header;
        }

        public string GetRow()
        {
            StringBuilder rowResultToString = new StringBuilder(1024);
            rowResultToString.Append(UikToString()).Append(";");

            rowResultToString.Append(VisitAddrToString()).Append(";");
            rowResultToString.Append(VisitPlaceToString()).Append(";");
            rowResultToString.Append(VisitPhoneToString()).Append(";");

            rowResultToString.Append(OfficeAddrToString()).Append(";");
            rowResultToString.Append(OfficePlaceToString()).Append(";");
            rowResultToString.Append(OfficePhoneToString()).Append(";");

            rowResultToString.Append(FullAddressVisit).Append(" / ").Append(FullAddressOffice).Append(";");
            rowResultToString.Append(";");

            return rowResultToString.ToString();
        }

        public void Clear(Place place)
        {
            throw new NotImplementedException();
            
            if (place == Place.office)
            {
                place_addr_v = null;
                street_addr_v = null;
                building_addr_v = null;

            }
            else
            {


            }
        }

        public void SetAddrPlace(string addr, Place place)
        {
            if (place == Place.office) place_addr_o = addr; else place_addr_v = addr;
        }
    }
}
