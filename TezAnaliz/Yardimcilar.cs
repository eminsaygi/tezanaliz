namespace TezAnaliz
{
    public static class Yardimcilar
    {
        public static bool TextIsSekil(string metin)
        {
            if (metin == null) return false;
            if (!metin.StartsWith(Sabitler.Figure)) return false;

            if (metin.Length < 10) return false;

            if (metin[5] != ' ') return false;
            if (metin[7] != '.') return false;
            if (metin[9] != '.') return false;

            if (!IsNumeric(metin.Substring(6, 1))) return false;
            if (!IsNumeric(metin.Substring(6, 1))) return false;

            return true;
        }

        public static bool TextIsTablo(string metin)
        {
            if (metin == null) return false;

            if (!metin.StartsWith(Sabitler.Table)) return false;

            if (metin.Length < 10) return false;

            if (metin[5] != ' ') return false;
            if (metin[7] != '.') return false;
            if (metin[9] != '.') return false;

            if (!IsNumeric(metin.Substring(6, 1))) return false;
            if (!IsNumeric(metin.Substring(6, 1))) return false;

            return true;
        }

        public static bool IsNumeric(this string metin) => double.TryParse(metin, out _);
    }
}