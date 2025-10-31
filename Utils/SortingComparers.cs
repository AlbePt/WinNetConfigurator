using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Sockets;

namespace WinNetConfigurator.Utils
{
    class CabinetNameComparer : IComparer<string>
    {
        static readonly StringComparer TextComparer = StringComparer.CurrentCultureIgnoreCase;

        public int Compare(string x, string y)
        {
            bool xHasNumber = TryParseCabinet(x, out int xNumber, out string xText);
            bool yHasNumber = TryParseCabinet(y, out int yNumber, out string yText);

            if (xHasNumber && yHasNumber)
            {
                int numberCompare = xNumber.CompareTo(yNumber);
                if (numberCompare != 0)
                    return numberCompare;

                return TextComparer.Compare(xText, yText);
            }

            if (xHasNumber != yHasNumber)
                return xHasNumber ? -1 : 1;

            return TextComparer.Compare(Normalize(x), Normalize(y));
        }

        static bool TryParseCabinet(string value, out int number, out string text)
        {
            number = 0;
            text = string.Empty;

            if (string.IsNullOrWhiteSpace(value))
            {
                text = string.Empty;
                return false;
            }

            value = value.Trim();
            int index = 0;
            while (index < value.Length && char.IsDigit(value[index]))
                index++;

            if (index == 0)
            {
                text = value;
                return false;
            }

            if (!int.TryParse(value.Substring(0, index), out number))
            {
                text = value;
                return false;
            }

            text = value.Substring(index).TrimStart();
            return true;
        }

        static string Normalize(string value) => string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();
    }

    class IpAddressComparer : IComparer<string>
    {
        public int Compare(string x, string y)
        {
            bool xEmpty = string.IsNullOrWhiteSpace(x);
            bool yEmpty = string.IsNullOrWhiteSpace(y);
            if (xEmpty || yEmpty)
            {
                if (xEmpty && yEmpty)
                    return 0;
                return xEmpty ? 1 : -1;
            }

            bool xValid = TryParseIPv4(x, out uint xValue);
            bool yValid = TryParseIPv4(y, out uint yValue);

            if (xValid && yValid)
                return xValue.CompareTo(yValue);

            if (xValid != yValid)
                return xValid ? -1 : 1;

            return StringComparer.CurrentCultureIgnoreCase.Compare(x, y);
        }

        static bool TryParseIPv4(string value, out uint numeric)
        {
            numeric = 0;
            if (!IPAddress.TryParse(value, out var address) || address.AddressFamily != AddressFamily.InterNetwork)
                return false;

            var bytes = address.GetAddressBytes();
            numeric = ((uint)bytes[0] << 24) | ((uint)bytes[1] << 16) | ((uint)bytes[2] << 8) | bytes[3];
            return true;
        }
    }
}
