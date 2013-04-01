using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace PPT2Vedio
{
    class LogFactory
    {
        public static void Info(String message, params Object[] args)
        {
            Console.ForegroundColor = ConsoleColor.Green;

            message = ParseMessage(message, args);

            Console.WriteLine(message);
        }

        public static void Error(String message, params Object[] args)
        {
            Console.ForegroundColor = ConsoleColor.Red;

            message = ParseMessage(message, args);

            Console.WriteLine(message);
        }

        public static void Debug(String message, params Object[] args)
        {
            Console.ForegroundColor = ConsoleColor.DarkBlue;

            message = ParseMessage(message, args);

            Console.WriteLine(message);
        }

        private static string ParseMessage(String message, Object[] args)
        {
            var reg = new Regex(@"(?<index>\{\d\})");
            var matches = reg.Matches(message);
            if (matches.Count <= args.Length)
            {
                for (int i = 0; i < matches.Count; i++)
                {
                    Match match = matches[i];
                    message = message.Replace(match.Value, args[i].ToString());
                }
            }
            else
            {
                throw new ArgumentException("格式化参数数量不符!");
            }

            return message;
        }
    }
}
