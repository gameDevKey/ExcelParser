using System;
using System.Collections.Generic;

namespace ExcelParserGUI
{
    internal class Config
    {
        internal static Dictionary<string, Type> Handlers = new Dictionary<string, Type>
        {
            {"json", typeof(JsonHandler) },
        };
    }
}
