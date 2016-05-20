using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClouDeveloper.HWP
{
    public enum HwpParseState
    {
        HWP_PARSE_STATE_NORMAL,
        HWP_PARSE_STATE_PASSING,
        HWP_PARSE_STATE_DOCSUMMARY,
        HWP_PARSE_STATE_P,
        HWP_PARSE_STATE_TEXT,
        HWP_PARSE_STATE_CHAR
    }
}
