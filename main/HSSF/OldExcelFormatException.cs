﻿using System;

namespace NPOI.HSSF
{
    [Serializable]
    public class OldExcelFormatException:Exception
    {
        public OldExcelFormatException(string s)
            : base(s)
        { }

    }
}
