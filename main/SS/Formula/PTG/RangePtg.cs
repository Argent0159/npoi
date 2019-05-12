/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */

namespace NPOI.SS.Formula.PTG
{
    using System;
    using System.Text;
    
    using NPOI.Util;

    /**
     * @author Daniel Noll (daniel at nuix dot com dot au)
     */
    public class RangePtg : OperationPtg
    {
        public const int SIZE = 1;
        public const byte sid = 0x11;

        public static OperationPtg instance = new RangePtg();

        private RangePtg()
        {
            // enforce singleton
        }

        public override bool IsBaseToken
        {
            get { return true; }
        }

        public override int Size
        {
            get { return SIZE; }
        }

        public override void Write(ILittleEndianOutput out1)
        {
            out1.WriteByte(sid + PtgClass);
        }

        public override string ToFormulaString()
        {
            return ":";
        }


        /** implementation of method from OperationsPtg*/
        public override string ToFormulaString(string[] operands)
        {
            StringBuilder buffer = new StringBuilder();

            buffer.Append(operands[0]);
            buffer.Append(":");
            buffer.Append(operands[1]);
            return buffer.ToString();
        }

        public override int NumberOfOperands
        {
            get { return 2; }
        }

    }
}
