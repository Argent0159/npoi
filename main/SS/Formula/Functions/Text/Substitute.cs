/*
* Licensed to the Apache Software Foundation (ASF) Under one or more
* contributor license agreements.  See the NOTICE file distributed with
* this work for Additional information regarding copyright ownership.
* The ASF licenses this file to You Under the Apache License, Version 2.0
* (the "License"); you may not use this file except in compliance with
* the License.  You may obtain a copy of the License at
*
*     http://www.apache.org/licenses/LICENSE-2.0
*
* Unless required by applicable law or agreed to in writing, software
* distributed Under the License is distributed on an "AS Is" BASIS,
* WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
* See the License for the specific language governing permissions and
* limitations Under the License.
*/
/*
 * Created on May 15, 2005
 *
 */
namespace NPOI.SS.Formula.Functions
{
    using System;
    using System.Text;
    using NPOI.SS.Formula.Eval;

    /**
     * An implementation of the SUBSTITUTE function:
     * Substitutes text in a text string with new text, some number of times.
     * @author Manda Wilson &lt; wilson at c bio dot msk cc dot org &gt;
     */
    public class Substitute : TextFunction
    {
        // fix warning CS0414 "never used": private static int Replace_ALL = -1;

        /**
         *Substitutes text in a text string with new text, some number of times.
         * 
         * @see org.apache.poi.hssf.record.formula.eval.Eval
         */
        public override ValueEval EvaluateFunc(ValueEval[] args, int srcCellRow, int srcCellCol)
        {
            if (args.Length < 3 || args.Length > 4)
            {
                return ErrorEval.VALUE_INVALID;
            }

            string oldStr = EvaluateStringArg(args[0], srcCellRow, srcCellCol);
            string searchStr = EvaluateStringArg(args[1], srcCellRow, srcCellCol);
            string newStr = EvaluateStringArg(args[2], srcCellRow, srcCellCol);


            string result;

            switch (args.Length)
            {
                case 4:
                    int instanceNumber = EvaluateIntArg(args[3], srcCellRow, srcCellCol);
                    if (instanceNumber < 1)
                    {
                        return ErrorEval.VALUE_INVALID;
                    }
                    result = ReplaceOneOccurrence(oldStr, searchStr, newStr, instanceNumber);
                    break;
                case 3:
                    result = ReplaceAllOccurrences(oldStr, searchStr, newStr);
                    break;
                default:
                    throw new InvalidOperationException("Cannot happen");

            }

            return new StringEval(result);
        }

        private static string ReplaceAllOccurrences(string oldStr, string searchStr, string newStr)
        {
            StringBuilder sb = new StringBuilder();
            int startIndex = 0;
            int nextMatch = -1;
            while (true)
            {
                nextMatch = oldStr.IndexOf(searchStr, startIndex, StringComparison.CurrentCulture);
                if (nextMatch < 0)
                {
                    // store everything from end of last match to end of string
                    sb.Append(oldStr.Substring(startIndex));
                    return sb.ToString();
                }
                // store everything from end of last match to start of this match
                sb.Append(oldStr.Substring(startIndex, nextMatch - startIndex));
                sb.Append(newStr);
                startIndex = nextMatch + searchStr.Length;
            }
        }

        private static string ReplaceOneOccurrence(string oldStr, string searchStr, string newStr, int instanceNumber)
        {
            if (searchStr.Length < 1)
            {
                return oldStr;
            }
            int startIndex = 0;
            int nextMatch = -1;
            int count = 0;
            while (true)
            {
                nextMatch = oldStr.IndexOf(searchStr, startIndex, StringComparison.CurrentCulture);
                if (nextMatch < 0)
                {
                    // not enough occurrences found - leave unchanged
                    return oldStr;
                }
                count++;
                if (count == instanceNumber)
                {
                    StringBuilder sb = new StringBuilder(oldStr.Length + newStr.Length);
                    sb.Append(oldStr.Substring(0, nextMatch));
                    sb.Append(newStr);
                    sb.Append(oldStr.Substring(nextMatch + searchStr.Length));
                    return sb.ToString();
                }
                startIndex = nextMatch + searchStr.Length;
            }
        }
    }
}