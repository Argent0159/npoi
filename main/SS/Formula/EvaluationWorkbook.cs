/* ====================================================================
   Licensed To the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file To You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed To in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace NPOI.SS.Formula
{

    using System;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.Udf;
    using NPOI.SS.Formula.PTG;

    public class ExternalSheet
    {
        private string _workbookName;
        private string _sheetName;

        public ExternalSheet(string workbookName, string sheetName)
        {
            _workbookName = workbookName;
            _sheetName = sheetName;
        }
        public string WorkbookName
        {
            get
            {
                return _workbookName;
            }
        }
        public string SheetName
        {
            get
            {
                return _sheetName;
            }
        }
    }

    public class ExternalSheetRange : ExternalSheet
    {
        private string _lastSheetName;
        public ExternalSheetRange(string workbookName, string firstSheetName, string lastSheetName)
            : base(workbookName, firstSheetName)
        {
            this._lastSheetName = lastSheetName;
        }

        public string FirstSheetName
        {
            get
            {
                return SheetName;
            }
        }
        public string LastSheetName
        {
            get
            {
                return _lastSheetName;
            }
        }
    }

    /**
     * Abstracts a workbook for the purpose of formula evaluation.<br/>
     * 
     * For POI internal use only
     * 
     * @author Josh Micich
     */
    public interface IEvaluationWorkbook
    {
        string GetSheetName(int sheetIndex);
        /**
         * @return -1 if the specified sheet is from a different book
         */
        int GetSheetIndex(IEvaluationSheet sheet);
        int GetSheetIndex(string sheetName);

        IEvaluationSheet GetSheet(int sheetIndex);

        /**
         * HSSF Only - fetch the external-style sheet details
         * <p>Return will have no workbook set if it's actually in our own workbook</p>
         */
        ExternalSheet GetExternalSheet(int externSheetIndex);
        /**
         * XSSF Only - fetch the external-style sheet details
         * <p>Return will have no workbook set if it's actually in our own workbook</p>
         */
        ExternalSheet GetExternalSheet(string firstSheetName, string lastSheetName, int externalWorkbookNumber);
        /**
         * HSSF Only - convert an external sheet index to an internal sheet index,
         *  for an external-style reference to one of this workbook's own sheets 
         */
        int ConvertFromExternSheetIndex(int externSheetIndex);
        /**
         * HSSF Only - fetch the external-style name details
         */
        ExternalName GetExternalName(int externSheetIndex, int externNameIndex);
        /**
         * XSSF Only - fetch the external-style name details
         */
        ExternalName GetExternalName(string nameName, string sheetName, int externalWorkbookNumber);
    
        IEvaluationName GetName(NamePtg namePtg);
        IEvaluationName GetName(string name, int sheetIndex);
        string ResolveNameXText(NameXPtg ptg);
        Ptg[] GetFormulaTokens(IEvaluationCell cell);
        UDFFinder GetUDFFinder();
    }

    public class ExternalName
    {
        private string _nameName;
        private int _nameNumber;
        private int _ix;

        public ExternalName(string nameName, int nameNumber, int ix)
        {
            _nameName = nameName;
            _nameNumber = nameNumber;
            _ix = ix;
        }
        public string Name
        {
            get
            {
                return _nameName;
            }
        }
        public int Number
        {
            get
            {
                return _nameNumber;
            }
        }
        public int Ix
        {
            get
            {
                return _ix;
            }
        }
    }
}