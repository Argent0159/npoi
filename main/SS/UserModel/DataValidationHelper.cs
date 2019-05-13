/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */
namespace NPOI.SS.UserModel
{
    using System;

    using NPOI.SS.Util;

    /**
     * @author <a href="rjankiraman@emptoris.com">Radhakrishnan J</a>
     * 
     */
    public interface IDataValidationHelper
    {

        IDataValidationConstraint CreateFormulaListConstraint(string listFormula);

        IDataValidationConstraint CreateExplicitListConstraint(string[] listOfValues);

        IDataValidationConstraint CreateNumericConstraint(int validationType, int operatorType, string formula1, string formula2);

        IDataValidationConstraint CreateTextLengthConstraint(int operatorType, string formula1, string formula2);

        IDataValidationConstraint CreateDecimalConstraint(int operatorType, string formula1, string formula2);

        IDataValidationConstraint CreateintConstraint(int operatorType, string formula1, string formula2);

        IDataValidationConstraint CreateDateConstraint(int operatorType, string formula1, string formula2, string dateFormat);

        IDataValidationConstraint CreateTimeConstraint(int operatorType, string formula1, string formula2);

        IDataValidationConstraint CreateCustomConstraint(string formula);

        IDataValidation CreateValidation(IDataValidationConstraint constraint, CellRangeAddressList cellRangeAddressList);
    }

}