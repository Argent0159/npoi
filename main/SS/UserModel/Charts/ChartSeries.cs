﻿using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.UserModel.Charts
{
    public enum TitleType
    {
        String,
        CellReference
    }

    public interface IChartSeries
    {
        /**
 * Sets the title of the series as a string literal.
 *
 * @param title
 */
        void SetTitle(string title);

        /**
         * Sets the title of the series as a cell reference.
         *
         * @param titleReference
         */
        void SetTitle(CellReference titleReference);

        /**
         * @return title as string literal.
         */
        string GetTitleString();

        /**
         * @return title as cell reference.
         */
        CellReference GetTitleCellReference();

        /**
         * @return title type.
         */
        TitleType? GetTitleType();
    }
}
