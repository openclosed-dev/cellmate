#region copyright
/*
 * Copyright 2020-2024 the original author or authors.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
#endregion
using System.Management.Automation;
using Microsoft.Office.Interop.Excel;

namespace Cellmate
{
    [Cmdlet(VerbsData.Merge, "BorderedWorksheetRange"),
     OutputType(typeof(Workbook))]
    public class MergeBorderedWorksheetRangeCmdlet : RangeCmdlet
    {
        [Parameter(Mandatory = true)]
        [ValidatePattern(RangePattern)]
        public override string[] Range { get; set; }

        [Parameter(
            HelpMessage = "Column offset starting from zero"
        )]
        public int ColumnOffset { get; set; }

        public override bool UsedRangeOnly { get => false; }

        protected override void ProcessRange(Workbook book, Worksheet sheet, Range range)
        {
            WriteVerbose($"Default worksheet range: {ToAddress(range)}");
            Range rangeToMerge = GetBorderedRange(range);
            if (rangeToMerge == null)
                return;

            WriteVerbose($"Bordered worksheet range: {ToAddress(rangeToMerge)}");
            if (ColumnOffset > 0)
            {
                rangeToMerge = GetSkippedRange(rangeToMerge, ColumnOffset);
                if (rangeToMerge == null)
                    return;
            }
            WriteVerbose($"Merging worksheet range: {ToAddress(rangeToMerge)}");
            rangeToMerge.Merge();
        }

        private Range GetSkippedRange(Range range, int columnsToSkip)
        {
            Range topLeft = range.Cells[1, 1] as Range;
            Range moved = topLeft.Offset[0, columnsToSkip];

            if (moved.Column > range.Columns.Count)
                return null;

            return range.Range[moved, range.Cells[range.Rows.Count, range.Columns.Count]];
        }

        private Range GetBorderedRange(Range range)
        {
            int column = 0;
            while (true)
            {
                var cell = range.Cells[1, column + 1] as Range;
                if (HasTopBorder(cell))
                {
                    column++;
                }
                else
                {
                    break;
                }
            }

            if (column <= 0)
                return null;

            return range.Range[range.Cells[1, 1], range.Cells[range.Rows.Count, column]];
        }

        static bool HasTopBorder(Range cell)
        {
            var border = cell.Borders[XlBordersIndex.xlEdgeTop];
            if (border.LineStyle is int i)
            {
                return i != (int)XlLineStyle.xlLineStyleNone;
            }
            return false;
        }

        static string ToAddress(Range range)
        {
            return range.Address[false, false];
        }
    }
}
