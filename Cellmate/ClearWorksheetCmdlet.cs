#region copyright
/*
 * Copyright 2024 the original author or authors.
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
    public enum Area {
        NonPrint
    }

    [Cmdlet(VerbsCommon.Clear, "Worksheet"), OutputType(typeof(Workbook))]
    public class ClearWorksheetCmdlet : WorksheetCmdlet
    {
        [Parameter()]
        public Area Area { get; set; } = Area.NonPrint;

        protected override void ProcessSheet(Workbook book, Worksheet sheet)
        {
            if (Area == Area.NonPrint)
                ClearNonPrintArea(book, sheet);
        }

        private void ClearNonPrintArea(Workbook book, Worksheet sheet)
        {
            string printArea = sheet.PageSetup.PrintArea;
            if (printArea != null && printArea != "")
            {
                // Converts the value in R1C1 style to A1 style.
                if (book.Application.ReferenceStyle == XlReferenceStyle.xlR1C1)
                    printArea = book.Application.ConvertFormula(printArea, XlReferenceStyle.xlR1C1, XlReferenceStyle.xlA1) as string;
                var printRange = sheet.Range[printArea];
                var outer = ComputeOuterRange(sheet, printRange);
                if (outer != null)
                {
                    WriteVerbose($"Clearing non-print areas: {outer.Address} on sheet {sheet.Name}");
                    outer.UnMerge();
                    outer.Clear();
                }
            }
        }

        private Range ComputeOuterRange(Worksheet sheet, Range inner)
        {
            var used = sheet.UsedRange;

            int innerRowCount = inner.Rows.Count;
            int innerColumnCount = inner.Columns.Count;
            int usedRowCount = used.Rows.Count;
            int usedColumnCount = used.Columns.Count;

            Range outer = null;
            var lastUsedCell = used.Cells[usedRowCount, usedColumnCount];

            if (innerColumnCount < usedColumnCount) {
                outer = used.Range[used.Cells[1, innerColumnCount + 1], lastUsedCell];
            }

            if (innerRowCount < usedRowCount) {
                var bottom = used.Range[used.Cells[innerRowCount + 1, 1], lastUsedCell];
                if (outer == null) {
                    outer = bottom;
                } else {
                    outer = sheet.Application.Union(outer, bottom);
                }
            }

            return outer;
        }
    }
}
