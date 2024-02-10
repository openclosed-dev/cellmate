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
            var printArea = sheet.PageSetup.PrintArea;
            if (printArea != null && printArea != "")
            {
                var outer = ComputeOuterRange(book, sheet, sheet.Range[printArea]);
                if (outer != null)
                {
                    outer.Clear();
                    WriteVerbose($"Cleared non-print areas: {outer.Address}");
                }
            }
        }

        private Range ComputeOuterRange(Workbook book, Worksheet sheet, Range inner)
        {
            Application app = sheet.Application;

            Range outer = null;
            foreach (Range cell in sheet.UsedRange)
            {
                if (app.Intersect(cell, inner) == null)
                {
                    if (outer == null)
                        outer = cell;
                    else 
                        outer = app.Union(outer, cell);
                }
            }
            
            return outer;
        }
    }
}