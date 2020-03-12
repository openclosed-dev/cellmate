#region copyright
/*
 * Copyright 2020 the original author or authors.
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

namespace Cellmate.Cmdlets
{
    public abstract class SheetRangeCmdlet : SheetCmdlet
    {
        [Parameter()]
        public string[] Range { get; set; }

        protected override void ProcessSheet(Workbook book, Worksheet sheet)
        {
            ProcessRange(book, sheet, CalculateRange(sheet));
        }

        protected abstract void ProcessRange(Workbook book, Worksheet sheet, Range range);

        Range CalculateRange(Worksheet sheet)
        {
            return sheet.UsedRange;
        }
    }
}