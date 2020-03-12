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
using System;
using System.Management.Automation;
using Microsoft.Office.Interop.Excel;

namespace Cellmate.Cmdlets
{
    public abstract class DateCellCmdlet : SheetRangeCmdlet
    {
        [Parameter()]
        public DateTime After { get; set; } = DateTime.MinValue;

        [Parameter()]
        public DateTime Before { get; set; } = DateTime.MaxValue;

        protected override void ProcessRange(Workbook book, Worksheet sheet, Range range)
        {
            int rowCount = range.Rows.Count;
            int columnCount = range.Columns.Count;
            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= columnCount; j++)
                {
                    var cell = range[i, j] as Range;
                    DateTime? value = Values.AsDateTime(cell.Value);
                    if (value.HasValue)
                    {
                        DateTime dateTime = value.Value;
                        if (IsDateInRange(dateTime))
                        {
                            ProcessDate(book, sheet, cell, dateTime);
                        }
                    } 
                }
            }
        }

        protected abstract void ProcessDate(Workbook book, Worksheet sheet, Range cell, DateTime value);

        bool IsDateInRange(DateTime value)
        {
            return After < value && value < Before;
        }
    }
}
