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
using Microsoft.Office.Interop.Excel;

namespace Cellmate.Cmdlets
{
    public abstract class CellCmdlet : RangeCmdlet
    {
        protected override void ProcessRange(Workbook book, Worksheet sheet, Range range)
        {
            int rowCount = range.Rows.Count;
            int columnCount = range.Columns.Count;
            var values = range.Value as object[,];
            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= columnCount; j++)
                {
                    var value = values[i, j];
                    if (TestCell(value))
                    {
                        ProcessCell(book, sheet, range[i, j] as Range, value);
                    }
                }
            }
        }

        protected abstract bool TestCell(object value);

        protected abstract void ProcessCell(Workbook book, Worksheet sheet, Range cell, object value);
    }
}
