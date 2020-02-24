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
using System.IO;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using CommandLine;

namespace Cellmate 
{
    abstract class ExcelCommand : Command
    {
        private static readonly Regex RangeRegex = new Regex(
            @"^([A-Z]+|\d+|[A-Z]+\d+)(:([A-Z]+|\d+|[A-Z]+\d+))?$",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private string range;
        private readonly bool editable;

        protected ExcelCommand()
        {
            this.editable = this is IEditable;
        }

        [Option("visible", 
            Default = true,
            HelpText = "Visibility of the Excel window.")]
        public bool? Visible { get; set; }

        [Option("range", 
            HelpText = "Range of cells to be processed, e.g. \"A1:Z99\"")]
        public string Range 
        { 
            get { return range; } 
            set
            {
                ValidateRange(value);
                this.range = value;
            } 
        }

        public string NewSuffix { get; set; }

        public bool IsEditable => editable;

        public override int Execute()
        {
            var excel = new Application();
            excel.Visible = Visible.Value;
            excel.DisplayAlerts = false;
            try
            {
                foreach (var file in Files)
                {
                    var path = Path.GetFullPath(file);
                    Workbook book = excel.Workbooks.Open(path);
                    try
                    {
                        ProcessBook(book);
                        if (IsEditable)
                        {
                            SaveBookAs(book, GenerateNewPath(path));
                        }
                    }
                    finally
                    {
                        book.Close();
                    }
                }
                return 0;
            }
            catch (Exception e)
            {
                Error.WriteLine(e.Message);
                return 1;
            }
            finally
            {
                excel.Quit();
            }
        }

        protected virtual void ProcessBook(Workbook book)
        {
            foreach (Worksheet sheet in book.Worksheets)
            {
                ProcessSheet(book, sheet);
            }
        }

        protected virtual void ProcessSheet(Workbook book, Worksheet sheet)
        {
            ProcessRange(book, sheet, CalculateRange(sheet));
        }

        protected abstract void ProcessRange(Workbook book, Worksheet sheet, Range range);

        void ValidateRange(string value)
        {
            if (RangeRegex.IsMatch(value))
            {
                this.range = value;
            }
            else
            {
                throw new ArgumentException();
            }
        }

        Range CalculateRange(Worksheet sheet)
        {
            Range usedRange = sheet.UsedRange;
            if (this.Range != null) 
            {
                return sheet.Application.Intersect(usedRange, sheet.Range[this.Range]);
            }
            else
            {
                return usedRange;
            }
        }

        void SaveBookAs(Workbook book, String path)
        {
            book.SaveAs(path);
        }

        string GenerateNewPath(string path)
        {
            return Path.ChangeExtension(path, this.NewSuffix);
        }
    }
}