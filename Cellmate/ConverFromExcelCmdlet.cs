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
using System.Management.Automation;
using Microsoft.Office.Interop.Excel;

namespace Cellmate
{
    [Cmdlet(VerbsData.ConvertFrom, "Excel"),
     OutputType(typeof(Workbook))]
    public class ConvertFromExcelCmdlet : WorkbookCmdlet
    {
        [Parameter()]
        public string Suffix { get; set; }
        
        [Parameter(Mandatory = true)]
        [ValidateSet("pdf", "xps", IgnoreCase = true)]
        public string Format { get; set; }

        protected override void ProcessBook(Workbook book)
        {
            ExportBookAsFixedFormat(book, Format);
        }

        void ExportBookAsFixedFormat(Workbook book, string format)
        {
            switch (format.ToLower())
            {
                case "pdf":
                    ExportBookAs(book, XlFixedFormatType.xlTypePDF, ".pdf"); 
                    break;
                case "xps":
                    ExportBookAs(book, XlFixedFormatType.xlTypeXPS, ".xps"); 
                    break;
                default:
                    throw new InvalidOperationException();
            }
        }

        void ExportBookAs(Workbook book, XlFixedFormatType format, string suffix)
        {
            var filename = Path.ChangeExtension(book.FullName, suffix);
            book.ExportAsFixedFormat(format, filename);
        }
    }
}
