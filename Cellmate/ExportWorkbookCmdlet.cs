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
using System.IO;
using System.Management.Automation;
using Microsoft.Office.Interop.Excel;

namespace Cellmate
{
    [Cmdlet(VerbsData.Export, "Workbook"),
     OutputType(typeof(Workbook))]
    public class ExportWorkbookCmdlet : WorkbookCmdlet
    {
        public enum OutputFormat 
        {
            Default,
            Csv,
            Pdf,
            Xps
        }

        [Parameter()]
        public OutputFormat As { get; set; } = OutputFormat.Default;

        [Parameter(Position = 0)]
        public string Destination { get; set; }
        
        protected override void ProcessBook(Workbook book)
        {
            SaveBook(book);
        }

        void SaveBook(Workbook book)
        {
            switch (As)
            {
                case OutputFormat.Default:
                    SaveBookAs(book, XlFileFormat.xlWorkbookDefault, ".xlsx");
                    break; 
                case OutputFormat.Csv:
                    SaveBookAs(book, XlFileFormat.xlCSV, ".csv");
                    break; 
                case OutputFormat.Pdf:
                    ExportBookAs(book, XlFixedFormatType.xlTypePDF, ".pdf");
                    break; 
                case OutputFormat.Xps:
                    ExportBookAs(book, XlFixedFormatType.xlTypeXPS, ".xps");
                    break; 
            }
        }

        void SaveBookAs(Workbook book, XlFileFormat format, string extension)
        {
            var filename = GetTargetName(book, extension);
            WriteVerbose($"Exporting a workbook: {filename}");
            book.SaveAs(filename, format, Local : true);
        }

        void ExportBookAs(Workbook book, XlFixedFormatType format, string extension)
        {
            var filename = GetTargetName(book, extension);
            WriteVerbose($"Exporting a workbook: {filename}");
            book.ExportAsFixedFormat(format, filename);
        }

        string GetTargetName(Workbook book, string extension)
        {
            string filename = book.FullName;
            
            if (Destination != null)
                filename = Path.Combine(ResolvePath(Destination), Path.GetFileName(filename));
            
            if (extension != null)
                filename = Path.ChangeExtension(filename, extension); 
            
            return filename;
        }
    }
}
