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
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace Cellmate.Cmdlets
{
    [Cmdlet(VerbsData.Merge, "Workbook"),
     OutputType(typeof(Workbook))]
    public class MergeWorkbookCmdlet : WorkbookCmdlet
    {
        private PdfDocument targetPdf;
        private int pageTotal;

        [Parameter(Mandatory = true)]
        public string Destination { get; set; }

        [Parameter()]
        public SwitchParameter PageNumber { get; set; }

        [Parameter()]
        public SwitchParameter Keep { get; set; }

        protected override void BeginProcessing()
        {
           this.targetPdf = new PdfDocument();
           this.pageTotal = 0;
        }

        protected override void EndProcessing()
        {
            if (this.targetPdf.PageCount > 0)
            {
                string path = ResolvePath(Destination);
                WriteVerbose($"Writing a PDF: {path}");
                this.targetPdf.Save(path);
                WriteVerbose($"Total pages written: {pageTotal}");
            }
            this.targetPdf.Close();
        }

        protected override void StopProcessing()
        {
            this.targetPdf.Close();
        }

        protected override void ProcessBook(Workbook book)
        {
            if (PageNumber.IsPresent)
            {
                AddPageNumber(book);
            }
            AddBook(book);
        }

        void AddPageNumber(Workbook book)
        {
            bool firstPage = true;
            
            foreach (Worksheet sheet in book.Worksheets)
            {
                if (sheet.Visible == XlSheetVisibility.xlSheetVisible)
                {
                    if (firstPage)
                    {
                        sheet.PageSetup.FirstPageNumber = this.pageTotal + 1;
                        firstPage = false;
                    }
                    sheet.PageSetup.RightFooter = "&P";
                }
            }
        }

        void AddBook(Workbook book)
        {
            WriteVerbose($"Appending a workbook: {book.FullName}");

            string path = System.IO.Path.ChangeExtension(book.FullName, ".pdf");
            book.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, path);

            try
            {
                AppendPdf(path);
            }
            finally
            {
                if (!Keep.IsPresent)
                {
                    System.IO.File.Delete(path);
                }
            }
        }

        void AppendPdf(string path)
        {
            using (PdfDocument pdf = PdfReader.Open(path, PdfDocumentOpenMode.Import))
            {
                foreach (var page in pdf.Pages)
                {
                    this.targetPdf.AddPage(page);
                }
                this.pageTotal += pdf.PageCount;
            }
        }
    }
}
