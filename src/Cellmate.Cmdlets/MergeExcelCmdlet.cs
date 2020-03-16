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
    [Cmdlet(VerbsData.Merge, "Excel"),
     OutputType(typeof(Workbook))]
    public class MergeExcelCmdlet : BookCmdlet
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
                WriteVerbose($"Writing a PDF: {Destination}");
                this.targetPdf.Save(Destination);
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
            int pageCount = 0;
            Window window = book.Windows[1];
            
            foreach (Worksheet sheet in book.Worksheets)
            {
                if (sheet.Visible == XlSheetVisibility.xlSheetVisible)
                {
                    sheet.Activate();
                    window.View = XlWindowView.xlPageBreakPreview;
                    
                    if (pageCount == 0)
                    {
                        sheet.PageSetup.FirstPageNumber = this.pageTotal + 1;
                    }
                    sheet.PageSetup.RightFooter = "&P";
                    pageCount += sheet.PageSetup.Pages.Count;
                }
            }

            this.pageTotal += pageCount;
        }

        void AddBook(Workbook book)
        {
            WriteVerbose($"Merging a book: {book.FullName}");

            string path = System.IO.Path.ChangeExtension(book.FullName, ".pdf");
            book.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, path);

            try
            {
                AddPdf(path);
            }
            finally
            {
                if (!Keep.IsPresent)
                {
                    System.IO.File.Delete(path);
                }
            }
        }

        void AddPdf(string path)
        {
            using (PdfDocument pdf = PdfReader.Open(path, PdfDocumentOpenMode.Import))
            {
                foreach (var page in pdf.Pages)
                {
                    this.targetPdf.AddPage(page);
                }
            }
        }
    }
}
