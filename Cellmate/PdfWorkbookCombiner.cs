#region copyright
/*
 * Copyright 2020-2024 the original author or authors.
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
using Microsoft.Office.Interop.Excel;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace Cellmate
{
    class PdfWorkbookCombiner : IWorkbookCombiner
    {
        private readonly PdfDocument targetPdf;
        private readonly IPageNumberRenderer pageNumberRenderer;
        private readonly bool saveEach;
        private int totalPages;

        public PdfWorkbookCombiner(IPageNumberRenderer pageNumberRenderer, bool saveEach)
        {
            this.targetPdf = new PdfDocument();
            this.pageNumberRenderer = pageNumberRenderer;
            this.saveEach = saveEach;
            this.totalPages = 0;
        }

        public string FormatName
        {
            get => "PDF";
        }

        public void Append(Workbook book)
        {
            pageNumberRenderer.RenderPageNumber(book, totalPages);

            string path = System.IO.Path.ChangeExtension(book.FullName, ".pdf");
            book.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, path);

            try
            {
                AppendPdf(path);
            }
            finally
            {
                if (!this.saveEach)
                    File.Delete(path);
            }
        }

        public void SaveAs(string path)
        {
            if (this.targetPdf.PageCount > 0)
            {
                this.targetPdf.Save(path);
            }
        }

        public void Close()
        {
            this.targetPdf.Close();
        }

        void AppendPdf(string path)
        {
            using (PdfDocument pdf = PdfReader.Open(path, PdfDocumentOpenMode.Import))
            {
                foreach (var page in pdf.Pages)
                {
                    this.targetPdf.AddPage(page);
                }
                this.totalPages += pdf.PageCount;
            }
        }
    }
}
