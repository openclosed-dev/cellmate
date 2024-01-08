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
using System;
using Microsoft.Office.Interop.Excel;

namespace Cellmate
{
    interface IPageNumberRenderer
    {
        void RenderPageNumber(Workbook book, int totalPages);
    }

    class NoopPageRenderer : IPageNumberRenderer
    {
        public void RenderPageNumber(Workbook book, int totalPages)
        {
            // Do nothing
        }
    }

    class SimplePageNumberRenderer : IPageNumberRenderer
    {
        private readonly Action<Worksheet, string> pageNumberingAction;
        private readonly string format;
        private string text;

        public SimplePageNumberRenderer(PageNumberPosition position, string format) 
        {
            this.pageNumberingAction = NumberingPageAt(position);
            this.format = format;
            this.text = format;
        }

        public void RenderPageNumber(Workbook book, int totalPages)
        {
            bool firstPage = true;

            foreach (Worksheet sheet in book.Worksheets)
            {
                if (sheet.Visible != XlSheetVisibility.xlSheetVisible)
                    continue;

                if (firstPage)
                {
                    ProcessFirstSheet(book, sheet, totalPages);
                    firstPage = false;
                }
                
                pageNumberingAction.Invoke(sheet, this.text);
            }
        }

        protected virtual void ProcessFirstSheet(Workbook book, Worksheet sheet, int totalPages)
        {
            string basename = book.Name.Substring(0, book.Name.LastIndexOf('.'));
            this.text = this.format.Replace("{basename}", basename);
        }

        static Action<Worksheet, string> NumberingPageAt(PageNumberPosition position)
        {
             switch (position)
            {
                case PageNumberPosition.Left:
                    return (sheet, format) => sheet.PageSetup.LeftFooter = format;
                case PageNumberPosition.Center:
                    return (sheet, format) => sheet.PageSetup.CenterFooter = format;
                case PageNumberPosition.Right:
                    return (sheet, format) => sheet.PageSetup.RightFooter = format;
            }
            return (sheet, format) => {};
       }
    }

    class ContinuousPageNumberRenderer : SimplePageNumberRenderer
    {

        public ContinuousPageNumberRenderer(PageNumberPosition position, string format)
        : base(position, format)
        {
        }

        protected override void ProcessFirstSheet(Workbook book, Worksheet sheet, int totalPages)
        {
            base.ProcessFirstSheet(book, sheet, totalPages);
            sheet.PageSetup.FirstPageNumber = totalPages + 1;
        }
    }
}