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
using System.Management.Automation;
using Microsoft.Office.Interop.Excel;

namespace Cellmate
{
    [Cmdlet(VerbsData.Merge, "Workbook"),
     OutputType(typeof(Workbook))]
    public class MergeWorkbookCmdlet : WorkbookCmdlet
    {
        public enum OutputFormat
        {
            Pdf
        }

        private IWorkbookCombiner combiner;

        [Parameter(Mandatory = true)]
        public OutputFormat As { get; set; }

        [Parameter(Position = 0, Mandatory = true)]
        public string Destination { get; set; }

        [Parameter()]
        public PageNumberPosition PageNumber { get; set; } = PageNumberPosition.None;

        [Parameter()]
        public SwitchParameter SaveEach { get; set; }

        [Parameter()]
        public SwitchParameter RestartPageNumber { get; set; }

        [Parameter()]
        public string PageNumberFormat { get; set; } = "&P";

        protected override void BeginProcessing()
        {
            this.combiner = new PdfWorkbookCombiner(
                CreatePageNumberRenderer(PageNumber, PageNumberFormat, RestartPageNumber.IsPresent),
                SaveEach.IsPresent);
        }

        protected override void EndProcessing()
        {
            string path = ResolvePath(Destination);
            WriteVerbose($"Writing a {combiner.FormatName}: {path}");
            this.combiner.SaveAs(path);
            this.combiner.Close();
        }

        protected override void StopProcessing()
        {
            this.combiner.Close();
        }

        protected override void ProcessBook(Workbook book)
        {
            WriteVerbose($"Appending a workbook: {book.FullName}");
            this.combiner.Append(book);
        }

        private static IPageNumberRenderer CreatePageNumberRenderer(
            PageNumberPosition position, string format, bool restartPageNumber) {
            if (position == PageNumberPosition.None) 
            {
                return new NoopPageRenderer();
            } 
            else if (restartPageNumber) 
            {
                return new SimplePageNumberRenderer(position, format);
            } 
            else {
                return new ContinuousPageNumberRenderer(position, format);
            }
        }
    }
}
