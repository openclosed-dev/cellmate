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

namespace Cellmate.Cmdlets
{
    [Cmdlet(VerbsData.Import, "Excel"),
     OutputType(typeof(Workbook))]
    public class ImportExcelCmdlet : Cmdlet
    {
        private Application excel;
        private bool visible;

        [Parameter(
            ValueFromPipeline = true,
            Mandatory = true)]
        public FileInfo InputObject { get; set; }

        [Parameter()]
        public SwitchParameter Visible 
        { 
            get { return visible; }
            set { visible = value; }
        }
        protected override void BeginProcessing()
        {
            excel = new Application();
            excel.Visible = this.Visible;
            excel.DisplayAlerts = false;
        }

        protected override void ProcessRecord()
        {
            var fullName = InputObject.FullName; 
            WriteVerbose($"Loading a book: {fullName}");
            Workbook book = excel.Workbooks.Open(fullName);
            try 
            {
                WriteObject(book);
            }
            finally
            {
                book.Close();
            }
        }

        protected override void StopProcessing()
        {
            CloseExcel();
        }

        protected override void EndProcessing()
        {
            CloseExcel();
        }

        private void CloseExcel()
        {
            if (excel != null)
            {
                excel.Quit();
            }
        }
    }
}
