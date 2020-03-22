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
using System.IO.Compression;
using System.Text;
using System.Management.Automation;
using Microsoft.Office.Interop.Excel;

namespace Cellmate
{
    [Cmdlet(VerbsData.Compress, "Workbook"),
     OutputType(typeof(Workbook))]
    public class CompressWorkbookCmdlet : WorkbookCmdlet
    {
        private Encoding entryEncoding;
        private ZipArchive zipArchive;

        [Parameter(Mandatory = true)]
        public string Destination { get; set; }

        [Parameter()]
        public string Encoding
        { 
            get
            {
                return entryEncoding.ToString();
            }
            set
            {
                this.entryEncoding = System.Text.Encoding.GetEncoding(value);
            }
        }

        public CompressWorkbookCmdlet()
        {
            this.entryEncoding = System.Text.Encoding.UTF8;
        }

        protected override void BeginProcessing()
        {
            string path = ResolvePath(Destination);
            var zipStream = new FileStream(path, FileMode.OpenOrCreate);
            this.zipArchive = new ZipArchive(zipStream, ZipArchiveMode.Update, false, entryEncoding);
        }

        protected override void EndProcessing()
        {
            WriteVerbose($"Zip entries written: {zipArchive.Entries.Count}");
            CloseZip();
        }

        protected override void StopProcessing()
        {
            CloseZip();
        }

        protected override void ProcessBook(Workbook book)
        {
            var fullName = book.FullName;
            var baseName = this.CurrentLocation;
            var entryName = fullName.Substring(baseName.Length + 1);
            WriteVerbose($"Compressing a workbook {fullName} as {entryName}");

            var sourceFileName = SaveAsTemporaryFile(book);
            try 
            {
                this.zipArchive.CreateEntryFromFile(sourceFileName, entryName); 
            }
            finally
            {
                File.Delete(sourceFileName);
            }
        }

        string SaveAsTemporaryFile(Workbook book)
        {
            string filename = Path.GetTempFileName();
            book.SaveCopyAs(filename);
            return filename;
        }

        void CloseZip()
        {
            if (zipArchive != null)
            {
                zipArchive.Dispose();
                zipArchive = null;
            }
        }
    }
}
