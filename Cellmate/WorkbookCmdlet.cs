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
    public abstract class WorkbookCmdlet : PSCmdlet
    {
        [Parameter(
            ValueFromPipeline = true,
            Mandatory = true)]
        public Workbook InputObject { get; set; }

        public string CurrentLocation
        {
            get => SessionState.Path.CurrentLocation.Path;
        }

        protected override void ProcessRecord()
        {
            var book = InputObject;
            ProcessBook(book);
            WriteObject(book);
        }

        protected abstract void ProcessBook(Workbook book);

        protected string ResolvePath(string path)
        {
            return Path.Combine(CurrentLocation, path);
        }
    }
}