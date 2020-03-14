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
using System.Collections.Generic;
using System.Management.Automation;
using System.IO;
using System.Text.RegularExpressions;
using CommandLine;
using Cellmate.Cmdlets;

namespace Cellmate 
{
    abstract class ExcelCommand : Command
    {
        private static readonly Regex RangeRegex = new Regex(
            @"^([A-Z]+|\d+|[A-Z]+\d+)(:([A-Z]+|\d+|[A-Z]+\d+))?$",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private string range;
        private readonly bool editable;

        protected ExcelCommand()
        {
            this.editable = this is IEditable;
        }

        [Value(0, MetaName = "files",
            HelpText = "Excel files to be processed.",
            Required = true)]
        public IEnumerable<string> Files { get; set; }

        [Option("visible", 
            HelpText = "Visibility of the Excel window.")]
        public bool Visible { get; set; }

        [Option("range", 
            HelpText = "Range of cells to be processed, e.g. \"A1:Z99\"")]
        public string Range 
        { 
            get { return range; } 
            set
            {
                ValidateRange(value);
                this.range = value;
            } 
        }

        public string NewSuffix { get; set; }

        public bool Inplace { get; set; }

        public bool IsEditable => editable;

        public override int Execute()
        {
            using (var powerShell = BuildPipeline())
            {
                var executor = new PowerShellExecutor(this.Out, this.Error);
                return executor.Execute(powerShell, Files);
            }
        }

        protected PowerShell BuildPipeline()
        {
            PowerShell ps = PowerShell.Create();

            ps.AddCommand("Get-Item");
            ps.AddCommand(new CmdletInfo("Import-Excel", typeof(ImportExcelCmdlet)));
            if (Visible)
            {
                ps.AddParameter("Visible");
            }

            AddCmdletsTo(ps);
            
            if (IsEditable)
            {
                AddExportExcelCmdlet(ps);
            }
            
            return ps;
        }

        protected abstract void AddCmdletsTo(PowerShell pipeline);

        protected void AddExportExcelCmdlet(PowerShell pipeline)
        {
            pipeline.AddCommand(new CmdletInfo("Export-Excel", typeof(ExportExcelCmdlet)));
            if (!Inplace)
            {
                pipeline.AddParameter("Suffix", NewSuffix);
            }
        }

        void ValidateRange(string value)
        {
            if (RangeRegex.IsMatch(value))
            {
                this.range = value;
            }
            else
            {
                throw new ArgumentException();
            }
        }
    }
}