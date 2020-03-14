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
using System.Management.Automation;
using CommandLine;
using Cellmate.Cmdlets;

namespace Cellmate
{
    [Verb("edit-date", HelpText = "Edit date cells")]
    class EditDateCommand : DateCommand, IEditable
    {
        [Option("value",
            Required = true,
            HelpText = "New value with which dates are replaced.")]
        public DateTime NewDate { get; set; }

        [Option("format",
            Default = "m/d/yyyy",
            HelpText = "Date format to assign")]
        public string DateFormat { get; set; } 

        protected override void AddCmdletsTo(PowerShell pipeline)
        {
            pipeline.AddCommand(new CmdletInfo("Edit-DateCell", typeof(EditDateCellCmdlet)))
                .AddParameter("Verbose")
                .AddParameter("Range", Range)
                .AddParameter("Before", Before)
                .AddParameter("After", After)
                .AddParameter("Value", NewDate);
            
            if (DateFormat != null)
            {
                pipeline.AddParameter("Format", DateFormat);
            }
        }
    }
}