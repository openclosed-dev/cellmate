﻿#region copyright
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
using CommandLine;
using Microsoft.Office.Interop.Excel;

namespace Cellmate
{
    [Verb("rewrite-date", HelpText = "Rewrite dates")]
    class RewriteDateCommand : DateCommand, IEditable
    {
        [Option("new-date",
            Required = true,
            HelpText = "New value with which dates are replaced.")]
        public DateTime NewDate { get; set; }

        [Option("date-format",
            Default = "m/d/yyyy",
            HelpText = "Date format to assign")]
        public string DateFormat { get; set; } 

        protected override void ProcessDate(Workbook book, Worksheet sheet, Range cell, DateTime value)
        {
            cell.NumberFormat = this.DateFormat;
            cell.Value = this.NewDate;
            var address = cell.Address[false, false];
            Out.WriteLine($"{book.Name}:{sheet.Name}:{address} {value} => {this.NewDate}");
        }
    }
}