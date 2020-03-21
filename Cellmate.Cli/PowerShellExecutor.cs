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
using System.Collections.Generic;
using System.IO;
using System.Management.Automation;

namespace Cellmate
{
    class PowerShellExecutor
    {
        private TextWriter output;
        private TextWriter error;

        public PowerShellExecutor(TextWriter output, TextWriter error)
        {
            this.output = output;
            this.error = error;
        }

        public int Execute<T>(PowerShell powerShell, IEnumerable<T> input)
        {
            PrepareStreams(powerShell);
       
            powerShell.Invoke(input);
       
            return powerShell.HadErrors ? 1 : 0;
        }

        void PrepareStreams(PowerShell powerShell)
        {
            var streams = powerShell.Streams;
            streams.Verbose.DataAdded += OutputVerbose;
            streams.Warning.DataAdded += OutputWarning;
            streams.Error.DataAdded += OutputError;
        }

        void OutputVerbose(object sender, DataAddedEventArgs e)
        {
            var record = (sender as PSDataCollection<VerboseRecord>) [e.Index];
            output.WriteLine($"[VERBOSE] {record.Message}");
        }

        void OutputWarning(object sender, DataAddedEventArgs e)
        {
            var record = (sender as PSDataCollection<WarningRecord>) [e.Index];
            output.WriteLine($"[WARN] {record.Message}");
        }
        void OutputError(object sender, DataAddedEventArgs e)
        {
            var record = (sender as PSDataCollection<ErrorRecord>) [e.Index];
            output.WriteLine($"[ERROR] {record}");
        }
    }
}