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

namespace Cellmate.Cmdlets.Test
{
    class ScriptExecutor
    {
        public static int Execute(string path)
        {
            string script = File.ReadAllText(path);
            using (PowerShell ps = CreatePowerShell())
            {
                ps.AddScript(script);
                ps.Invoke();
                return ps.HadErrors ? 1 : 0;
            }
        }

        static PowerShell CreatePowerShell()
        {
            PowerShell powerShell = PowerShell.Create();
            powerShell.Streams.Information.DataAdded += OutputInformation;
            powerShell.Streams.Verbose.DataAdded += OutputVerbose;
            powerShell.Streams.Warning.DataAdded += OutputWarning;
            return powerShell;
        }

        static void OutputInformation(object sender, DataAddedEventArgs e)
        {
            var record = (sender as PSDataCollection<InformationRecord>) [e.Index];
            Console.WriteLine($"INFO: {record.MessageData}");
        }

        static void OutputVerbose(object sender, DataAddedEventArgs e)
        {
            var record = (sender as PSDataCollection<VerboseRecord>) [e.Index];
            Console.WriteLine($"VERBOSE: {record.Message}");
        }

        static void OutputWarning(object sender, DataAddedEventArgs e)
        {
            var record = (sender as PSDataCollection<WarningRecord>) [e.Index];
            Console.WriteLine($"WARN: {record.Message}");
        }
    }
}