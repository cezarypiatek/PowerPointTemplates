using System;
using System.Threading.Tasks;
using Typin;

namespace PowerPointTemplates
{
    class Program
    {
        public static async Task<int> Main() =>
            await new CliApplicationBuilder()
                .AddCommandsFromThisAssembly()
                .Build()
                .RunAsync();
        
    }
}
