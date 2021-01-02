using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Typin;
using Typin.Attributes;
using Typin.Console;

namespace PowerPointTemplates
{
    [Command("dump", Description = "Generate values file base on the template document")]
    public class DumpPlaceholdersCommand : ICommand
    {
        [CommandParameter(0, Description = "Template document")]
        public string TemplateDocument { get; set; }

        [CommandOption("outputFile", 'o', Description = "Output file")]
        public string OutputFile { get; set; }

        public ValueTask ExecuteAsync(IConsole console)
        {
            var ppt = new ApplicationClass();
            var presentation = ppt.Presentations.Open(TemplateDocument, WithWindow: MsoTriState.msoFalse);
            try
            {
                DumpPlaceholders(presentation, OutputFile);
            }
            finally
            {
                presentation.Close();
                ppt.Quit();
            }

            return default;
        }

        private static void DumpPlaceholders(Presentation presentation, string file)
        {
            var dict = new Dictionary<string, PlaceholderReplacement>();
            for (var i = 1; i <= presentation.Slides.Count; i++)
                foreach (var placeholder in GetPlaceholders(presentation.Slides[i]))
                    dict[placeholder.Key] = placeholder;

            var serializedData = JsonSerializer.Serialize(dict.OrderBy(x => x.Key).Select(x => x.Value),
                new JsonSerializerOptions
                {
                    WriteIndented = true,
                    Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                });
            File.WriteAllText(file, serializedData);
        }

        private static IEnumerable<PlaceholderReplacement> GetPlaceholders(Slide templateSlide)
        {
            for (var i = 1; i <= templateSlide.Shapes.Count; i++)
                if (templateSlide.Shapes[i] is { } shape && shape.AlternativeText is { } key &&
                    string.IsNullOrWhiteSpace(key) == false)
                {
                    if (shape.Type == MsoShapeType.msoAutoShape)
                        yield return new PlaceholderReplacement(key, "");
                    else if (shape.Type == MsoShapeType.msoTextBox)
                        yield return new PlaceholderReplacement(key, shape.TextFrame?.TextRange?.Text);
                }
        }
    }
}