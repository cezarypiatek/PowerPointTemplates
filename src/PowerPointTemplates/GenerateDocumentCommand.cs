using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Typin;
using Typin.Attributes;
using Typin.Console;

namespace PowerPointTemplates
{
    [Command("transform", Description = "Transform template document using values")]

    public class GenerateDocumentCommand : ICommand
    {
        [CommandParameter(0, Description = "Template document")]
        public string TemplateDocument { get; set; }

        [CommandOption("valuesFile", 'v', Description = "File with values for placeholders replacement")]
        public string ValuesFile { get; set; }

        public ValueTask ExecuteAsync(IConsole console)
        {
            var ppt = new ApplicationClass();
            var presentation = ppt.Presentations.Open(TemplateDocument, WithWindow: MsoTriState.msoFalse);
            var values = JsonSerializer.Deserialize<List<PlaceholderReplacement>>(File.ReadAllText(ValuesFile));
            try
            {
                for (var i = 1; i <= presentation.Slides.Count; i++)
                {
                    var templateSlide = presentation.Slides[i];
                    ReplacePlaceholders(templateSlide, values);
                    templateSlide.Export($"{Directory.GetCurrentDirectory()}\\slide_{i}.png", "PNG");
                }
            }
            finally
            {
                presentation.Close();
                ppt.Quit();
            }

            return default;
        }

        private static void ReplacePlaceholders(Slide templateSlide, List<PlaceholderReplacement> templateElements)
        {
            var dict = templateElements.ToDictionary(x => x.Key, x => x);
            for (var i = 1; i <= templateSlide.Shapes.Count; i++)
                if (templateSlide.Shapes[i] is { } shape && shape.AlternativeText is { } key &&
                    dict.TryGetValue(key, out var placeholderReplacement))
                    switch (shape.Type)
                    {
                        case MsoShapeType.msoTextBox:
                            templateSlide.Shapes[i].TextFrame.TextRange.Text = placeholderReplacement.Value;
                            break;
                        case MsoShapeType.msoAutoShape:
                            templateSlide.Shapes[i].Fill.UserPicture(placeholderReplacement.Value);
                            break;
                    }
        }
    }
}