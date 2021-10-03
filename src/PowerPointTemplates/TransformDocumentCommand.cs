using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
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
    [Command("transform", Description = "Transform template document using values")]

    public class TransformDocumentCommand : ICommand
    {
        [CommandParameter(0, Description = "Template document")]
        public string TemplateDocument { get; set; }

        [CommandOption("valuesFile", 'v', Description = "File with values for placeholders replacement", IsRequired = true)]
        public string ValuesFile { get; set; }

        [CommandOption("exportSlides", 'e', Description = "Export every slide into a separated file in a specified format like: JPG, PNG.", IsRequired = false)]
        public string ExportSlides { get; set; }


        [CommandOption("save", 's', Description = "Save transformed document", IsRequired = false)]
        public bool SaveTransformed { get; set; }
        
        [CommandOption("leaveOpen", 'l', Description = "Leave document open", IsRequired = false)]
        public bool LeaveOpen { get; set; }

        public ValueTask ExecuteAsync(IConsole console)
        {
            var ppt = new ApplicationClass();
            var presentation = ppt.Presentations.Open(TemplateDocument, WithWindow: LeaveOpen? MsoTriState.msoTrue : MsoTriState.msoFalse);
            var values = JsonSerializer.Deserialize<Dictionary<string, string>>(File.ReadAllText(ValuesFile));
            
            try
            {
                for (var i = 1; i <= presentation.Slides.Count; i++)
                {
                    var templateSlide = presentation.Slides[i];
                    ReplacePlaceholders(templateSlide, values);
                    if (string.IsNullOrWhiteSpace(ExportSlides) == false)
                    {
                        templateSlide.Export($"{Directory.GetCurrentDirectory()}\\{templateSlide.Name}.{ExportSlides.ToLower()}", ExportSlides.ToUpper());
                    }
                }

                if (SaveTransformed)
                {
                    var newName = $"{Directory.GetCurrentDirectory()}\\{Path.GetFileNameWithoutExtension(TemplateDocument)}_transformed.{Path.GetExtension(TemplateDocument)}";
                    presentation.SaveCopyAs(newName);
                }
            }
            finally
            {
                if(LeaveOpen == false)
                {
                    presentation.Close();
                    ppt.Quit();
                }
            }

            return default;
        }

        private static void ReplacePlaceholders(Slide templateSlide, Dictionary<string,string> dict)
        {
            for (var i = 1; i <= templateSlide.Shapes.Count; i++)
                if (templateSlide.Shapes[i] is { AlternativeText: { } key} shape && dict.TryGetValue(key, out var placeholderReplacement))
                    
                    switch (shape.Type)
                    {
                        case MsoShapeType.msoTextBox:
                            templateSlide.Shapes[i].TextFrame.TextRange.Text = placeholderReplacement;
                            break;
                        case MsoShapeType.msoAutoShape:
                            templateSlide.Shapes[i].Fill.UserPicture(placeholderReplacement);
                            break;
                    }
        }
    }
}