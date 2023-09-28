# PowerPointTemplates

This is a command line tool that allows to replace placeholders inside PowerPoint document with values provided as JSON file.

## Requirements
- Minimum `dotnet core 3.1`
- Microsoft Office


Nuget https://www.nuget.org/packages/PowerPointTemplates

## Hot to create placeholders

Use `Alt Text` to convert Text Box and Shapes into a placeholders. `Alt Text` is used to match values from the JSON file.

![image](https://github.com/cezarypiatek/OfficeTemplates/assets/7759991/d4373128-0283-47f9-88fa-042c7aa3cb6b)


## How to use it

Populate PowerPoint template with values and export slides as PNG images

```
dotnet tool run powerpointtemplates transform my_template.pptx -v data.json -e 'PNG' -o ./output
```

Populate PowerPoint template with values and save the result as a new document

```
dotnet tool run powerpointtemplates transform my_template.pptx -v data.json -s -f result.pptx -o ./output
```
