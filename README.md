# PowerPointTemplates

This is a command line tool that allows for replacing placeholders inside PowerPoint documents with values provided as JSON files.

## Requirements
- Minimum `dotnet core 3.1`
- Microsoft Office

## How to install

```
dotnet tool install --global PowerPointTemplates -no-cache --ignore-failed-sources
```

Nuget https://www.nuget.org/packages/PowerPointTemplates

## Hot to create placeholders

Use `Alt Text` to convert Text Box and Shapes into a placeholders. `Alt Text` is used to match values from the JSON file.

![image](https://github.com/cezarypiatek/PowerPointTemplates/assets/7759991/f7d0590b-762a-473a-8b95-7fb529c0127a)

JSON file with a related placeholder value:

```json
{
 "Survey_Link": "http://tinyurl.com/sample",
 "Survey_Code": "c:\\images\\code.png"
}
```

## How to use it

Populate PowerPoint template with values and export slides as PNG images

```
dotnet tool run powerpointtemplates transform my_template.pptx -v data.json -e 'PNG' -o ./output
```

Populate PowerPoint template with values and save the result as a new document

```
dotnet tool run powerpointtemplates transform my_template.pptx -v data.json -s -f result.pptx -o ./output
```
