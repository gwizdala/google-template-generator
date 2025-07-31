# Google Template Generator
A Google Apps Script Project that enables you to generate custom Google Files (e.g. Slides, Docs, Sheets) from a template

[How-To Guide](https://gwizkid.com/posts/google-slides-presentation-generator/)

## Steps to Use

1. Create a Google Slide that has some variables in it marked by handlebars (`{{}}`) notation and slide sections marked by handlebars with the tag `_slide`, (e.g. `{{sectionName_slide}}`)
2. Create a Google Form that populates those variables and sections
3. Send those Form results to a Sheet that has a "Variables" and a "Sections" tab
4. Attach an Apps Script to the Form and add the files in the [`apps-script`](./apps-script/) folder
5. Run the `attachTrigger` function and accept the permissions

_Variables Example_

| Variable           | Default                                      | Form Key          | Form Index |
|--------------------|----------------------------------------------|-------------------|------------|
| TimeStamp          |                                              | Timestamp         |          0 |
| PresenterEmail     |                                              | Email Address     |          1 |
| FileName           | My Presentation                              | File Name         |          2 |
| FolderId           | myDefaultFolderId                            | Folder ID         |          3 |
| CompanyName        | Customer                                     | Company Name      |          4 |
| Sections           |                                              | Sections          |          6 |
| Status             |                                              | Generation Status |          7 |
| FileLinks          |                                              | File Link(s)      |          8 |
| TemplateTypes      | Slides                                       |                   |         -1 |
| SlideTemplateId    | mySlideTemplateId                            |                   |         -1 |
| DocumentTemplateId |                                              |                   |         -1 |
| Year               |                                              |                   |         -1 |

_Images Example_
| Variable           | Default                                      | Form Key          | Form Index |
|--------------------|----------------------------------------------|-------------------|------------|
| CompanyLogo        |                                              | Company Logo      |          5 |

_Sections Example_

| Form Value | Tag     | Title             |
|------------|---------|-------------------|
| Gizmos     | gizmos  | Gizmos are Good   |
| Widgets    | widgets | Widgets are Weird |
| Gadgets    | gadgets | Gadgets are Great |
