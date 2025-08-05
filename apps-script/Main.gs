/**
 * Template Generator - Main Function
 * Makes a copy of the Template(s) given criteria made by the user.
 * 
 * @author: @gwizdala
 */

//// TRIGGERS
// Run this once to get the appropriate trigger attached to your form and sheet
function attachTrigger(){
  const sheet = SpreadsheetApp.getActive();
  ScriptApp.newTrigger('submit')
    .forSpreadsheet(sheet)
    .onFormSubmit()
    .create();
}

/**
 * Generates the slide given what is returned from the form submission
 * 
 * @param {object} e the event trigger: https://developers.google.com/apps-script/guides/triggers/events#google_forms_events
 */
function submit(e) {
  const TODAY = new Date();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const variableMapping = getSheetDataAsObjects(spreadsheet, 'Variables');
  const imageMapping = getSheetDataAsObjects(spreadsheet, 'Images');
  const sectionMapping = getSheetDataAsObjects(spreadsheet, 'Sections');

  const sheet = spreadsheet.getActiveSheet();
  const range = e.range;
  const row = parseInt(range.getRow());
  const statusColumnIndex = variableMapping["Status"]?.formIndex;
  const linkColumnIndex = variableMapping["FileLinks"]?.formIndex;

  try {
    setCellValue(sheet, row, statusColumnIndex, GENERATION_STATUS.GENERATING);

    const namedValues = e.namedValues;
    Logger.log(`Form Entry: ${JSON.stringify(namedValues)}`);
    
    const values = e.values;
    // Figure out which templatized values to replace given what's in the form
    // This is to protect from malformed content in the form or in the constants file
    const mapValues = (formValues, mappingObject) => {
      const mapToObject = {};
      if (!!mappingObject) {  
        for (const templateVariable in mappingObject) {
          const templateObject = mappingObject[templateVariable];
          if (templateObject?.formIndex && templateObject.formIndex >= 0 && formValues[templateObject.formIndex]) {
            const formValue = formValues[templateObject.formIndex];
            const templateValue = !!formValue ? formValue : templateObject.default;
            mapToObject[templateVariable] = templateValue;
          } else {
            mapToObject[templateVariable] = templateObject.default;
          }
        }
      }

      return mapToObject;
    };

    const templateValues = mapValues(values, variableMapping);
    const templateImages = mapValues(values, imageMapping);

    // Arguments you can update
    templateValues["Year"] = TODAY.getFullYear();

    // Arguments you'll need
    const email = templateValues["PresenterEmail"];
    const name = templateValues["FileName"];
    const folderId = templateValues["FolderId"];
    const companyName = templateValues["CompanyName"];
    const sectionList = templateValues["Sections"].split(', ');

    // Arguments you could add
    // Checkbox input that gathers different file types to generate
    const templateTypes = templateValues["TemplateTypes"].split(', ');
    // Date inputs that gather start and end date
    // const startDate = templateValues["startDate"];
    // const endDate = templateValues["endDate"];
    // const templateValues["Duration"] = calculateDuration(startDate, endDate);


    // Generate each of the file types requested from the form
    let files = [];

    templateTypes.forEach((templateType) => {
      let template;
      switch(templateType) {
        case "Slides":
          template = new SlideTemplate(templateValues["SlideTemplateId"], `Presentation: ${name}`, folderId);
          break;
        // case "Doc": // Add other filetypes for people to select
        //   template = new DocumentTemplate(templateValues["DocumentTemplateId"], `Document: ${name}`, folderId);
        //   break;
        default:
      }

      if (template) {
        template.generate({
          variables: templateValues,
          images: templateImages,
          sections: sectionList,
          sectionMapping: sectionMapping
        });
        template.setOwnership(email);
        files.push({ name: template.getFileName(), link: template.getFileUrl() });
      }
    });
    
    setCellValue(sheet, row, statusColumnIndex, GENERATION_STATUS.SUCCESS);
    setCellValue(sheet, row, linkColumnIndex, files.map(file => file.link).join("\n"));
    sendSuccessEmail(email, files, companyName);
  } catch(err) {
    setCellValue(sheet, row, statusColumnIndex, GENERATION_STATUS.ERROR);
    Logger.log(`Generation Failed: ${err}`);
    sendErrorEmail(JSON.stringify(err));
  }
}
