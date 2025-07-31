/**
 * Template Generator
 * Manages Template Functions - the parent object
 * 
 * @author: @gwizdala
 * 
 */
class Template {
  /**
   * Creates a copy of the template with the name provided at the desired location.
   * 
   * @param {integer} templateId The ID of the template to duplicate
   * @param {string} name The name of the new deck
   * @param {string} destinationId The ID of the folder to duplicate to.
   */
  constructor(templateId, name, destinationId) {
    const template = DriveApp.getFileById(templateId);
    const destination = DriveApp.getFolderById(destinationId);
    const copiedFile = template.makeCopy(name, destination);

    this.fileName = name;
    this.fileId = copiedFile.getId();
    this.file = DriveApp.getFileById(this.fileId);
  }

  /**
   * Returns the value in the tag format needed to find/replace
   * 
   * @param {string} tag the tag to tagify
   * @return the string wrapped with the appropriate tagging
   */
  tagify(tag) {
    return `{{${tag}}}`;
  }

 /**
 * Replaces the values in the template with the values provided by the form
 * 
 * @param {Array[object]} templateVariables the placeholder values to be replaced
 */
  replaceTemplateVariables(templateVariables) {}

  /**
   * Replaces the images in the template with the values provided by the form
   * 
   * @param {Array[object]} templateImages the images to be replaced
   */
  replaceTemplateImages(templateImages) {}

  /**
   * Hides content and moves to the appendix section given a list of tag to hide/move
   * 
   * @param {Array[string]} tags The regex tags in which to search from
   */
  hideContent(tags) {}

  /**
   * Show or Hide content based on requested Sections
   * 
   * @param {Array[string]} sections the list of requested sections
   * @param {object} mapping the mapped metadata to render
   */
  setSections(sections, mapping) {}

  /**
   * Removes extra data, like instructions, in the file
   */
  cleanup() {}

  /**
   * Sets the ownership of the slide to the submitter of the form
   * 
   * @param {string} email the email to set the ownership to
   */
  setOwnership(email) {
    DriveApp.getFileById(this.fileId).setOwner(email);
  }

  /**
   * Gets the url of the file
   * 
   * @return the string URL of the file
   */
  getFileUrl() {
    return this.file.getUrl();
  }

  /**
   * Gets the name of the file
   * 
   * @return the string name of the file
   */
  getFileName() {
    return this.fileName;
  }

  /**
   * Generates the template with all of the values loaded
   * 
   * @param {object} config the config object used to generate the template
   * @param {Array[object]} config.variables the placeholder values to be replaced
   * @param {Array[object]} config.images the placeholder images to be replaced
   * @param {Array[string]} config.sections the list of requested sections
   * @param {object} config.sectionMapping the mapped metadata to render
   * 
   * @return the file URL
   */
  generate(config) {
    if (config.variables) {
      this.replaceTemplateVariables(config.variables);
    }
    if (config.images) {
      this.replaceTemplateImages(config.images)
    }
    if (config.sections && config.sectionMapping) {
      this.setSections(config.sections, config.sectionMapping);
    }
    
    this.cleanup();
  }
}