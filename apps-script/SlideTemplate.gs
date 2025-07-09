/**
 * Slide Template Generator
 * Makes a copy of the Presentation Template
 * 
 * @author: @gwizdala
 * 
 */
class SlideTemplate extends Template {
  constructor(templateId, name, destinationId) {
    super(templateId, name, destinationId);
    this.presentation = SlidesApp.openById(this.fileId);
    this.slides = this.presentation.getSlides();
  }

  replaceTemplateVariables(templateVariables) {
    for (const templateVariable in templateVariables) {
      const templateValue = templateVariables[templateVariable];
      this.presentation.replaceAllText(this.tagify(templateVariable), templateValue);
    }
  }

  hideContent(tags) {
    this.slides.forEach(slide => {
      // Get all shapes (includes text) in the current slides
      // And search for the key term
      // We are using getPageElements here because of an apps script bug in getShapes() which pulls page elements
      const shapes = slide.getPageElements();
      var shapeIndex = 0;
      var tagFound = false;
      while(shapeIndex < shapes.length && !tagFound) {
        const shape = shapes[shapeIndex];
        if (shape.getPageElementType().toString() == 'SHAPE' ) {
          const text = shape.asShape().getText();
          var tagIndex = 0;
          while(tagIndex < tags.length && !tagFound) {
            const tag = `${tags[tagIndex]}`;
            // Quick n' dirty way to ensure that the string is escaped properly
            const regexEscapers = {
              "{": "\\{",
              "}": "\\}",
              "-": "\\-",
              "_": "\\_"
            };
            const reEscape = new RegExp(Object.keys(regexEscapers).join("|"), "gi");

            const regex = tag.replace(reEscape, (matched) => {
              return regexEscapers[matched];
            });
            if (text.find(regex).length > 0) {
              // Found a slide to hide. Move to the end of the slides and hide it
              slide.move(this.slides.length);
              slide.setSkipped(true);
              tagFound = true;
            }
            tagIndex += 1;
          }
        }
        shapeIndex += 1;
      }
    });
  }

  setSections(sections, mapping) {
    const SECTION = "Section"; // The name of the section title
    const TOC = "TableOfContents"; // Where the section list is added
    const SLIDE_TAG_SUFFIX = '_slide'; // How we are identifying the tag to a slide
    
    let sectionList = '';
    const sectionArray = Array.isArray(sections) ? sections : [sections]; // handling single input entry

    sectionArray.forEach((section, index) => {
      let secObject = mapping[section];
      if (secObject) {
        // Create the Section List
        let sectionTitle = `${SECTION} ${index+1} - ${secObject.title}`; // e.g. "Section 1 - Title"
        sectionList += `${sectionList != '' ? '\n' : ''}${sectionTitle}`; // adds a newline if needed
        this.presentation.replaceAllText(this.tagify(secObject.tag), sectionTitle); // Replaces titles if they've been used
      }
    });

    // Display the Section List
    this.presentation.replaceAllText(this.tagify(TOC), sectionList);

    // Show/Hide slides
    const hiddenSections = Object.keys(mapping).filter((key) => !sectionArray.includes(key));
    var slideTags = [];
    
    hiddenSections.forEach((hiddenSection) => {
      const hsecObject = mapping[hiddenSection];
      // Build tags list for slides to hide/remove
      slideTags.push(this.tagify(`${hsecObject.tag}${SLIDE_TAG_SUFFIX}`));
      // Push title without requirement of UC number
      this.presentation.replaceAllText(this.tagify(hsecObject.tag), hsecObject.title);
    });
    this.hideContent(slideTags);
    // Wipe the slide tags from the slides
    Object.keys(mapping).forEach((option) => {
      this.presentation.replaceAllText(this.tagify(`${mapping[option].tag}${SLIDE_TAG_SUFFIX}`), '');
    });
  }

  cleanup() {
    // Delete the instructions slide
    this.slides[0].remove();
  }
}