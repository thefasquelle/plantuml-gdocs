function onOpen(e) {
  if (DocumentApp.getActiveDocument()) {
    DocumentApp.getUi().createAddonMenu()
      .addItem('Insert or update diagram - modless', 'showPlantUMLDialogModless')
      .addItem('Insert or update diagram - modal', 'showPlantUMLDialogModal')
      .addToUi();
  }
}

function onInstall(e) {
  onOpen(e);
}

function showPlantUMLDialogModal() {
  DocumentApp.getUi().showModalDialog(
    buildPlantumlDialog().setWidth(1000).setHeight(1000), 
    "editor"
  );
}

function showPlantUMLDialogModless() {
  DocumentApp.getUi().showModelessDialog(
    buildPlantumlDialog().setWidth(600).setHeight(600), 
    "editor"
  );
}

function buildPlantumlDialog() {
  var template = HtmlService.createTemplateFromFile('plantuml-editor');

  var selectedImage = checkSelectedInlineImage();
  if (selectedImage) {
    let desc = selectedImage.getAltDescription();
    if (desc.match(/@startuml(?:.|\n)+@enduml/gm)) {
      template.initialSource = desc;
    } else {
      throw "Selected image does include not a plantuml diagram."
    }
  } else {
    template.initialSource = "";
  }

  var html = template.evaluate()
                     .setTitle('PlantUML');
  
  return html;
}

function insertDiagram(base64Png, sourceCode) {
  const blob = Utilities.newBlob(Utilities.base64Decode(base64Png), 'image/png', 'plantuml-diagram.png');
  const doc = DocumentApp.getActiveDocument();
  const cursor = doc.getCursor();

  var image;

  if (cursor) {
    // try to insert at cursor position
    image = cursor.insertInlineImage(blob);
  } else {
    // if no cursor, it may be because an image is selected
    image = checkSelectedInlineImage();
    if (image) {
      // replace it
      image = replaceInlineImage(image, blob);
    } else {
      throw new Error("Cannot find cursor or selected image.");
    }
  }
  // CRITICAL: Google Docs sometimes strips embedded PNG metadata (like tEXt chunks) 
  // when it processes the image on its servers. We save the source code into the
  // image's "Description" (Alt Text) as a bulletproof backup.
  image.setAltDescription(sourceCode);
  image.setAltTitle("plantuml");
}

