/**
 * Return an InlineImage currently selected by the user.
 * Return null if no InlineImage is selected.
 */
function checkSelectedInlineImage() {
  const selection = DocumentApp.getActiveDocument().getSelection();
  
  if (!selection) {
    Logger.log("No selection found.");
    return null;
  }

  const rangeElements = selection.getRangeElements();
  for (const rangeElement of rangeElements) {
    const element = rangeElement.getElement();
    
    // Check if the selected element is an InlineImage
    if (element.getType() === DocumentApp.ElementType.INLINE_IMAGE) {
      const image = element.asInlineImage();
      Logger.log("InlineImage found!");
      return image;
    }
  }

  Logger.log("Selection does not contain an InlineImage.");
  return null;
}

/**
 * Replace the provided InlineImage with another InlineImage from png blob
 */
function replaceInlineImage(oldImage, newBlob, preserveAspectRatio=false) {
  // 1. Get the parent element (usually a Paragraph or ListItem)
  const parent = oldImage.getParent();
  
  // 2. Get the index position of the current image within the parent
  const index = parent.getChildIndex(oldImage);
  
  // 3. Store the old image's properties
  if (preserveAspectRatio) {
    var width = oldImage.getWidth();
    var height = oldImage.getHeight();
  }
  const altTitle = oldImage.getAltTitle();
  const altDescription = oldImage.getAltDescription();
  const linkUrl = oldImage.getLinkUrl();

  // 4. Insert the new image blob at the same position
  const newImage = parent.insertInlineImage(index, newBlob);

  // 5. Re-apply the old dimensions and metadata to the new image
  if (preserveAspectRatio) {
    newImage.setWidth(width).setHeight(height);
  }
  if (altTitle) newImage.setAltTitle(altTitle);
  if (altDescription) newImage.setAltDescription(altDescription);
  if (linkUrl) newImage.setLinkUrl(linkUrl);

  // 6. Remove the old image element from the document
  oldImage.removeFromParent();

  return newImage;
}
