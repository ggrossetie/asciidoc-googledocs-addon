/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 */
function onOpen() {
  DocumentApp.getUi().createAddonMenu()
      .addItem('Start', 'showSidebar')
      .addToUi();
}

/**
 * Runs when the add-on is installed.
 */
function onInstall() {
  onOpen();
}

/**
 * Opens a sidebar in the document containing the add-on's user interface.
 */
function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('AsciiDoc Export');
  DocumentApp.getUi().showSidebar(ui);
}

function asciidocify() {
  var body = DocumentApp.getActiveDocument().getBody();
  var numChildren = body.getNumChildren();
  var paragraphs = body.getParagraphs();
  var asciidoc = '';
  for (var i = 0 ; i < numChildren; i++) {
    var child = body.getChild(i);
    if (child.getType() == DocumentApp.ElementType.PARAGRAPH) {
      asciidoc = asciidoc + asciidocHandleTitle(child);
      asciidoc = asciidoc + asciidocHandleText(child);
      if (i + 1 < numChildren && !isEmptyText(child)) {
        var nextChild = body.getChild(i + 1);
        if (nextChild.getType() == DocumentApp.ElementType.PARAGRAPH && !isEmptyText(nextChild)) {
          // Keep paragraph
          asciidoc = asciidoc + " +";
        }
      }
    } else {
      asciidoc = asciidoc + child.getText();
    }
    asciidoc = asciidoc + '\n';
  }
  return asciidoc;
}

function isEmptyText(child) {
  return child.getText().replace(/^\s+|\s+$/g, '').length == 0;
}

function asciidocHandleText(child) {
  var result = '';
  if (child.getHeading() == DocumentApp.ParagraphHeading.NORMAL) {
    var text = child.editAsText();
    var textAttributeIndices = text.getTextAttributeIndices();
    var content = child.getText();
    if (textAttributeIndices.length > 0) {
      for (var j = 0; j < textAttributeIndices.length; j++) {
        var offset = textAttributeIndices[j];
        var nextOffset = textAttributeIndices[j + 1];
        var distinctContent = content.substring(offset, nextOffset);
        result = result + asciidocHandleFontStyle(text, offset, distinctContent);
      }
    } else {
      result = content;
    }
  }
  return result;
}


function asciidocHandleFontStyle(text, offset, distinctContent) {
  var result = '';
  var numOccurence = (distinctContent.length > 1 ? 1 : 2) + 1;
  var isBold = text.isBold(offset);
  var isItalic = text.isItalic(offset);
  var isUnderline = text.isUnderline(offset);
  var isPlain = !isBold && !isItalic && !isUnderline
  // Prefix markup
  if (isUnderline) {
    // Underline doesn't play nice with others
    result = result + "+++<u>";
  } else {
    if (isBold) {
      result = result + Array(numOccurence).join("*");
    }
    if (isItalic) {
      result = result + Array(numOccurence).join("_");
    }
  }
  // Content
  result = result + distinctContent;
  // Suffix markup
  if (!isUnderline) {
    if (isItalic) {
      result = result + Array(numOccurence).join("_");
    }
    if (isBold) {
      result = result + Array(numOccurence).join("*");
    }
  } else {
    result = result + "</u>+++";
  }
  return result;
}

function asciidocHandleTitle(child) {
  var result = '';
  if (child.getHeading() == DocumentApp.ParagraphHeading.TITLE) {
    result = "= " + child.getText();
  } else if (child.getHeading() == DocumentApp.ParagraphHeading.HEADING1) {
    result = "== " + child.getText();
  } else if (child.getHeading() == DocumentApp.ParagraphHeading.HEADING2) {
    result = "=== " + child.getText();
  } else if (child.getHeading() == DocumentApp.ParagraphHeading.HEADING3) {
    result = "==== " + child.getText();
  } else if (child.getHeading() == DocumentApp.ParagraphHeading.HEADING4) {
    result = "===== " + child.getText();
  } else if (child.getHeading() == DocumentApp.ParagraphHeading.HEADING5) {
    result =  "====== " + child.getText();
  } else if (child.getHeading() == DocumentApp.ParagraphHeading.HEADING6) {
    result = "======= " + child.getText();
  }
  return result;
}
