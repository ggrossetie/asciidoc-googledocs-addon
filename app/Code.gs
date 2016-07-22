/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 */
function onOpen() {
  DocumentApp.getUi().createAddonMenu()
      .addItem('Convert all', 'showDialogConvertAll')
      .addItem('Convert selection', 'showDialogConvertSelection')
      .addToUi();
}

/**
 * Runs when the add-on is installed.
 */
function onInstall() {
  onOpen();
}

function showDialogConvertAll() {
  CacheService.getPrivateCache().put('convertMode', 'all');
  showDialog();
}

function showDialogConvertSelection() {
  CacheService.getPrivateCache().put('convertMode', 'selection');
  showDialog();
}

/**
 * Opens a dialog containing the add-on's user interface.
 */
function showDialog() {
  var ui = HtmlService.createHtmlOutputFromFile('Dialog')
      .setWidth(400)
      .setHeight(500)
      .setTitle('AsciiDoc Processor');

  DocumentApp.getUi().showModalDialog(ui, 'AsciiDoc Processor');
}

function asciidocify() {
  var selection = DocumentApp.getActiveDocument().getSelection();
  var elements = [];
  if (selection && 'selection' == CacheService.getPrivateCache().get('convertMode')) {
    var rangeElements = selection.getRangeElements();
    for (var i = 0; i < rangeElements.length; i++) {
      elements.push(rangeElements[i].getElement());
    }
  } else {
    var body = DocumentApp.getActiveDocument().getBody();
    for (var i = 0; i < body.getNumChildren(); i++) {
      elements.push(body.getChild(i));
    }
  }
  var asciidoc = '';
  var insideCodeBlock = false;
  var elementsLength = elements.length;
  for (var i = 0 ; i < elementsLength; i++) {
    var child = elements[i];
    var nextChild = undefined;
    if (i + 1 < elementsLength) {
      nextChild = elements[i + 1];
    }
    // Handle code block
    var isCurrentCode = isTextCode(child.editAsText());
    if (!insideCodeBlock) {
      if (isCurrentCode) {
        if (typeof nextChild !== 'undefined' && isTextCode(nextChild.editAsText())) {
          // Start code block
          asciidoc = asciidoc + '----\n';
          insideCodeBlock = true;
        }
      }
    }
    if (insideCodeBlock) {
      asciidoc = asciidoc + child.getText();
      if (typeof nextChild !== 'undefined') {
        var isNextChildCode = isTextCode(nextChild.editAsText());
        if (!isNextChildCode) {
          // End code block
          asciidoc = asciidoc + '\n----';
          insideCodeBlock = false;
        }
      } else {
        // End code block
        asciidoc = asciidoc + '\n----';
        insideCodeBlock = false;
      }
    } else {
      asciidoc = asciidoc + asciidocHandleChild(child, i, nextChild);
    }
    asciidoc = asciidoc + '\n';
  }
  return asciidoc;
}

function isEmptyText(child) {
  return child.getText().replace(/^\s+|\s+$/g, '').length == 0;
}

function asciidocHandleChild(child, i, nextChild) {
  var result = '';
  if (child.getType() == DocumentApp.ElementType.PARAGRAPH) {
    result = result + asciidocHandleTitle(child);
    result = result + asciidocHandleText(child);
    if (child.getHeading() == DocumentApp.ParagraphHeading.NORMAL
        && !isEmptyText(child)
        && typeof nextChild !== 'undefined'
        && nextChild.getType() == DocumentApp.ElementType.PARAGRAPH
        && !isEmptyText(nextChild)) {
      // Keep paragraph
      if (nextChild.getHeading() == DocumentApp.ParagraphHeading.NORMAL) {
        result = result + ' +';
      } else {
        result = result + '\n';
      }
    }
  } else if (child.getType() == DocumentApp.ElementType.TABLE) {
    result = result + asciidocHandleTable(child);
  } else if (child.getType() == DocumentApp.ElementType.LIST_ITEM) {
    result = result + asciidocHandleList(child);
  } else {
    result = result + child.getText();
  }
  return result;
}

function asciidocHandleText(child) {
  var result = '';
  if (child.getType() == DocumentApp.ElementType.TEXT || child.getHeading() == DocumentApp.ParagraphHeading.NORMAL) {
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
  var isStrikethrough = text.isStrikethrough(offset);
  var isLink = false;
  var linkURL = text.getLinkUrl(offset);
  var htmlBuf = ''
  // FIXME: getTextAttributeIndices doesn't split on different fonts,
  // makeing this almost useless
  var isCode = isTextCode(text);
  // Prefix markup
  if (linkURL !== null) {
    isLink = true;
    result = result + new Array(numOccurence).join(linkURL + '[');
  }
  if (isUnderline && !isLink) {
    htmlBuf += '<u>'; // or asciidoc.css class: underline
  }
  if (isStrikethrough) {
    htmlBuf += '<s>'; // or asciidoc.css class: line-through
  }
  if (htmlBuf !== '') {
    result += '+++' + htmlBuf + '+++';
    htmlBuf = '';
  }
  if (isBold) {
    result = result + new Array(numOccurence).join('*');
  }
  if (isItalic) {
    result = result + new Array(numOccurence).join('_');
  }
  if (isCode) {
    result = result + '+';
  }
  // Content
  result += distinctContent;
  // Suffix markup
  if (isLink) {
    result = result + new Array(numOccurence).join(']');
  }
  if (isCode) {
    result = result + '+';
  }
  if (isItalic) {
    result = result + new Array(numOccurence).join('_');
  }
  if (isBold) {
    result = result + new Array(numOccurence).join('*');
  }
  if (isStrikethrough) {
    htmlBuf += '</s>';
  }
  if (isUnderline && !isLink) {
    htmlBuf += '</u>';
  }
  if (htmlBuf !== '') {
    result += '+++' + htmlBuf + '+++';
  }
  return result;
}

function asciidocHandleTitle(child) {
  var result = '';
  var headingLevel;
  if (child.getText() && child.getText().trim() !== '') {
    if (child.getHeading() == DocumentApp.ParagraphHeading.TITLE) {
      headingLevel = 1;
    } else if (child.getHeading() == DocumentApp.ParagraphHeading.HEADING1) {
      headingLevel = 2;
    } else if (child.getHeading() == DocumentApp.ParagraphHeading.HEADING2) {
      headingLevel = 3;
    } else if (child.getHeading() == DocumentApp.ParagraphHeading.HEADING3) {
      headingLevel = 4;
    } else if (child.getHeading() == DocumentApp.ParagraphHeading.HEADING4) {
      headingLevel = 5;
    } else if (child.getHeading() == DocumentApp.ParagraphHeading.HEADING5) {
      headingLevel = 6;
    } else if (child.getHeading() == DocumentApp.ParagraphHeading.HEADING6) {
      headingLevel = 7;
    }
    if (typeof headingLevel !== 'undefined') {
      result = new Array(headingLevel + 1).join('=') + ' ' + child.getText() + '\n';
    }
  }
  return result;
}

function asciidocHandleTable(child) {
  var result = '';
  if (child.getType() == DocumentApp.ElementType.TABLE) {
    var numRows = child.getNumRows();
    if (numRows > 0) {
      result = result + '|===\n';
      for (var rowIndex = 0; rowIndex < numRows; rowIndex++) {
        var tableRow = child.getRow(rowIndex);
        var numCells = tableRow.getNumCells();
        for (var cellIndex = 0; cellIndex < numCells; cellIndex++) {
          var tableCell = tableRow.getCell(cellIndex);
          result = result + '|';
          result = result + asciidocHandleChild(tableCell.getChild(0));
        }
        if (rowIndex == 0) {
          result = result + '\n';
        }
        result = result + '\n';
      }
      result = result + '|===';
    }
  }
  return result;
}

function asciidocHandleList(child) {
  var result = '';
  var listSize = child.getNumChildren();
  if (listSize == 1) {
    result = asciidocHandleText(child.getChild(0));
    var listSyntax;
    if (child.getGlyphType() == DocumentApp.GlyphType.BULLET
        || child.getGlyphType() == DocumentApp.GlyphType.HOLLOW_BULLET
        || child.getGlyphType() == DocumentApp.GlyphType.SQUARE_BULLET) {
      listSyntax = new Array(child.getNestingLevel() + 2).join('*');
    } else {
      listSyntax = new Array(child.getNestingLevel() + 2).join('.');
    }
    result = ' ' + listSyntax + ' ' + result;
  } else {
    // Should never happen?
    result = child.getText();
  }
  return result;
}

/** Guess if the text is code by looking at the font family. */
function isTextCode(text) {
  // Things will be better if Google Fonts can tell us about a font
  var i, fontFamily = text.getFontFamily(), /* Now it returns a string! */
  monospaceFonts = ['Consolas', 'Courier New', 'Source Code Pro'];
  if (fontFamily === null) {
    return false; // Handle null early.. It means multiple values.
  }
  // See ES7 Array.prototype.includes(elem, pos)
  for (i = 0; i < monospaceFonts.length; i++) {
    if (fontFamily === monospaceFonts[i]) {
      return true;
    }
  }
  // Last Try: Assume it's mono if it ends with ' Mono'.
  // This works for all Google Fonts as of 2016-10-21.
  // See ES6 String.prototype.endsWith(str, pos).
  return fontFamily.indexOf(' Mono') === fontFamily.length - 5;
}
