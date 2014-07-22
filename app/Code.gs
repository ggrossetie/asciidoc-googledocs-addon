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
      .setTitle('AsciiDoc Converter');

  DocumentApp.getUi().showModalDialog(ui, 'AsciiDoc Converter');
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
      var nextChild = elements[i + 1];
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
    if (typeof nextChild !== 'undefined' && !isEmptyText(child)) {
      if (nextChild.getType() == DocumentApp.ElementType.PARAGRAPH && !isEmptyText(nextChild)) {
        // Keep paragraph
        if (nextChild.getHeading() == DocumentApp.ParagraphHeading.NORMAL) {
          result = result + ' +';
        } else {
          result = result + '\n';
        }
      }
    }
  } else if (child.getType() == DocumentApp.ElementType.TABLE) {
    result = result + asciidocHandleTable(child);
  } else {
    result = result + child.getText();
  }
  return result;
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
  var isPlain = !isBold && !isItalic && !isUnderline;
  var isCode = isTextCode(text);
  // Prefix markup
  if (isUnderline) {
    // Underline doesn't play nice with others
    result = result + '+++<u>';
  } else {
    if (isBold) {
      result = result + new Array(numOccurence).join('*');
    }
    if (isItalic) {
      result = result + new Array(numOccurence).join('_');
    }
    if (isCode) {
      result = result + '+';
    }
  }
  // Content
  result = result + distinctContent;
  // Suffix markup
  if (!isUnderline) {
    if (isItalic) {
      result = result + new Array(numOccurence).join('_');
    }
    if (isBold) {
      result = result + new Array(numOccurence).join('*');
    }
    if (isCode) {
      result = result + '+';
    }
  } else {
    result = result + '</u>+++';
  }
  return result;
}

function asciidocHandleTitle(child) {
  var result = '';
  if (child.getHeading() == DocumentApp.ParagraphHeading.TITLE) {
    result = '= ' + child.getText();
  } else if (child.getHeading() == DocumentApp.ParagraphHeading.HEADING1) {
    result = '== ' + child.getText();
  } else if (child.getHeading() == DocumentApp.ParagraphHeading.HEADING2) {
    result = '=== ' + child.getText();
  } else if (child.getHeading() == DocumentApp.ParagraphHeading.HEADING3) {
    result = '==== ' + child.getText();
  } else if (child.getHeading() == DocumentApp.ParagraphHeading.HEADING4) {
    result = '===== ' + child.getText();
  } else if (child.getHeading() == DocumentApp.ParagraphHeading.HEADING5) {
    result =  '====== ' + child.getText();
  } else if (child.getHeading() == DocumentApp.ParagraphHeading.HEADING6) {
    result = '======= ' + child.getText();
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

function isTextCode(text) {
  return text.getFontFamily() == DocumentApp.FontFamily.CONSOLAS || text.getFontFamily() == DocumentApp.FontFamily.COURIER_NEW;
}