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

/**
 * Gets the text the user has selected. If there is no selection,
 * this function displays an error message.
 *
 * @return {Array.<string>} The selected text.
 */
function getSelectedText() {
  var selection = DocumentApp.getActiveDocument().getSelection();
  if (selection) {
    var text = [];
    var elements = selection.getSelectedElements();
    for (var i = 0; i < elements.length; i++) {
      if (elements[i].isPartial()) {
        var element = elements[i].getElement().asText();
        var startIndex = elements[i].getStartOffset();
        var endIndex = elements[i].getEndOffsetInclusive();

        text.push(element.getText().substring(startIndex, endIndex + 1));
      } else {
        var element = elements[i].getElement();
        // Only translate elements that can be edited as text; skip images and
        // other non-text elements.
        if (element.editAsText) {
          var elementText = element.asText().getText();
          // This check is necessary to exclude images, which return a blank
          // text element.
          if (elementText != '') {
            text.push(elementText);
          }
        }
      }
    }
    if (text.length == 0) {
      throw 'Please select some text.';
    }
    return text;
  } else {
    throw 'Please select some text.';
  }
}

/**
 * Gets the stored user preferences for the origin and destination languages,
 * if they exist.
 *
 * @return {Object} The user's origin and destination language preferences, if
 *     they exist.
 */
function getPreferences() {
  var userProperties = PropertiesService.getUserProperties();
  var languagePrefs = {
    originLang: userProperties.getProperty('originLang'),
    destLang: userProperties.getProperty('destLang')
  };
  return languagePrefs;
}

/**
 * Gets the user-selected text and translates it from the origin language to the
 * destination language. The languages are notated by their two-letter short
 * form. For example, English is 'en', and Spanish is 'es'. The origin language
 * may be specified as an empty string to indicate that Google Translate should
 * auto-detect the language.
 *
 * @param {string} origin The two-letter short form for the origin language.
 * @param {string} dest The two-letter short form for the destination language.
 * @param {boolean} savePrefs Whether to save the origin and destination
 *     language preferences.
 * @return {string} The result of the translation.
 */
function runTranslation(origin, dest, savePrefs) {
  var text = getSelectedText();
  if (savePrefs == true) {
    var userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty('originLang', origin);
    userProperties.setProperty('destLang', dest);
  }

  var translated = [];
  for (var i = 0; i < text.length; i++) {
    translated.push(LanguageApp.translate(text[i], origin, dest));
  }

  return translated.join('\n');
}

/**
 * Replaces the text of the current selection with the provided text, or
 * inserts text at the current cursor location. (There will always be either
 * a selection or a cursor.) If multiple elements are selected, only inserts the
 * translated text in the first element that can contain text and removes the
 * other elements.
 *
 * @param {string} newText The text with which to replace the current selection.
 */
function insertText(newText) {
  var selection = DocumentApp.getActiveDocument().getSelection();
  if (selection) {
    var replaced = false;
    var elements = selection.getSelectedElements();
    if (elements.length == 1 &&
        elements[0].getElement().getType() ==
        DocumentApp.ElementType.INLINE_IMAGE) {
      throw "Can't insert text into an image.";
    }
    for (var i = 0; i < elements.length; i++) {
      if (elements[i].isPartial()) {
        var element = elements[i].getElement().asText();
        var startIndex = elements[i].getStartOffset();
        var endIndex = elements[i].getEndOffsetInclusive();

        var remainingText = element.getText().substring(endIndex + 1);
        element.deleteText(startIndex, endIndex);
        if (!replaced) {
          element.insertText(startIndex, newText);
          replaced = true;
        } else {
          // This block handles a selection that ends with a partial element. We
          // want to copy this partial text to the previous element so we don't
          // have a line-break before the last partial.
          var parent = element.getParent();
          parent.getPreviousSibling().asText().appendText(remainingText);
          // We cannot remove the last paragraph of a doc. If this is the case,
          // just remove the text within the last paragraph instead.
          if (parent.getNextSibling()) {
            parent.removeFromParent();
          } else {
            element.removeFromParent();
          }
        }
      } else {
        var element = elements[i].getElement();
        if (!replaced && element.editAsText) {
          // Only translate elements that can be edited as text, removing other
          // elements.
          element.clear();
          element.asText().setText(newText);
          replaced = true;
        } else {
          // We cannot remove the last paragraph of a doc. If this is the case,
          // just clear the element.
          if (element.getNextSibling()) {
            element.removeFromParent();
          } else {
            element.clear();
          }
        }
      }
    }
  } else {
    var cursor = DocumentApp.getActiveDocument().getCursor();
    var surroundingText = cursor.getSurroundingText().getText();
    var surroundingTextOffset = cursor.getSurroundingTextOffset();

    // If the cursor follows or preceds a non-space character, insert a space
    // between the character and the translation. Otherwise, just insert the
    // translation.
    if (surroundingTextOffset > 0) {
      if (surroundingText.charAt(surroundingTextOffset - 1) != ' ') {
        newText = ' ' + newText;
      }
    }
    if (surroundingTextOffset < surroundingText.length) {
      if (surroundingText.charAt(surroundingTextOffset) != ' ') {
        newText += ' ';
      }
    }
    cursor.insertText(newText);
  }
}

function myFunction() {
  var body = DocumentApp.getActiveDocument().getBody();
  var numChildren = body.getNumChildren();
  var paragraphs = body.getParagraphs();
  Logger.log('getNumChildren ' + numChildren);
  Logger.log('paragraphs ' + paragraphs);

  var asciidoc = '';
  
  for (var i = 0 ; i < numChildren; i++) {
    
    var child = body.getChild(i);
    Logger.log('child[' + i + ' ] ' + child);
    Logger.log('child[' + i + ' ].getType() ' + child.getType());
    
    if (child.getType() == DocumentApp.ElementType.PARAGRAPH) {
      if (child.getHeading() == DocumentApp.ParagraphHeading.TITLE) {
        asciidoc = asciidoc + "= " + child.getText();
      } else if (child.getHeading() == DocumentApp.ParagraphHeading.HEADING1) {
        asciidoc = asciidoc + "== " + child.getText();
      } else if (child.getHeading() == DocumentApp.ParagraphHeading.HEADING2) {
        asciidoc = asciidoc + "=== " + child.getText();
      } else if (child.getHeading() == DocumentApp.ParagraphHeading.HEADING3) {
        asciidoc = asciidoc + "==== " + child.getText();
      } else if (child.getHeading() == DocumentApp.ParagraphHeading.HEADING4) {
        asciidoc = asciidoc + "===== " + child.getText();
      } else if (child.getHeading() == DocumentApp.ParagraphHeading.HEADING5) {
        asciidoc = asciidoc + "====== " + child.getText();
      } else if (child.getHeading() == DocumentApp.ParagraphHeading.HEADING6) {
        asciidoc = asciidoc + "======= " + child.getText();
      } else if (child.getHeading() == DocumentApp.ParagraphHeading.NORMAL) {
        var text = child.editAsText();
        Logger.log('isBold() ' + text.isBold());

        var textAttributeIndices = text.getTextAttributeIndices();
        Logger.log('textAttributeIndices ' + textAttributeIndices);
        Logger.log('textAttributeIndices.length ' + textAttributeIndices.length);

        var content = child.getText();
        if (textAttributeIndices.length > 0) {
          for (var j = 0; j < textAttributeIndices.length; j++) {
            var offset = textAttributeIndices[j];
            var nextOffset = textAttributeIndices[j + 1];
            var distinctContent = content.substring(offset, nextOffset);
            if (text.isBold(offset)) {
              if (distinctContent.length > 1) {
                asciidoc = asciidoc + "*" + distinctContent + "*"; 
              } else {
                asciidoc = asciidoc + "**" + distinctContent + "**"; 
              }
            } else {
              asciidoc = asciidoc + distinctContent;
            }
          }
        } else {
          asciidoc = asciidoc + content;
        }

      } else {
        Logger.log('other ' + child.getText());
        asciidoc = asciidoc + child.getText();
      }
    }
    asciidoc = asciidoc + '\n';
  }

  return asciidoc;
}

