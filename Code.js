function initializeSettings() {
  return {
    newLogo: DriveApp.getFileById('1LwxRWnT5si5K6QaoADOQeb4ureA9b8lw').getAs(
      'image/png'
    ),
    newColor: '#ca2226',
    newColorRGB: { red: 0.792, green: 0.133, blue: 0.204 },
    newPhone: '(987) 654-3210',
    newUrl: 'www.samplecompany.net',
    newEmail: 'info@samplecompany.net',
  };
}

function updateDocsStyle() {
  const settings = initializeSettings();
  const docsFolder = DriveApp.getFolderById(
    '1sh622pEPAnsx-QJn0K3FLZznxYs_oscT'
  );

  // Update documents in specified folder
  const docs = docsFolder.getFiles();
  const result = [['Name', 'URL', 'Status']];
  while (docs.hasNext()) {
    let status = 'Success';
    let doc;
    try {
      doc = DocumentApp.openByUrl(docs.next().getUrl());
      updateHeader(doc, settings);
      updateFooter(doc, settings);
      updateHeaderFooterTables(doc, settings);
      Logger.log(
        Utilities.formatString(
          'Updated document: %s (%s)',
          doc.getName(),
          doc.getUrl()
        )
      );
    } catch (e) {
      status = 'Failure: ' + e.message;
    } finally {
      result.push([doc.getName(), doc.getUrl(), status]);
    }
  }

  // Create report
  const report = SpreadsheetApp.create('Change Style Report');
  report
    .getActiveSheet()
    .getRange(1, 1, result.length, result[0].length)
    .setValues(result)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  DriveApp.getFileById(report.getId()).moveTo(docsFolder);

  // Send report link by email
  MailApp.sendEmail(
    Session.getActiveUser().getEmail(),
    'Docs Style Update Report',
    'Please find the report at ' + report.getUrl()
  );
}

function updateHeader(doc, settings) {
  const header = doc.getHeader();
  const originalHeaderTableIndex = 1;
  const originalHeaderTableNumColumns = 3;

  if (!header) {
    throw Error('Header not found.');
  }
  const originalHeaderTable = header.getChild(originalHeaderTableIndex);
  if (originalHeaderTable.getType() !== DocumentApp.ElementType.TABLE) {
    throw Error("The header's second element is not a table.");
  }
  if (
    originalHeaderTable.getRow(0).getNumCells() !==
    originalHeaderTableNumColumns
  ) {
    throw Error('The header table has an unexpected number of columns.');
  }

  // The logo is located in a paragraph element inside the first cell
  const originalHeaderLogo = originalHeaderTable
    .getCell(0, 0)
    .getChild(0)
    .getChild(0)
    .asInlineImage();

  // The new logo is added to the original logo's parent paragraph with the same dimensions
  originalHeaderLogo
    .getParent()
    .asParagraph()
    .insertInlineImage(0, settings.newLogo)
    .setWidth(originalHeaderLogo.getWidth())
    .setHeight(originalHeaderLogo.getHeight());

  originalHeaderLogo.removeFromParent();

  // Change color of company initials, which are in the first paragraph of the second cell
  const originalHeaderNameParagraph = originalHeaderTable
    .getCell(0, 1)
    .getChild(0)
    .asParagraph();
  originalHeaderNameParagraph
    .editAsText()
    .setForegroundColor(0, 0, settings.newColor)
    .setForegroundColor(7, 7, settings.newColor);

  // Change phone number, which is in the third paragraph of the third cell
  originalHeaderTable
    .getCell(0, 2)
    .getChild(2)
    .asParagraph()
    .setText(settings.newPhone);
}

function updateFooter(doc, settings) {
  const footer = doc.getFooter();
  const originalFooterTableIndex = 1;
  const originalFooterTableNumColumns = 3;
  const originalFooterLogoIndex = 2;

  if (!footer) {
    throw Error('Footer not found.');
  }
  const originalFooterTable = footer.getChild(originalFooterTableIndex);
  if (originalFooterTable.getType() !== DocumentApp.ElementType.TABLE) {
    throw Error("The footer's second element is not a table.");
  }
  if (
    originalFooterTable.getRow(0).getNumCells() !==
    originalFooterTableNumColumns
  ) {
    throw Error('The footer table has an unexpected number of columns.');
  }
  const originalFooterLogo = footer
    .getChild(originalFooterLogoIndex)
    .asParagraph()
    .getChild(0);
  if (originalFooterLogo.getType() !== DocumentApp.ElementType.INLINE_IMAGE) {
    throw Error('The footer does not contain a logo.');
  }

  // Change footer email
  originalFooterTable
    .getCell(0, 0)
    .getChild(0)
    .asParagraph()
    .setText(settings.newEmail);

  // Change footer phone number
  originalFooterTable
    .getCell(0, 1)
    .getChild(0)
    .asParagraph()
    .setText(settings.newPhone);

  // Change footer URL
  originalFooterTable
    .getCell(0, 2)
    .getChild(0)
    .asParagraph()
    .setText(settings.newUrl);

  // Replace original logo
  originalFooterLogo
    .getParent()
    .asParagraph()
    .insertInlineImage(0, settings.newLogo)
    .setWidth(originalFooterLogo.getWidth())
    .setHeight(originalFooterLogo.getHeight());
  originalFooterLogo.removeFromParent();
}

function updateHeaderFooterTables(doc, settings) {
  // Change border color using Docs API Service since it's not supported directly by Apps Script
  const docID = doc.getId();
  const apiDoc = Docs.Documents.get(docID);
  const apiHeaderID = apiDoc.documentStyle.defaultHeaderId;
  const apiHeaderContent = apiDoc.headers[apiHeaderID].content;
  const apiHeaderTableStart = apiHeaderContent[1].startIndex;
  const apiFooterID = apiDoc.documentStyle.defaultFooterId;
  const apiFooterContent = apiDoc.footers[apiFooterID].content;
  const apiFooterTableStart = apiFooterContent[1].startIndex;
  const newBorderStyle = {
    width: { magnitude: 2, unit: 'PT' },
    dashStyle: 'SOLID',
    color: { color: { rgbColor: settings.newColorRGB } },
  };

  const requests = {
    requests: [
      {
        updateTableCellStyle: {
          tableRange: {
            tableCellLocation: {
              tableStartLocation: {
                segmentId: apiDoc.documentStyle.defaultHeaderId,
                index: apiHeaderTableStart,
              },
              rowIndex: 0,
              columnIndex: 0,
            },
            rowSpan: 1,
            columnSpan: 3,
          },
          tableCellStyle: {
            borderBottom: newBorderStyle,
          },
          fields: 'borderBottom',
        },
      },
      {
        updateTableCellStyle: {
          tableRange: {
            tableCellLocation: {
              tableStartLocation: {
                segmentId: apiDoc.documentStyle.defaultFooterId,
                index: apiFooterTableStart,
              },
              rowIndex: 0,
              columnIndex: 0,
            },
            rowSpan: 1,
            columnSpan: 3,
          },
          tableCellStyle: {
            borderTop: newBorderStyle,
          },
          fields: 'borderTop',
        },
      },
    ],
  };
  Docs.Documents.batchUpdate(requests, docID);
}
