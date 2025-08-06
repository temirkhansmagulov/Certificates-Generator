function getIdFromUrl(url) { return url.match(/[-\w]{25,}/); }

function onOpen() {
  var ui = SpreadsheetApp.getUi();  //For convenience
    //Creates a button 'Generate Certificates' in the top panel of the Spreadsheet
    ui.createMenu('Generate Certificates')
      .addItem('Style 1 Certificate','certificateGen_style1')
      .addItem('Style 2 Certificate','certificateGen_style2')
      .addItem('Style 3 Certificate','certificateGen_style3')
      .addItem('Style 4 Certificate','certificateGen_style4')
      .addItem('Style 5 Certificate','certificateGen_style5')
      .addItem('Style 6 Certificate','certificateGen_style6')
      .addItem('Style 7 Certificate','certificateGen_style7')
      .addItem('Style 8 Certificate','certificateGen_style8')
      .addItem('Style 9 Certificate','certificateGen_style9')
      .addItem('Style 10 Certificate','certificateGen_style10')
    .addToUi();
}

////////////////////////
//Style 1 Certificate//
////////////////////////
function certificateGen_style1(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settingsSheet = ss.getSheetByName('Settings');
  var template_url = settingsSheet.getRange(6,1).getValue();//gets template url from the cell (6th row and 1st column)
  certificateGen(template_url);
}

////////////////////////////
//Style 2 Certificate//
////////////////////////////
function certificateGen_style2(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settingsSheet = ss.getSheetByName('Settings');
  var template_url = settingsSheet.getRange(7,1).getValue();//gets template url from the cell (7th row and 1st column)
  certificateGen(template_url);
}

////////////////////////
//Style 3 Certificate//
////////////////////////
function certificateGen_style3(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settingsSheet = ss.getSheetByName('Settings');
  var template_url = settingsSheet.getRange(8,1).getValue();//gets template url from the cell (6th row and 1st column)
  certificateGen(template_url);
}

////////////////////////////
//Style 4 Certificate//
////////////////////////////
function certificateGen_style4(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settingsSheet = ss.getSheetByName('Settings');
  var template_url = settingsSheet.getRange(9,1).getValue();//gets template url from the cell (7th row and 1st column)
  certificateGen(template_url);
}

////////////////////////
//Style 5 Certificate//
////////////////////////
function certificateGen_style5(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settingsSheet = ss.getSheetByName('Settings');
  var template_url = settingsSheet.getRange(10,1).getValue();//gets template url from the cell (6th row and 1st column)
  certificateGen(template_url);
}

////////////////////////////
//Style 6 Certificate//
////////////////////////////
function certificateGen_style6(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settingsSheet = ss.getSheetByName('Settings');
  var template_url = settingsSheet.getRange(11,1).getValue();//gets template url from the cell (7th row and 1st column)
  certificateGen(template_url);
}

////////////////////////
//Style 7 Certificate//
////////////////////////
function certificateGen_style7(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settingsSheet = ss.getSheetByName('Settings');
  var template_url = settingsSheet.getRange(12,1).getValue();//gets template url from the cell (6th row and 1st column)
  certificateGen(template_url);
}

////////////////////////////
//Style 8 Certificate//
////////////////////////////
function certificateGen_style8(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settingsSheet = ss.getSheetByName('Settings');
  var template_url = settingsSheet.getRange(13,1).getValue();//gets template url from the cell (7th row and 1st column)
  certificateGen(template_url);
}

////////////////////////
//Style 9 Certificate//
////////////////////////
function certificateGen_style9(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settingsSheet = ss.getSheetByName('Settings');
  var template_url = settingsSheet.getRange(14,1).getValue();//gets template url from the cell (6th row and 1st column)
  certificateGen(template_url);
}

////////////////////////////
//Style 10 Certificate//
////////////////////////////
function certificateGen_style10(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settingsSheet = ss.getSheetByName('Settings');
  var template_url = settingsSheet.getRange(15,1).getValue();//gets template url from the cell (7th row and 1st column)
  certificateGen(template_url);
}

//////////////////////////
//Certificates Generator//
//////////////////////////
function certificateGen(template_url){
  
  //Spreadsheet navigation
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var date = Utilities.formatDate(new Date(), "GMT+6", 'MMMM dd, yyyy [HH:mm:ss]');

  //Sheets by names
  var listSheet = ss.getSheetByName('List');
  var historySheet = ss.getSheetByName('History');
  var settingsSheet = ss.getSheetByName('Settings');

  var range =  listSheet.getActiveRange();
  
  var docsFolderId = getIdFromUrl(settingsSheet.getRange(2,1).getValue());
  var pdfFolderId = getIdFromUrl(settingsSheet.getRange(4,1).getValue());
  var template_id = getIdFromUrl(template_url);

  const DOCS_folder = DriveApp.getFolderById(docsFolderId);
  const PDF_folder = DriveApp.getFolderById(pdfFolderId);
  const Template = DriveApp.getFileById(template_id);

  for(var i=0; i<range.getNumRows();i++){
    var newTempFile = Template.makeCopy(listSheet.getRange(range.getRow()+i,1).getValue(), DOCS_folder);
    var OpenDoc = DocumentApp.openById(newTempFile.getId());
    var body = OpenDoc.getBody();
    console.log(body);
    body.replaceText("{Name and Surname}", listSheet.getRange(range.getRow()+i,1).getValue());
    body.replaceText("{Code}",listSheet.getRange(range.getRow()+i,2).getValue());
    body.replaceText("{Text}",listSheet.getRange(range.getRow()+i,3).getValue());
    body.replaceText("{Signer}",listSheet.getRange(range.getRow()+i,4).getValue());
    body.replaceText("{Value1}",listSheet.getRange(range.getRow()+i,5).getValue());
    body.replaceText("{Value2}",listSheet.getRange(range.getRow()+i,6).getValue());
    body.replaceText("{Value3}",listSheet.getRange(range.getRow()+i,7).getValue());
    body.replaceText("{Value4}",listSheet.getRange(range.getRow()+i,8).getValue());
    body.replaceText("{Value5}",listSheet.getRange(range.getRow()+i,9).getValue());
    body.replaceText("{Value6}",listSheet.getRange(range.getRow()+i,10).getValue());
    body.replaceText("{Value7}",listSheet.getRange(range.getRow()+i,11).getValue());
    body.replaceText("{Value8}",listSheet.getRange(range.getRow()+i,12).getValue());
    body.replaceText("{Value9}",listSheet.getRange(range.getRow()+i,13).getValue());
    body.replaceText("{Value10}",listSheet.getRange(range.getRow()+i,14).getValue());
    body.replaceText("{Value11}",listSheet.getRange(range.getRow()+i,15).getValue());
    body.replaceText("{Value12}",listSheet.getRange(range.getRow()+i,16).getValue());

    OpenDoc.saveAndClose();
    var BLOBPDF = newTempFile.getAs(MimeType.PDF);
    var pdfFile =  PDF_folder.createFile(BLOBPDF).setName(listSheet.getRange(range.getRow()+i,1).getValue());
    historySheet.getRange(historySheet.getLastRow()+1,1).setValue(listSheet.getRange(range.getRow()+i,1).getValue());
    historySheet.getRange(historySheet.getLastRow(),2).setValue(pdfFile.getUrl());
    historySheet.getRange(historySheet.getLastRow(),3).setValue(date);
  }

  return;
}

function templateDocument(){
    var newTempFile = Template.makeCopy(listSheet.getRange(range.getRow()+i,1).getValue(), DOCS_folder);
    var OpenDoc = DocumentApp.openById(newTempFile.getId());
    var body = OpenDoc.getBody();
    console.log(body);
    body.appendImage
    body.appendParagraph
    body.replaceText("{Name and Surname}", "{Name and Surname "+i+"}");
    body.replaceText("{Code}","{Code "+i+"}");
    body.replaceText("{Text}","{Text "+i+"}");
    body.replaceText("{Signer}","{Signer "+i+"}");
    body.replaceText("{Value1}","{Value1 "+i+"}");
    body.replaceText("{Value2}","{Value2 "+i+"}");
    body.replaceText("{Value3}","{Value3 "+i+"}");
    body.replaceText("{Value4}","{Value4 "+i+"}");
    OpenDoc.saveAndClose();
    var BLOBPDF = newTempFile.getAs(MimeType.PDF);
    var pdfFile =  PDF_folder.createFile(BLOBPDF).setName(listSheet.getRange(range.getRow()+i,1).getValue());
  return;
}

function testfunction (){
  var doc = DocumentApp.openById("14MEVZTWE9wdsxmjOI4Y_FqaTaz6sVmGO21nqvd4QFgI");
  //Copy rich text
  var body = doc.getBody();
  // console.log(body);
  var text = body.getText();
  //console.log(text);

  //Get number of children
  //Store children Text and Attributes as a Class
  //Append Children Text and set attributes to them one by one
  var child = body.getChild(0);
  //console.log(child);
  console.log(child.asText().getText());
  var childAtts = child.getAttributes();
  //console.log(childAtts);
  
  var style = attributeGetter(childAtts);
  //console.log(style);
  //Add page break
  body.appendPageBreak();
  var paragraph = body.appendParagraph(text);
  paragraph.setAttributes(style);
  //Append rich text
}

function attributeGetter(sourceAttributes){
  var style = {};
  style[DocumentApp.Attribute.BACKGROUND_COLOR] =     sourceAttributes[DocumentApp.Attribute.BACKGROUND_COLOR];
  style[DocumentApp.Attribute.BOLD] =                 sourceAttributes[DocumentApp.Attribute.BOLD];
  style[DocumentApp.Attribute.BORDER_COLOR] =         sourceAttributes[DocumentApp.Attribute.BORDER_COLOR];
  style[DocumentApp.Attribute.BORDER_WIDTH] =         sourceAttributes[DocumentApp.Attribute.BORDER_WIDTH];
  style[DocumentApp.Attribute.CODE] =                 sourceAttributes[DocumentApp.Attribute.CODE];
  style[DocumentApp.Attribute.FONT_FAMILY] =          sourceAttributes[DocumentApp.Attribute.FONT_FAMILY];
  style[DocumentApp.Attribute.FONT_SIZE] =            sourceAttributes[DocumentApp.Attribute.FONT_SIZE];
  style[DocumentApp.Attribute.FOREGROUND_COLOR] =     sourceAttributes[DocumentApp.Attribute.FOREGROUND_COLOR];
  style[DocumentApp.Attribute.GLYPH_TYPE] =           sourceAttributes[DocumentApp.Attribute.GLYPH_TYPE];
  style[DocumentApp.Attribute.HEADING] =              sourceAttributes[DocumentApp.Attribute.HEADING];
  style[DocumentApp.Attribute.HEIGHT] =               sourceAttributes[DocumentApp.Attribute.HEIGHT];
  style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = sourceAttributes[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT];
  style[DocumentApp.Attribute.INDENT_END] =           sourceAttributes[DocumentApp.Attribute.INDENT_END];
  style[DocumentApp.Attribute.INDENT_FIRST_LINE] =    sourceAttributes[DocumentApp.Attribute.INDENT_FIRST_LINE];
  style[DocumentApp.Attribute.INDENT_START] =         sourceAttributes[DocumentApp.Attribute.INDENT_START];
  style[DocumentApp.Attribute.ITALIC] =               sourceAttributes[DocumentApp.Attribute.ITALIC];
  style[DocumentApp.Attribute.LEFT_TO_RIGHT] =        sourceAttributes[DocumentApp.Attribute.LEFT_TO_RIGHT];
  style[DocumentApp.Attribute.LINE_SPACING] =         sourceAttributes[DocumentApp.Attribute.LINE_SPACING];
  style[DocumentApp.Attribute.LINK_URL] =             sourceAttributes[DocumentApp.Attribute.LINK_URL];
  style[DocumentApp.Attribute.LIST_ID] =              sourceAttributes[DocumentApp.Attribute.LIST_ID];
  style[DocumentApp.Attribute.MARGIN_BOTTOM] =        sourceAttributes[DocumentApp.Attribute.MARGIN_BOTTOM];
  style[DocumentApp.Attribute.MARGIN_LEFT] =          sourceAttributes[DocumentApp.Attribute.MARGIN_LEFT];
  style[DocumentApp.Attribute.MARGIN_RIGHT] =         sourceAttributes[DocumentApp.Attribute.MARGIN_RIGHT];
  style[DocumentApp.Attribute.MARGIN_TOP] =           sourceAttributes[DocumentApp.Attribute.MARGIN_TOP];
  style[DocumentApp.Attribute.MINIMUM_HEIGHT] =       sourceAttributes[DocumentApp.Attribute.MINIMUM_HEIGHT];
  style[DocumentApp.Attribute.NESTING_LEVEL] =        sourceAttributes[DocumentApp.Attribute.NESTING_LEVEL];
  style[DocumentApp.Attribute.PADDING_BOTTOM] =       sourceAttributes[DocumentApp.Attribute.PADDING_BOTTOM];
  style[DocumentApp.Attribute.PADDING_LEFT] =         sourceAttributes[DocumentApp.Attribute.PADDING_LEFT];
  style[DocumentApp.Attribute.PADDING_RIGHT] =        sourceAttributes[DocumentApp.Attribute.PADDING_RIGHT];
  style[DocumentApp.Attribute.PADDING_TOP] =          sourceAttributes[DocumentApp.Attribute.PADDING_TOP];
  style[DocumentApp.Attribute.PAGE_HEIGHT] =          sourceAttributes[DocumentApp.Attribute.PAGE_HEIGHT];
  style[DocumentApp.Attribute.PAGE_WIDTH] =           sourceAttributes[DocumentApp.Attribute.PAGE_WIDTH];
  style[DocumentApp.Attribute.SPACING_AFTER] =        sourceAttributes[DocumentApp.Attribute.SPACING_AFTER];
  style[DocumentApp.Attribute.SPACING_BEFORE] =       sourceAttributes[DocumentApp.Attribute.SPACING_BEFORE];
  style[DocumentApp.Attribute.STRIKETHROUGH] =        sourceAttributes[DocumentApp.Attribute.STRIKETHROUGH];
  style[DocumentApp.Attribute.UNDERLINE] =            sourceAttributes[DocumentApp.Attribute.UNDERLINE];
  style[DocumentApp.Attribute.VERTICAL_ALIGNMENT] =   sourceAttributes[DocumentApp.Attribute.VERTICAL_ALIGNMENT];
  style[DocumentApp.Attribute.WIDTH] =                sourceAttributes[DocumentApp.Attribute.WIDTH];
  return style;
}

function bodyCopier(body){
  var children = {};
  var childTexts = [];
  var childAttributes = [];
  return children;
}
