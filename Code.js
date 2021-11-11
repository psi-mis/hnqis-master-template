// ------------------------------------------------------------------------------------------------
// Generate the Menu
// ------------------------------------------------------------------------------------------------

function onOpen() {
  
  resetData();
  
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();  
  var entries = [
    {
      name : "To File",
      functionName : "call_generateXMLToFile"
    }
    ,{
      name: "To Sheet",
      functionName : "call_generateXMLToSheet"
    }
  ];

  var entries2 = [
    {
      name: "To Sheet",
      functionName: "call_generateTemplateToSheet"
    }
  ];
  spreadsheet.addMenu("XML", entries);
  spreadsheet.addMenu("HNQIS 1.6 Template", entries2);  
};

// Reset data in 'Create XML' sheet when the template is opened

function resetData()
{
   var ss = SpreadsheetApp.getActiveSpreadsheet();
   var xmlSheet = ss.getSheetByName("Create XML");
   // HNQIS 1.6
   var hnqis16 = ss.getSheetByName("HNQIS 1.6");
   hnqis16.clear();

   var rowNo = xmlSheet.getLastRow();
   
   // Set property 'Public Sharing' of all DEs  as 'Read Only' as default
   xmlSheet.getRange("B7").setValue("r-------");
  
   // Set property 'Store zero data values' of all DEs  as 'false' as default
   xmlSheet.getRange("B8").setValue("FALSE");
  
   // Set property 'Domain Type' of all DEs  as 'TRACKER' as default
   xmlSheet.getRange("B9").setValue("TRACKER");
  
   // Set empty values of all DEs 'User Group Sharing' as default
   xmlSheet.getRange("A11:B" + rowNo ).setValue("");
}

// ------------------------------------------------------------------------------------------------
// Global variable
// ------------------------------------------------------------------------------------------------

var mappingDataIdx = 5;

var uidAttrRowIdx = 2;
var dataRowIdx = 3;

var uidIdx = 0;
var deNameIdx = 1;
var deShortNameIdx = 2;
var deCodeIdx = 3;
var deDescriptionIdx = 4;
var deFormNameIdx = 5;
var aggOperatorIdx = 6;
var optionSetIdx = 7;
var tabNameIdx = 8;
var headerIdx = 9;
var deTypeIdx = 10;
var questionIdx = 11;
var questionOrderIdx = 12;
var questionHideIdx = 13;
var questionHideGroupIdx = 14;
var questionMatchIdx = 15;
var questionMatchGroupIdx = 16;
var numWIdx = 17;
var denWIdx = 18;
var compIndicatorIdx = 19;
var imgUrlIdx = 20;
var videoUrlIdx = 21;
var questionParentsIdx = 22;
var questionParentOptsIdx = 23;
var rowTagIdx = 24;
var columnTagIdx = 25;
var tabGroupIdx = 26;
var compulsoryIdx = 27;



// ------------------------------------------------------------------------------------------------
// Generate XML Data
// ------------------------------------------------------------------------------------------------

/** 
*
* Generate XML Import data as a file. 
* Users can download this XML Import file in 'My driver' folder in their Google driver
* 
**/

function call_generateXMLToFile()
{
   generateXML( true );
}

/** 
* 
* Generate XML file and put in 'XML Generated' sheet
* 
**/

function call_generateXMLToSheet()
{ 
   showProgressMessage('Cleaning data in XML Generated sheet  ...');
  
   var ss = SpreadsheetApp.getActiveSpreadsheet();
   var xmlSheet = ss.getSheetByName("XML Generated");
   if( xmlSheet.getLastRow() > 0 )
   {
     xmlSheet.insertRowBefore(1);
     xmlSheet.deleteRows( uidAttrRowIdx, xmlSheet.getLastRow() - 1 );
   }
  
   generateXML( false );
}


/**
* 
* Generate XML import data
* @params isFile 
*           True is generating into XML file. Users can download XML file in 'My driver' folder in their Google driver
*           False is generating to 'XML Generate' sheet
*
**/

function generateXML( isFile ) 
{  
  if( runValidation() )
  {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    showProgressMessage('Loading meta data ...');
    
    // Load meta data in 'Mapping' sheet
    
    loadMappingData();
   
    
    showProgressMessage('Generating XML data. Please wait ...');
    
    
    var deSheet = ss.getSheetByName("DE master");
    var dataParamSheet = ss.getSheetByName("Create XML");
    var data = deSheet.getDataRange().getValues();
	
    if( data.length > dataRowIdx )
    {
      var arrayOptionSet = _optionSets;
      var arrQuestionType = _questionTypes;
      var arrDeTypes = _deTypes;
      var defaultCatCombo = _defaultCatCombo;
      
      var extraData = dataParamSheet.getDataRange().getValues();
      var publicAccess = extraData[6][1];
      var zeroIsSignificant = extraData[7][1].toLowerCase();
      var domainType = extraData[8][1];
      
      // Get uids of atttributes
      
      var uid_TabGroup = getTabGroupValue( data, uidAttrRowIdx );
	  var uid_RowTag = getRowTagValue( data, uidAttrRowIdx );
      var uid_ColumnTag = getColumnTagValue( data, uidAttrRowIdx );
      var uid_Tab = getTabValue( data, uidAttrRowIdx );
      var uid_Header = getHeaderValue( data,uidAttrRowIdx );
      var uid_Order = getOrderValue( data, uidAttrRowIdx );
      var uid_DeType = getDETypeValue( data, uidAttrRowIdx );     
      var uid_QType = getQTypeValue( data, uidAttrRowIdx );
      var uid_QHide = getQHideValue( data, uidAttrRowIdx );
      var uid_QHideGroup = getQHideGroupValue( data, uidAttrRowIdx );
      var uid_QMatch = getQMatchValue( data, uidAttrRowIdx );
      var uid_QMatchGroup = getQMatchGroupValue( data, uidAttrRowIdx );
      var uid_NumQ = getNumWValue( data, uidAttrRowIdx );
      var uid_DenW = getDenWValue( data, uidAttrRowIdx );
      var uid_CompIndicator = getCompIndicatorValue( data, uidAttrRowIdx );
      var uid_ImgUrl = getImageUrlValue( data, uidAttrRowIdx );
      var uid_VideoUrl = getVideoUrlValue( data, uidAttrRowIdx );
      var uid_QuestionParents = getQuestionParentsValue( data, uidAttrRowIdx );
      var uid_QuestionParentOpts = getQuestionParentOptsValue( data, uidAttrRowIdx );
      
      var url, xmlroot, xmlDEs, xmlSheet;
      var xmlData = new Array();
      var curDate = getCurrentDate();
      
      if( isFile ) 
      {
        url = XmlService.getNamespace("http://dhis2.org/schema/dxf/2.0");
        xmlroot = XmlService.createElement('metaData',url);
        xmlDEs = XmlService.createElement('dataElements');
      }
      else
      {
        xmlSheet = ss.getSheetByName("XML Generated");
        xmlSheet.getRange('A1').setValue("<?xml version='1.0' encoding='UTF-8'?>");
        xmlSheet.getRange('A2').setValue("<metaData xmlns='http://dhis2.org/schema/dxf/2.0' created='" + curDate + "'>");
        xmlSheet.getRange('A3').setValue("<dataElements>"); 
      }
      
      
      // Generate the Group Sharing XML Note
      var hasGroupSharing = ( extraData.length >= 10 );     
      
      for( var i=dataRowIdx; i<data.length; i++ )
      {
        var uid = data[i][uidIdx];
        var code = data[i][deCodeIdx];
        var name = data[i][deNameIdx];
        var formName = data[i][deFormNameIdx];
        var shortName = data[i][deShortNameIdx];
        var description = data[i][deDescriptionIdx];
        
        var aggregationType= data[i][aggOperatorIdx];          
        if( aggregationType == "" )
        {
          aggregationType = "SUM";
        }
        
        var optionSet = data[i][optionSetIdx];
        optionSet = arrayOptionSet[optionSet];
        if( optionSet == undefined )
        {
          optionSet = "";
        }
        
        if( shortName == "" )
        {
          shortName = name.substring(0, 45 );
        }
        
        var xmlDE = XmlService.createElement('dataElement');
        if( code != "" )
        {
          xmlDE.setAttribute('code', code);
        }
        xmlDE.setAttribute('name', name);
        xmlDE.setAttribute('created', curDate);
        xmlDE.setAttribute('lastUpdated', curDate);
        xmlDE.setAttribute('shortName', shortName); 
		
		var xmlDefaultCatCombo = XmlService.createElement("categoryCombo");
		xmlDefaultCatCombo.setAttribute( 'id', defaultCatCombo );
		xmlDE.addContent( xmlDefaultCatCombo );
        
        if( uid != "" )
        {
          xmlDE.setAttribute('id', uid);
        }
        
        xmlDE.addContent(generateNote( 'publicAccess', publicAccess ));
        xmlDE.addContent(generateNote( 'zeroIsSignificant', zeroIsSignificant ));
        xmlDE.addContent(generateNote( 'domainType', domainType ));        
        xmlDE.addContent(generateNote( 'aggregationType', aggregationType  )); // This information is not needed in version 2.20 
        
        
        if( formName != "" )
        {
          xmlDE.addContent(generateNote( 'formName', formName ));
        }
        
        if( description != "")
        {
          xmlDE.addContent(generateNote( 'description', description ));
        }
        
        if( optionSet != "" )
        {
          var xmlOptionSet = XmlService.createElement( "optionSet" );
          xmlOptionSet.setAttribute( "id", optionSet );
          xmlDE.addContent( xmlOptionSet );
        }
        
        // Group Sharing
        var groupSharing =  generateGroupSharingsNote( extraData );
        if( groupSharing != "" )
        {
           xmlDE.addContent( groupSharing );
        }
        
        // Attributes
        
        var value_TabGroup = getTabGroupValue( data,i ) + "";
        var value_RowTag = getRowTagValue( data,i ) + "";
        var value_ColumnTag = getColumnTagValue( data,i ) + "";
        var value_Tab = getTabValue( data,i ) + "";
        var value_Header = getHeaderValue( data, i ) + "";
        var value_Order = getOrderValue( data, i ) + "";  
        var value_DeType = getDETypeValue( data, i ) + ""; 
        var value_QType = getQTypeValue( data, i ) + "";
        var value_QHide = getQHideValue( data, i ) + "";
        var value_QHideGroup = getQHideGroupValue( data, i ) + "";
        var value_QMatch = getQMatchValue( data, i ) + "";
        var value_QMatchGroup = getQMatchGroupValue( data, i ) + "";
        var value_NumQ = getNumWValue( data, i ) + "";
        var value_DenW = getDenWValue( data, i ) + "";
        var value_CompIndicator = getCompIndicatorValue( data, i ) + "";
        var value_ImageUrl = getImageUrlValue( data, i ) + "";
        var value_VideoUrl = getVideoUrlValue( data, i ) + "";
        var value_QuestionParents = getQuestionParentsValue( data, i ) + "";
        var value_QuestionParentOpts = getQuestionParentOptsValue( data, i ) + "";
        
        // 'Q. Type' value
        
        var defaultType = arrQuestionType["default"].split(";");
        var type = defaultType[1];
        value_QType = arrQuestionType[value_QType];
        if( value_QType == undefined )
        {
          value_QType = "";
        }
        else
        {
           var qType = value_QType.split(";");
           value_QType = qType[0]; // uid
           type = ( qType[1]=="" ) ? type : qType[1];
        }
        
        // 'DE Type' value
        
		var _value_DeType = "";

		if( value_DeType == '' )
		{
			value_DeType = _value_DeType;
		}
		if( value_DeType !== '' && arrDeTypes[value_DeType] == undefined )
		{
		  value_DeType = "{" + value_DeType + "} : [not exist in Mapping]";
		}
		else
		{
		  value_DeType = arrDeTypes[value_DeType];
		}
		
		if( value_DeType == undefined )
		{
			value_DeType = "";
		}
        
        xmlDE.addContent(generateNote( 'valueType', type )); // the attribute name is "type" in version 2.20
        
        if( /* value_TabGroup != "" || */
			value_RowTag != "" || value_ColumnTag != "" 
		   || value_Tab != "" || value_Header != "" || value_Order != "" 
           || value_QType != "" || value_QHide != "" || value_QHideGroup != "" || value_QMatch != "" 
           || value_QMatchGroup != "" || value_NumQ != "" || value_DenW != "" || value_CompIndicator != "" 
           || value_DeType != "" )
        {
          var xmlAttributeValues = XmlService.createElement('attributeValues');       
          
          generateAttributeValue( xmlAttributeValues, value_Tab, uid_Tab, curDate );
          generateAttributeValue( xmlAttributeValues, value_Header, uid_Header, curDate );
          generateAttributeValue( xmlAttributeValues, value_Order, uid_Order, curDate );
          generateAttributeValue( xmlAttributeValues, value_DeType, uid_DeType, curDate );
          generateAttributeValue( xmlAttributeValues, value_QType, uid_QType, curDate );
          generateAttributeValue( xmlAttributeValues, value_QHide, uid_QHide, curDate );
          generateAttributeValue( xmlAttributeValues, value_QHideGroup, uid_QHideGroup, curDate );
          generateAttributeValue( xmlAttributeValues, value_QMatch, uid_QMatch, curDate );
          generateAttributeValue( xmlAttributeValues, value_QMatchGroup, uid_QMatchGroup, curDate );
          generateAttributeValue( xmlAttributeValues, value_NumQ, uid_NumQ, curDate );
          generateAttributeValue( xmlAttributeValues, value_DenW, uid_DenW, curDate );
          generateAttributeValue( xmlAttributeValues, value_CompIndicator, uid_CompIndicator, curDate );
          generateAttributeValue( xmlAttributeValues, value_ImageUrl, uid_ImgUrl, curDate );
          generateAttributeValue( xmlAttributeValues, value_VideoUrl, uid_VideoUrl, curDate );
          generateAttributeValue( xmlAttributeValues, value_QuestionParents, uid_QuestionParents, curDate );
          generateAttributeValue( xmlAttributeValues, value_QuestionParentOpts, uid_QuestionParentOpts, curDate );
		  generateAttributeValue( xmlAttributeValues, value_TabGroup, uid_TabGroup, curDate );			
		  generateAttributeValue( xmlAttributeValues, value_RowTag, uid_RowTag, curDate );
		  generateAttributeValue( xmlAttributeValues, value_ColumnTag, uid_ColumnTag, curDate );
		  
          
          xmlDE.addContent(xmlAttributeValues);
          
        }
        
        if( isFile )
        {
          xmlDEs.addContent(xmlDE);
        }
        else
        {
          var rowIdx = i - 3;
          var colData = new Array();
          colData[0] = convertXMLToString( xmlDE );
          xmlData[rowIdx]= colData;
        }
        
      }
      
      if( isFile )
      {
        xmlroot.addContent( xmlDEs );
        var document = XmlService.createDocument(xmlroot);
        var xmldoc = XmlService.getPrettyFormat().format(document);
        
        //  Save the XML output to a document using the name of the active spreadsheet. 
        // - Verify that it's saved in the current directory -
        
        var folder = DriveApp.getFileById(SpreadsheetApp.getActive().getId());
        var fileName = ss.getName() + '.xml';
        var file = DriveApp.createFile(fileName, xmldoc, 'application/xml');
        
        Browser.msgBox("XML file was created successfully. You can download the file in '" + file.getUrl() );
      }
      else
      {
        xmlSheet.getRange( 4,1, data.length - dataRowIdx ).setValues( xmlData );
        
        var rowNo = data.length + 1;
        xmlSheet.getRange('A' + rowNo).setValue( "</dataElements>");
        
        rowNo++;
        xmlSheet.getRange('A' + rowNo).setValue( "</metaData>");
        
        ss.setActiveSheet(xmlSheet);
        Browser.msgBox("Export data successfully.");
      }
      
    }
    else
    {
      Browser.msgBox("No data to export.");
    }
    
  }
  else
  {
    Browser.msgBox("There are invalid data in DE master sheet. Please check red cells in this sheet.");
  }
  
}

// ------------------------------------------------------------------------------------------------
// Dialog box for processing/error/warming
// ------------------------------------------------------------------------------------------------


// Dialog box - Progressing message

function showProgressMessage( message ) {
	
	var html = HtmlService.createHtmlOutputFromFile('message.html')
      .setWidth(500)
      .setHeight(80);
	  
	html.append('<p>' + message + '</p>');
	SpreadsheetApp.getUi().showModalDialog(html, 'Processing ...');
	
}


// --------------------------------------------------------------------------------------------------------------
// Get Attribute Values from sheet
// --------------------------------------------------------------------------------------------------------------

function getTabGroupValue( data, rowIdx )
{  
  return data[rowIdx][tabGroupIdx];
}

function getTabValue( data, rowIdx )
{  
  return data[rowIdx][tabNameIdx];
}

function getHeaderValue( data, rowIdx )
{
  return data[rowIdx][headerIdx];
}

function getOrderValue( data, rowIdx )
{
  return data[rowIdx][questionOrderIdx];
}

function getDETypeValue( data, rowIdx )
{
	return data[rowIdx][deTypeIdx];
}

function getQTypeValue( data, rowIdx )
{
  return data[rowIdx][questionIdx];
}

function getQHideValue( data, rowIdx )
{
  return data[rowIdx][questionHideIdx];
}

function getQHideGroupValue( data, rowIdx )
{
  return data[rowIdx][questionHideGroupIdx];
}

function getQMatchValue( data, rowIdx )
{
  return data[rowIdx][questionMatchIdx];
}

function getQMatchGroupValue( data, rowIdx )
{
  return data[rowIdx][questionMatchGroupIdx];
}

function getNumWValue( data, rowIdx )
{
  return data[rowIdx][numWIdx];
}

function getDenWValue( data, rowIdx )
{
  return data[rowIdx][denWIdx];
}

function getCompIndicatorValue( data, rowIdx )
{
  return data[rowIdx][compIndicatorIdx];
}

function getRowTagValue( data, rowIdx )
{
  return data[rowIdx][rowTagIdx];
}

function getColumnTagValue( data, rowIdx )
{
  return data[rowIdx][columnTagIdx];
}

function getImageUrlValue( data, rowIdx )
{				
  return data[rowIdx][imgUrlIdx];
}

function getVideoUrlValue( data, rowIdx )
{
  return data[rowIdx][videoUrlIdx];
}

function getQuestionParentsValue( data, rowIdx )
{
  return data[rowIdx][questionParentsIdx];
}

function getQuestionParentOptsValue( data, rowIdx )
{
  return data[rowIdx][questionParentOptsIdx];
}


// Create simple XML note with noteName and text
function generateNote( noteName, text )
{
  var xmlNote = XmlService.createElement(noteName);
  xmlNote.setText(text);
  return xmlNote;
}

// Load Mapping data, include option sets, q.Type attribute value, programs

function loadMappingData()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Mapping");
  
  
  var rowNo = sheet.getLastRow(); 
  var optionSetList = {};
  var questionTypeList = {};
  var invQuestionList = {};
  var programList = {};
  var deTypeList = {};
  var invDeTypeList = {};
  var defaultCatCombo = "";
  
  var data = sheet.getDataRange().getValues();
  
  if( rowNo>=4 )
  {
    for (var i=4; i<data.length; i++) {
      
      // Get 'Option Set' UID from 'Mapping'
      var name = data[i][0];
      if( name != "" )
      {
        var uid = data[i][1];     
        optionSetList[name] = uid;
      }
      
      // 'Question Type' && 'Inv  Question Type'
      name = data[i][4];
      if( name != "" )
      {
        var uid = data[i][3];  
        var deType = data[i][5]; 
        questionTypeList[name] = uid + ";" + deType;
        
        invQuestionList[uid + " - " + deType] = name;
      }
      
      // 'Program'      
      name = data[i][14];
      if( name != "" )
      {
        var uid = data[i][15];
        programList[name] = uid;
      }
      
      // 'deType'
      uid = data[i][20];
      if( uid != "" )
      {
        var name = data[i][21];
        deTypeList[name] = uid;
        invDeTypeList[uid] = name;
      }
	  
	  // defaultCategoryUID
	  uid = data[i][23];
	  if( uid != "" )
	  {
		  defaultCatCombo = uid;
	  }
      
    }
  }
  
  // Save mapping data into Global variables
 
   _optionSets = optionSetList;
   _questionTypes = questionTypeList;
   _invQuestionTypes = invQuestionList;
   _programs = programList;
   _deTypes = deTypeList;
   _invDeTypes = invDeTypeList;
   _defaultCatCombo = defaultCatCombo;
  
}

// Generate an Attribue Value Note

function generateAttributeValue( xmlParentNote, value, uid, curDate )
{
  if( value!= "" )
  {  
    var xmlAttributeValue = XmlService.createElement("attributeValue");
    xmlAttributeValue.setAttribute( 'created', curDate );
    xmlAttributeValue.setAttribute( 'lastUpdated', curDate );
    
    var xmlAttribute = XmlService.createElement("attribute");
    xmlAttribute.setAttribute( 'id', uid );
    xmlAttributeValue.addContent( xmlAttribute );
    
    var xmlValue = XmlService.createElement("value");
    xmlValue.setText( value );
    xmlAttributeValue.addContent( xmlValue );
        
   
    
    if( uid == 'olcVXnDPG1U' )
    {
       Logger.log( convertXMLToString( xmlAttributeValue ) );
    }
    
    xmlParentNote.addContent( xmlAttributeValue );
  }
  
}

function generateGroupSharingsNote( extraData )
{
  var xmlGroupSharings = XmlService.createElement("userGroupAccesses");
  var hasGroupSharing = false;
  
  for( var i=10;i<extraData.length; i++ )
  {
    var access = extraData[i][1];
    var uid = extraData[i][2];
   
    if( access!="" && uid!="" && uid!="#N/A"){
      hasGroupSharing = true;
      generateGroupSharingNote( xmlGroupSharings, uid, access );
    }
  }  
  
  
  return ( hasGroupSharing ) ? xmlGroupSharings : "" ;
}

// Generate an Group Sharing Note
function generateGroupSharingNote( xmlParentNote, userGroupUid, access )
{      
    var xmlGroupSharing = XmlService.createElement("userGroupAccess");
   
    var xmlAccess = XmlService.createElement("access");
    xmlAccess.setText( access );
    xmlGroupSharing.addContent( xmlAccess );
    
    var xmlUserGroup = XmlService.createElement("userGroupUid");
    xmlUserGroup.setText( userGroupUid );
    xmlGroupSharing.addContent( xmlUserGroup );
  
    var xmlUid = XmlService.createElement("uid");
    xmlUid.setText( userGroupUid );
    xmlGroupSharing.addContent( xmlUid );
        
    xmlParentNote.addContent( xmlGroupSharing );
  
}

// ------------------------------------------------------------------------------------------------
// Check data in "DE master" sheet
// ------------------------------------------------------------------------------------------------

var sheetValid = true;
var warmingColor = "#ff5050";
var errorColor = "#66ccff";
	
function runValidation() 
{
  showProgressMessage("Validate data. Please wait .....");
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName("DE master");

  
  var numRows = sourceSheet.getLastRow();
  var numCols = sourceSheet.getLastColumn();
  sourceSheet.getRange( dataRowIdx, 1, numRows,numCols ).setBackground("white");
  
  
  // Get the active sheet and info about it
    
  
  if( !checkDataValidationColumn( sourceSheet ) )
  {
    return false;
  }
  
    /** List the columns you want to check by number (A = 1) **/
    var CHECK_COLUMNS = [uidIdx, deCodeIdx, deNameIdx, deShortNameIdx, deFormNameIdx];  
    
    
    // Create the temporary working sheet
    var newSheet = ss.getSheetByName("FindDupes");
    
    if( newSheet !== null )
    {
      ss.deleteSheet(newSheet);
    }
    
    newSheet = ss.insertSheet("FindDupes");
    newSheet.hideSheet();
    
    
    // Copy the desired rows to the FindDupes sheet
    for (var i = 0; i < CHECK_COLUMNS.length; i++) {
      var sourceRange = sourceSheet.getRange(1,CHECK_COLUMNS[i] + 1,numRows);
      var nextCol = newSheet.getLastColumn() + 1;
      sourceRange.copyTo(newSheet.getRange(1,nextCol,numRows));
    }
    
    // Find duplicates in the FindDupes sheet and color them in the main sheet
    var sheetValid = true;
    var data = newSheet.getDataRange().getValues();
    for (i = dataRowIdx; i < data.length; i++) 
    {
		// Check max length of data
        var rowValid = true;
		var uid = data[i][0];
		var code = data[i][1];
		var name = data[i][2];
		var shortName = data[i][3];
		var formName = data[i][4];
	
      if( name == "" 
         || ( uid != "" && uid.length != 11 ) 
        || ( code != "" && code.length > 50 ) 
        || ( name != "" && name.length > 230 )
        || ( shortName != "" && shortName.length > 50 )
        || ( formName != "" && formName.length > 230 ) )
        {
          sheetValid = false;
          rowValid = false;
          sourceSheet.getRange(i+1,1,1,numCols).setBackground("#FF8080");
        }
      
      
      if( rowValid )
      {
        for (j = i+1; j < data.length; j++) {
          for (var k = 0; k< CHECK_COLUMNS.length - 1; k++) {
            // Check duplicate for uid, code, name, shortName
            // Don't check formName ( the last item in CHECK_COLUMNS )
            if( data[i][k]!="" && data[j][k]!="" && data[i][k] == data[j][k]) 
            {
              sheetValid = false;
              rowValid = false;
              sourceSheet.getRange(i+1,1,1,numCols).setBackground("#FF8080");
              sourceSheet.getRange(j+1,1,1,numCols).setBackground("#FF8080");
              break;
            }
            
          }
        }
      }
    }
    
    // Remove the FindDupes temporary sheet
    ss.deleteSheet(newSheet);
  
    return sheetValid;
};


function checkDataValidationColumn( sheet )
{
   var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("DE master");
  
  var cells = sheet.getDataRange().getBackgrounds();
  var rows = cells.length;
  var cols = cells[0].length;
  
   for (var i = dataRowIdx; i < rows; i++){
    for (var j = 0; j < cols; j++){
      if (cells[i][j] == '#ff0000'){ // first color to change
        return false;
      }
    }
   }
  return true;
}

// ------------------------------------------------------------------------------------------------
// Utilities
// ------------------------------------------------------------------------------------------------

function getCurrentDate()
{
  var curDate = new Date();
  return formatDate( curDate );
}

function formatDate( date )
{
  var month = date.getMonth() + 1;
  var days = date.getDate();
  
  if( month < 10 ) {
   month = '0' + month; 
  }
  
  if( days < 10 ) {
   days = '0' + days; 
  }
  
  return date.getYear() + "-" + month + "-" + days;
}

function convertXMLToString( xmlDoc )
{
  return XmlService.getPrettyFormat().format(xmlDoc).replace(/\n/g, ' ');
}

// ------------------------------------------------------------------------------------
// Download data elements from server
// ------------------------------------------------------------------------------------


// Copy a new sheet
function cloneGoogleSheet( sourceSheet ) {
  
  var name = sourceSheet.getName() + " - Backup";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = sourceSheet.copyTo(ss);
  
  /* Before cloning the sheet, delete any previous copy */
  var old = ss.getSheetByName(name);
  if (old) ss.deleteSheet(old); // or old.setName(new Name);
  
  SpreadsheetApp.flush(); // Utilities.sleep(2000);
  sheet.setName(name);
 
}

function removeDataRows( dataSheet )
{
  if( dataSheet.getLastRow() > dataRowIdx ){
    dataSheet.insertRowAfter( dataRowIdx );
    dataSheet.deleteRows( dataRowIdx + 2, dataSheet.getLastRow() - dataRowIdx - 1 ); 
  }
}

// Connect server and retreive data
// Fill data into "DE master" sheet
function retrieveDataFromServer( uname, upass, servername, dataURL )
{
  showProgressMessage("Accessing server ...");
  
  var url = servername + dataURL;
  var user = uname + ":" + upass;
  
  // PULL URL from Config spreadsheet to allow the user to change which server will be used
  
  var options =
      {
        "contentType" : "application/json",
        "method" : "get",
        "headers": {
          "Authorization": "Basic " + Utilities.base64Encode(user)
        }
      };
  
  var response;
  var valid = true;
  try
  {
    response = UrlFetchApp.fetch(url, options);
  } catch (e) {
    valid = false;
	
	if( e.message.indexOf("\"httpStatus\":\"Unauthorized\"") >= 0 )
    {
		Browser.msgBox("Username/Password is wrong. ");
	}
	else
	{
		Browser.msgBox("Invalid URL. " + e.message);
	}
	
    showPasswordDialog( dataURL );
  }
  
  if(valid)
  {
    writeData( response );
  }

}

function checkDuplicateAndGetDataElementListInStages( jsonProgram )
{
	programStages = jsonProgram.programStages;
	var resultList = [];
	var idx = 0;
  //Logger.log(programStages);
	
	for( var i=0; i<programStages.length; i++ )
	{
		try
		{
			var psDEs = programStages[i].programStageDataElements;
			for( var j=0; j<psDEs.length; j++ )
			{
				resultList[idx] = psDEs[j]; //.dataElement;
        //var test = resultList[idx];
        //var test2 = psDEs[j].compulsory;
        //Logger.log(psDEs[j].dataElement.name); // .push("compulsory"+ "=" + psDEs[j].compulsory)
				idx++;
			}
		}
		catch( e )
		{
			Browser.msgBox("Error : " + e.message + " in line "+ e.lineNumber);
		}
	}
	
	return resultList;
}

function writeData( response )
{
 
  var jsonData = JSON.parse( response );
  var dataElementList = checkDuplicateAndGetDataElementListInStages( jsonData );
  
  if( dataElementList.length > 0 )
  {
    try
	{
		var ss = SpreadsheetApp.getActiveSpreadsheet();
	  
		showProgressMessage("Loading meta data ...");
		loadMappingData();
		
		var mappingSheet = ss.getSheetByName("Mapping");    
		var arrayOptionSet = _optionSets;
		var arrQuestionType = _questionTypes;
		var arrInvQuestionType = _invQuestionTypes;
		var arrDeTypes = _deTypes;
		var arrInvDeType = _invDeTypes;
		
	 
		var errorCells = { "optionSet" : [], "qType" : [], "deTypes" : [] };
		
		
		var dataSheet = ss.getSheetByName("DE master");
	  
	    // Create backup sheet for dataSheet
		showProgressMessage("Back up data in the \"DE master\" sheet. Please wait ..." );
		cloneGoogleSheet( dataSheet );
	  
		// Remove data rows from dataSheet if any
	    showProgressMessage("Cleaning data in the \"DE master\" sheet. Please wait ..." );
		removeDataRows( dataSheet );
		
	    showProgressMessage("Processing data and populating the \"DE master\" sheet with " + dataElementList.length + " data elements. Please wait ...");
	  
		  
		var exportSheetData = dataSheet.getDataRange().getValues();
		var exportSheetColNo = dataSheet.getLastColumn();
	  
		var exportData = new Array();
		
		var idxOpt = arrayOptionSet.length + 1;
		
		var uid_TabGroup = getTabGroupValue( exportSheetData, uidAttrRowIdx );
		var uid_RowTag = getRowTagValue( exportSheetData, uidAttrRowIdx );
		var uid_ColumnTag = getColumnTagValue( exportSheetData, uidAttrRowIdx );
		var uid_Tab = getTabValue( exportSheetData, uidAttrRowIdx );
		var uid_Header = getHeaderValue( exportSheetData, uidAttrRowIdx );
		var uid_Order = getOrderValue( exportSheetData, uidAttrRowIdx );  
		var uid_DeType = getDETypeValue( exportSheetData, uidAttrRowIdx );      
		var uid_QType = getQTypeValue( exportSheetData, uidAttrRowIdx );
		var uid_QHide = getQHideValue( exportSheetData, uidAttrRowIdx );
		var uid_QHideGroup = getQHideGroupValue( exportSheetData, uidAttrRowIdx );
		var uid_QMatch = getQMatchValue( exportSheetData, uidAttrRowIdx );
		var uid_QMatchGroup = getQMatchGroupValue( exportSheetData, uidAttrRowIdx );
		var uid_NumQ = getNumWValue( exportSheetData, uidAttrRowIdx );
		var uid_DenW = getDenWValue( exportSheetData, uidAttrRowIdx );
		var uid_CompIndicator = getCompIndicatorValue( exportSheetData, uidAttrRowIdx );
		var uid_ImgUrl = getImageUrlValue( exportSheetData, uidAttrRowIdx );
		var uid_VideoUrl = getVideoUrlValue( exportSheetData, uidAttrRowIdx );
		var uid_QuestionParents = getQuestionParentsValue( exportSheetData, uidAttrRowIdx );
		var uid_QuestionParentOpts = getQuestionParentOptsValue( exportSheetData, uidAttrRowIdx );
	  
		for( var i=0; i<dataElementList.length; i++ )
		{
		  var cols = new Array( exportSheetColNo );
		  
		  // Set default values for row
		  for (var j=0; j<exportSheetColNo; j++){
			cols[j] = "";
		  } 
		  
		  var idx = 5 + i;
		  var data = dataElementList[i];
		  var valueType = data.dataElement.valueType; // specify parent
		  
		  cols[uidIdx] = data.dataElement.id; // specify parent
		  cols[deNameIdx] = data.dataElement.name; // specify parent
		  cols[deShortNameIdx] = data.dataElement.shortName; // specify parent
		  cols[aggOperatorIdx] = data.dataElement.aggregationType; // specify parent
		  if(  data.dataElement.description !== undefined )
		  {
			 cols[deDescriptionIdx] = data.dataElement.description; // specify parent
		  }
		  
		  if( data.dataElement.code !== undefined )
		  {
			cols[deCodeIdx] = data.dataElement.code; // specify parent
		  }

		  if( data.dataElement.formName !== undefined )
		  {
			cols[deFormNameIdx] = data.dataElement.formName; // specify parent
		  }
		  
		  if( data.dataElement.optionSet !== undefined )
		  {
			var optionSetId = data.dataElement.optionSet.id; // specify parent
			var optionSetName = data.dataElement.optionSet.name; // specify parent
			if( arrayOptionSet[optionSetName] == undefined )
			{
			  optionSetName = optionSetName + " : [not exist in Mapping]";         
			  
			  errorCells.optionSet.push( i + dataRowIdx + 1);
			  
			  //arrayOptionSet[optionSetName] = optionSetId;
			  
			  //idxOpt++;   
			  //mappingSheet.getRange("A" + idxOpt ).setValue( optionSetName ); 
			  //mappingSheet.getRange("B" + idxOpt ).setValue( optionSetId );
			}
			
			cols[optionSetIdx] = optionSetName;
		  }
		  
		  if( data.dataElement.attributeValues.length > 0 )
		  {
			for( var j=0; j<data.dataElement.attributeValues.length; j++ )
			{
			  var attribueValue = data.dataElement.attributeValues[j];
			  
			  if( attribueValue.attribute.id == uid_TabGroup )
			  {
				cols[tabGroupIdx] = attribueValue.value;
			  } 
			  else if( attribueValue.attribute.id == uid_RowTag )
			  {
				cols[rowTagIdx] = attribueValue.value;
			  }
			  else if( attribueValue.attribute.id == uid_ColumnTag )
			  {
				cols[columnTagIdx] = attribueValue.value;
			  }
			  else if( attribueValue.attribute.id == uid_Tab )
			  {
				cols[tabNameIdx] = attribueValue.value;
			  }
			  else if( attribueValue.attribute.id == uid_Header )
			  {
				cols[headerIdx] = attribueValue.value;
			  }
			  else if( attribueValue.attribute.id == uid_Order )
			  {
				cols[questionOrderIdx] = attribueValue.value;
			  }
			  else if( attribueValue.attribute.id == uid_DeType )
			  {
				var code = attribueValue.value;
				var value = arrInvDeType[code];
				if( arrInvDeType[code] == undefined )
				{
				  value = code + " : [not exist in Mapping]"; 
				  errorCells.deTypes.push( i + dataRowIdx + 1);
				}
				
				cols[deTypeIdx] = value;
			  }
			  else if( attribueValue.attribute.id == uid_QType )
			  {
				var code = attribueValue.value + " - " + valueType;
				
				var idxQType = arrInvQuestionType.length + 2; 
				var value = arrInvQuestionType[code];
				if( arrInvQuestionType[code] == undefined )
				{
				  value = attribueValue.value + " : [not exist in Mapping ]" + code;   
								
				  errorCells.qType.push( i + dataRowIdx + 1);
				  
				  /*arrInvQuestionType[code] = code;
				  idxQType = arrInvQuestionType.length + 2;
				  mappingSheet.getRange("D" + idxQType ).setValue( code );
				  mappingSheet.getRange("E" + idxQType ).setValue( code ); */  
				}
				 
				cols[questionIdx] = value;   
				
			  }
			  else if( attribueValue.attribute.id == uid_QHide )
			  {
				cols[questionHideIdx] = attribueValue.value;
			  }
			  else if( attribueValue.attribute.id == uid_QHideGroup )
			  {
				cols[questionHideGroupIdx] = attribueValue.value;
			  }
			  else if( attribueValue.attribute.id == uid_QMatch )
			  {
				cols[questionMatchIdx] = attribueValue.value;
			  }
			  else if( attribueValue.attribute.id == uid_QMatchGroup )
			  {
				cols[questionMatchGroupIdx] = attribueValue.value;
			  }
			  else if( attribueValue.attribute.id == uid_NumQ )
			  {
				cols[numWIdx] = attribueValue.value;
			  }
			  else if( attribueValue.attribute.id == uid_DenW )
			  {
				cols[denWIdx] = attribueValue.value;
			  }
			  else if( attribueValue.attribute.id == uid_CompIndicator )
			  {
				cols[compIndicatorIdx] = attribueValue.value;
			  }
			  else if( attribueValue.attribute.id == uid_ImgUrl )
			  {  
				cols[imgUrlIdx] = attribueValue.value;
			  }
			  else if( attribueValue.attribute.id == uid_VideoUrl )
			  {
				cols[videoUrlIdx] = attribueValue.value;
			  }
			  else if( attribueValue.attribute.id == uid_QuestionParents )
			  {
				cols[questionParentsIdx] = attribueValue.value;
			  }
			  else if( attribueValue.attribute.id == uid_QuestionParentOpts )
			  {
				cols[questionParentOptsIdx] = attribueValue.value;
			  }
			}
			
		  } // END attribute values

      // add compulsory check as well
      cols[compulsoryIdx] = data.compulsory;
		  
		  exportData[i] = cols;
		  
		} // END for dataElementList
		
		
		showProgressMessage("Writing data. Please wait ...");
		
		dataSheet.insertRowsAfter(dataRowIdx, dataElementList.length-1);
		dataSheet.getRange(dataRowIdx + 1, 1, dataElementList.length, exportSheetColNo ).setValues( exportData );    
		Utilities.sleep(1000);
		
		showProgressMessage("Setting data validation ...");
		
		setDataValidation( dataSheet, arrayOptionSet, arrQuestionType, arrDeTypes );    
		Utilities.sleep(1000);
		
		if( errorCells.qType.length > 0 
		   || errorCells.optionSet.length > 0 
		   || errorCells.deTypes.length > 0 )
		{
		  showProgressMessage( "Setting color for invalid cells  ..." );
		  
		  for( var k in errorCells.qType )
		  {
			dataSheet.getRange( errorCells.qType[k], questionIdx + 1 ).setBackground("red");
		  }
		  
		  for( var k in errorCells.optionSet )
		  {
			dataSheet.getRange( errorCells.optionSet[k], optionSetIdx + 1 ).setBackground("red");
		  }
		  
		  for( var k in errorCells.deTypes )
		  {
			dataSheet.getRange( errorCells.deTypes[k], deTypeIdx + 1 ).setBackground("red");
		  }
		}
		
		  /* setBgColorValidation( dataSheet ); */
		Utilities.sleep(1000);
		
		ss.setActiveSheet( dataSheet );
		
		Browser.msgBox("Downloaded and wrote data successfully.");
	}
	catch( e )
	{
		Browser.msgBox("Error : " + e.message + " in line "+ e.lineNumber);
	}
   
  } // END if
  else
  {
    Browser.msgBox("No data element in selected program."); 
  }
    
}

function setBgColorValidation( dataSheet )
{
  
    var ss = SpreadsheetApp.getActiveSpreadsheet();
  
    var dataSheet = ss.getSheetByName("DE master");
  
  
  // Check for option sets
    
  for( var i=dataRowIdx + 1; i<dataSheet.getLastColumn(); i++ )
  {
    setBgColorCellValidation( dataSheet.getRange(i, optionSetIdx + 1) );
    setBgColorCellValidation( dataSheet.getRange(i, questionIdx + 1) );
    setBgColorCellValidation( dataSheet.getRange(i, deTypeIdx + 1 ) );
  }
  
}

function setDataValidation( deSheet )
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();  
  var mappingSheet = ss.getSheetByName("Mapping");
  
  // Set data validation for 'Option Set' column
  setDataValidationInColumn( deSheet, "H", mappingSheet, "A" );
  
  
  // Set data validation for 'Question Type' column
  setDataValidationInColumn( deSheet, "L", mappingSheet,"E" );
  
  
  // Set data validation for 'DE Type' column
  setDataValidationInColumn( deSheet, "K", mappingSheet,"V" );
}

function setDataValidationInColumn( deSheet, deSheetColIdx, mappingSheet, mappingSheetColIdx )
{
  var rowNo = deSheet.getLastRow();
  
  // Set the data validation for cells to require a value from 'Mapping' sheet
  
  var dataColNo = dataRowIdx + 1;
  var cell = deSheet.getRange( deSheetColIdx + dataColNo + ':' + deSheetColIdx + rowNo );
  var mappingColNo = 5;
  var range = mappingSheet.getRange( mappingSheetColIdx + mappingColNo + ':' + mappingSheetColIdx + mappingSheet.getLastRow() );
  var rule = SpreadsheetApp.newDataValidation().requireValueInRange( range ).build();
  cell.setDataValidation( rule );
}


// -----------------------------------------------------------------------------------------
// Password dialog box
// -----------------------------------------------------------------------------------------

// Show "Set Password" dialog.

function showPasswordDialog( dataURL, message )
{
	var html = HtmlService.createHtmlOutputFromFile('login.html')
      .setWidth(400)
      .setHeight(120);
	html.append("<input type='hidden' id='dataURL' value='" + dataURL + "'/>");
	
	showProgressMessage( message + " ..." );
	if( message !== undefined )
	{
		html.append("<div style='color:red;font-weight:bold;'><br>" + message + "<div/>");
	}
	
	SpreadsheetApp.getUi().showModalDialog( html, 'Login' );
	  
}

function callLogin( username, password, dataURL )
{ 
	if( username == "" || password == "" )
	{
		showPasswordDialog( dataURL, "Please enter username / password" );
		return;
	}
	
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var paramSheet = ss.getSheetByName("Download Data Form"); 
   
    var serverName = paramSheet.getRange("C2").getValue().toString();
   // app.close();
    
    var retreiveDataType = PropertiesService.getScriptProperties().getProperty('retrieveData');

    if( retreiveDataType == "DataElement" )
    {
      retrieveDataFromServer( username, password, serverName, dataURL ); 
    }
    else
    {
		 if( retreiveDataType == "MetaData" )
		 {
			 var ss = SpreadsheetApp.getActiveSpreadsheet();
			 var paramSheet = ss.getSheetByName( "Download Data Form" );
			 var programChecked = paramSheet.getRange( "A9" ).getValue().toString();
			 var optionSetChecked = paramSheet.getRange( "A10" ).getValue().toString();
				   
			 if( programChecked == "√" ){
				retrieveMetaDataFromServer( username, password, serverName, "Program" ) ; 
			 }
			   
			 if( optionSetChecked == "√" ){
				retrieveMetaDataFromServer( username, password, serverName, "OptionSet" ) ; 
			 }
		 }
		 else
		 {
			 retrieveMetaDataFromServer( username, password, serverName, retreiveDataType ) ; 
		 }
	 
   
    }
  
}

// ------------------------------------------------------------------------------------------------
// Check data in 'DE master' Sheet
// ------------------------------------------------------------------------------------------------


function onEdit() 
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var deSheet = ss.getSheetByName('DE master');
  var activeCell = deSheet.getActiveCell();
  
  // if( activeCell.getColumn() == 16 ) // Option Set
  // else if( activeCell.getColumn() == 11 ) // Q. type
  
  setBgColorCellValidation( activeCell );
}

function setBgColorCellValidation( cell )
{
    var rule = cell.getDataValidation();
    if (rule != null) {
    var cellValue = cell.getValue().toString();
    var criteria = rule.getCriteriaValues();
    
    if (criteria[0].indexOf(cellValue) === -1)
    {
      cell.setBackground("#ff0000");
    }
    else
    {
       cell.setBackground("white");
    }
  }
}


// ------------------------------------------------------------------------------------------------
// Populate data from server
// ------------------------------------------------------------------------------------------------


function setUp_ProgramCheckBox()
{
  setUp_CheckBox( "Program" );
}

function setUp_OptionSetCheckBox()
{
  setUp_CheckBox( "OptionSet" );
}

function setUp_CheckBox( dataType )
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var paramSheet = ss.getSheetByName( "Download Data Form" );
  
  var cell;
  if( dataType== "Program" ){
	cell = paramSheet.getRange( "A9" );
  }
  else if( dataType== "OptionSet" ){
    cell = paramSheet.getRange( "A10" );
  }   
  
  var checked = cell.getValue().toString();
  if( checked == "√" ){
	cell.setValue("");
  }
  else{
	cell.setValue("√");  
  }
}
 
 
// Validate the parametters. If they are valid, connect the server and retreive data
 
function call_retrieveDataFromServer( dataType )
{
   var ss = SpreadsheetApp.getActiveSpreadsheet();
   var paramSheet = ss.getSheetByName( "Download Data Form" );
   var serverName = paramSheet.getRange( "C2" ).getValue().toString();
   var program = paramSheet.getRange( "C5" ).getValue().toString();
  
    if( serverName == "" || ( dataType == 'DataElement' && program == "" )  )
    {
	  Browser.msgBox( "Please enter a server name and choose a program in \"Download Data Form\" sheet" );
    }
    // Retrieve Meta data
	else
    {
	  loadDataFromServer( dataType );
    }
}

function loadDataFromServer( dataType )
{
	PropertiesService.getScriptProperties().setProperty( "retrieveData", dataType );
	// Retrieve Meta data ( programs, option sets )
	if( dataType != "DataElement" )
    {
      showPasswordDialog( "" );
    }
	// Retrieve Data Elements
	else
    {
	  var ss = SpreadsheetApp.getActiveSpreadsheet();
      var paramSheet = ss.getSheetByName( "Download Data Form" );
      var program = paramSheet.getRange( "C5" ).getValue().toString();
	  
      loadMappingData();
	  var arrPrograms = _programs;
      var programUid = arrPrograms[program];
      var dataURL = "/api/programs/" + programUid + ".json?fields=programStages[id,name,programStageDataElements[compulsory,dataElement[id,name,shortName,code,formName,description,aggregationType,valueType,optionSet[name,id],attributeValues[attribute[id],value]]]]";
      
      showPasswordDialog( dataURL );
    }
}

function call_retrieveDataElementFromServer()
{
   call_retrieveDataFromServer( "DataElement" );
}

function call_retrieveMetaDataFromServer()
{
	call_retrieveDataFromServer( "MetaData" );
}

function call_retrieveProgramFromServer()
{
	call_retrieveDataFromServer( "Program" );
}

function call_retrieveOptionSetFromServer()
{
	call_retrieveDataFromServer( "OptionSet" );
}

// Connect server and retreive data
// Fill data into "DE master" sheet
function retrieveMetaDataFromServer( uname, upass, serverName, dataType )
{
  showProgressMessage("Accessing server to retrieve " + dataType + " ...");
  
  var user = uname + ":" + upass;
  
  // PULL URL from Config spreadsheet to allow the user to change which server will be used
  var url = serverName;
  if( dataType == "Program" ){
	url += "/api/programs.json?paging=false&fields=id,name";
  }
  else if( dataType == "OptionSet" ){
	url += "/api/optionSets.json?paging=false&fields=id,name";
  }
  
  var options =
      {
        "contentType" : "application/json",
        "method" : "get",
        "headers": {
          "Authorization": "Basic " + Utilities.base64Encode(user)
        }
      };
  
  var response;
  var valid = true;
  try
  {
    response = UrlFetchApp.fetch(url, options);
  } 
  catch (e) {
    valid = false;
    
	if( e.message.indexOf("\"httpStatus\":\"Unauthorized\"") >= 0 )
    {
		Browser.msgBox("Username/Password is wrong. " );
	}
	else
	{
		Browser.msgBox("Invalid URL. " + e.message);
	}
	
    showPasswordDialog( dataURL  );
  }
  
  if(valid)
  {
    writeMetaData( dataType, response );
  }

}

/**
	startRowIdx : start from 1
	nameColIdx: start from 1
**/
function writeMetaData( dataType, response )
{
   // -------------------------------------------------------------------------------
   // Get meta data list
   var jsonData = JSON.parse( response );
   var dataList = [];
   var nameColIdx = 0;
   if( dataType == "Program" ){
	 dataList = jsonData.programs;
     nameColIdx = 15;
   }
   else if( dataType == "OptionSet" ){
	 dataList = jsonData.optionSets;
     nameColIdx = 1;
   }
  
   // -------------------------------------------------------------------------------
   // Get meta data list
   
   var ss = SpreadsheetApp.getActiveSpreadsheet();
   var mappingSheet = ss.getSheetByName("Mapping");
   var formSheet = ss.getSheetByName("Download Data Form");
  
   showProgressMessage("Removing " + dataType + " data in Mapping sheet ...");
   mappingSheet.getRange(mappingDataIdx, nameColIdx, mappingSheet.getLastRow(), 2).setValue( "" );
   
   if( dataList.length > 0 )
   { 
     showProgressMessage("Writing " + dataType + " data in Mapping sheet ...");
     
     var nameList = new Array();
     var exportData = new Array();     
     for( var i in dataList)
     {
       var cols = new Array( 2 );
       cols[0] = dataList[i].name;
       cols[1] = dataList[i].id;
       exportData[i] = cols;
       
       nameList[i] = dataList[i].name;
     }
     
     mappingSheet.getRange(mappingDataIdx, nameColIdx, dataList.length, 2).setValues( exportData ).setBackground("#D0E0E3");     
     
     if( dataType == "Program" ){
		showProgressMessage("Set data validation for Program field ...");
		formSheet.getRange("C5").setDataValidation( SpreadsheetApp.newDataValidation().requireValueInList( nameList ).build() );          
     }
     
     Browser.msgBox("Wrote " + dataType + " data in Mapping file.") 
   }
   else
   {
    Browser.msgBox("No " + dataType + " on server.") 
   }
    
}

function writeProgramData( dataType, response )
{
   var jsonData = JSON.parse( response );
   var programList = jsonData.programs;
   
   var ss = SpreadsheetApp.getActiveSpreadsheet();
   var mappingSheet = ss.getSheetByName("Mapping");
   var formSheet = ss.getSheetByName("Download Data Form");
  
   showProgressMessage("Removing Program list in Mapping sheet ...");
   mappingSheet.getRange("O5:P").setValue( "" );
     
   
   if( programList.length > 0 )
   { 
     showProgressMessage("Writing programs in Mapping sheet ...");
     
     var programNameList = new Array();
     var exportData = new Array();     
     for( var i in programList)
     {
       var cols = new Array( 2 );
       cols[0] = programList[i].name;
       cols[1] = programList[i].id;
       exportData[i] = cols;
       
       programNameList[i] = programList[i].name;
     }
     
     mappingSheet.getRange(5, 15, programList.length, 2).setValues( exportData ).setBackground("#D0E0E3");     
     
     
     showProgressMessage("Set data validation for Program field ...");
     
     formSheet.getRange("C5").setDataValidation( SpreadsheetApp.newDataValidation().requireValueInList( programNameList ).build() );          
     
     
     Browser.msgBox("Wrote data in Mapping file.") 
   }
   else
   {
    Browser.msgBox("No program on server.") 
   }
    
}

function test()
{
	retrieveDataFromServer( "jamesc", "Passmed0", "https://leap.psi-mis.org", "/api/programs/kA0IzIoVyj0.json?fields=programStages[id,name,programStageDataElements[compulsory,dataElement[id,name,shortName,code,formName,description,aggregationType,valueType,optionSet[name,id],attributeValues[attribute[id],value]]]]" );
}

// Test API call
// function testAPI()
// {
// 	let testObj = retrieveDataFromServer( "UGdemo1", "Ugandademo1!", "https://data.psi-mis.org", "/api/programs/xUZKZbqE6aR.json?fields=programStages[id,name,programStageDataElements[compulsory,dataElement[id,name,shortName,code,formName,description,aggregationType,valueType,optionSet[name,id],attributeValues[attribute[id],value]]]]" );
//   Logger.log(testObj);
// }
