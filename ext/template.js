// global var
let cScoreNamePosition;
let checklistName;
let activeSheet = SpreadsheetApp.getActiveSheet().getName();

// ---------------------------------------------------------------------------------
// Download Template
//
// ---------------------------------------------------------------------------------

/*
display a sidebar menu with options to download the template
*/
function openSidebar() {
  let html = HtmlService.createHtmlOutputFromFile("download").setTitle("Download Template");

  if (activeSheet === "HNQIS 1.6"){
    SpreadsheetApp.getUi().showSidebar(html);
  } else{
    html = HtmlService.createHtmlOutput("<script>google.script.host.close();</script>");
    SpreadsheetApp.getUi().showSidebar(html);
  }
}

/* 
 create a data url to download the checklist as excel or pdf
*/
function createDataUrl(type) {
  let mimeTypes = { xlsx: MimeType.MICROSOFT_EXCEL, pdf: MimeType.PDF };
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let template = ss.getSheetByName("HNQIS 1.6");

  let configForm = ss.getSheetByName("Download Data Form");
  let templateName = configForm.getRange("C5").getValue();
  templateName = templateName + " 2021.09"
  const url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=${type}&gid=${template.getSheetId()}`;

  const blob = UrlFetchApp.fetch(url, {
    headers: { authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
  }).getBlob();
  return {
    data:
      `data:${mimeTypes[type]};base64,` +
      Utilities.base64Encode(blob.getBytes()),
    filename: `${templateName}.${type}`,
  };

}



function call_generateTemplateToSheet(){

  generateTemplate();
  openSidebar();

}


// generateTemplate

function generateTemplate(){
  
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let deMaster = ss.getSheetByName("DE master");
  let data = deMaster.getDataRange().getValues();

  let configForm = ss.getSheetByName("Download Data Form");
  let serverName = configForm.getRange("C2").getValue();

  checklistName = configForm.getRange("C5").getValue();

  let template = data.map(function(v){
    return [
      v[5],v[7],v[22],v[14],v[14],v[14],v[14],v[11],v[17],v[19],  // replace v[23] with v[14] * 4
      v[4] + " " + v[20] + " " + v[21],
      v[12], v[0], v[1], v[9], v[10], v[8],v[27]
    ]
  });

  template = deleteRow(template,2);
  template = deleteRow(template,0);

  // rename header
  template = renameHeader(template);

  // rename parent child relations
  template = renameParentChildRelations(template);

  // transform answer options
  template = transformAnswerOptions(template);

  // sort by Qorders
  template = template.sort(function(v1,v2){
    return v1[11] - v2[11];
  });

  // filterout the blanks
  template = template.filter(function(v){
    if (v[11] === ""){
      return false;
    } else{
      return true;
    }
  });

  // transform qTypes
  template = transformQuestionTypes(template);

  // Fix Mandatory Des
  template = highlightMandatory(template);
  //Logger.log(template);

  // review Template & orderByHeaders
  template = reviewTemplate(template)

  // Review qOrders
  //template = orderByHeader(template);
  //orderByHeader(template);

  // remove the last X (n + 1) column i.e., id & column
  template = template.map(function(v){
    return v.slice(0,-7);
  });

  // Get a copy of summary
  let summary = template.filter(function(v){
    if (v[7] === "93 Compositive_Score"){
      return true;
    } else{
      return false;
    }
  });

  // re-transform the qTypes clearing the composite scores
  template = reTransformQuestionTypes(template);

  // remove the overall score
  compositeScores = template.map(function(v){return v[9];});
  overallScoreIndex = compositeScores.indexOf(0.0);

  //Logger.log(overallScoreIndex);
  template = deleteRow(template, overallScoreIndex);

  // unshift the first composite score
  let cs = template.map(function(v){return v[9];});
  let firstCS = cs.indexOf(1);
  //template = template.unshift(template[firstCS]);
  template.unshift(template[firstCS]);
  template.splice(firstCS+1,1);

  // swtich back the header
  let tabHeader = cs.indexOf("Composite Indicator");
  template.unshift(template[tabHeader + 1]);
  template.splice(tabHeader + 2,1);

  // write
  let target = ss.getSheetByName("HNQIS 1.6");
  let targetRange = target.getRange(4,1,template.length, template[0].length);
  targetRange.setValues(template);

  // set the list name & server
  target.getRange(1,1).setValue("Checklist: " + checklistName);
  //target.getRange(1,2).setValue(checklistName);

  target.getRange(2,1).setValue("Server: " + serverName);
  //target.getRange(2,2).setValue(serverName);

  // write summary names
  let targetSummaryNameRange = target.getRange(template.length + 5, 1, summary.length,1);
  targetSummaryNameRange.setValues(summary.map(function(v){ return [v[0]];}));

  // write summary composite scores
  let targetSummaryScoreRange = target.getRange(template.length + 5, 10, summary.length,1);
  targetSummaryScoreRange.setValues(summary.map(function(v){ return [v[9]];}));

  // set header
  target.getRange(template.length + 4,1, 1,1).setValue("Composite Scores");

  formatTemplate("HNQIS 1.6");

  




}



// review the template; 
// i.e., remap the CS scores for questions without CS core - use the first CS score for questions mathcing the header
function reviewTemplate(data){
  // filter blank CS
  let withBlankCS = data.filter(function(v){
    if (v[15] === "QUESTION" && v[9] === ""){
      return true;
    } else{
      return false;
    }
  });

  // get the header values
  let headerValuesToUpdate = withBlankCS.map(function(v){

    // also get the uids of questions with blank CS
    // v[14] == header index and v[16] == tab index
    return [v[14],v[16],v[12]];
    });
  

  // filter data with eq headers and get the CS score assigned to any of the results
  let headerValues = headerValuesToUpdate;
  for (var i = 0; i < headerValuesToUpdate.length; i++){
    // filter the header
    let dataByHeader = data.filter(function(v){
      if (headerValuesToUpdate[i][0] === headerValuesToUpdate[i][1]){
        // match by tab
        if (v[16] === headerValuesToUpdate[i][1] && v[9] !== ""){
          return true;
        } else if (v[0] === headerValuesToUpdate[i][1] && v[9] !== "" && v[15] === "COMPOSITE_SCORE"){ //also match by name for CS questions
          return true;
        }
         else {
          return false;
        }
      } else {
        // match by header
        if (v[14] === headerValuesToUpdate[i][0] && v[9] !== ""){
          return true;
        } else if (v[0] === headerValuesToUpdate[i][0] && v[9] !== "" && v[15] === "COMPOSITE_SCORE"){ //also match by name for CS questions
          return true;
        }
        else {
          return false;
        }
      }
    });

    // get CS scores
    let dataByHeaderCScores = dataByHeader.map(function(v){return v[9]});
    dataByHeaderCScores = dataByHeaderCScores.filter(onlyUnique);
    dataByHeaderCScores = dataByHeaderCScores.sort(function(v1,v2){ return v1 - v2;});

    headerValuesToUpdate[i][3] = dataByHeaderCScores[0];
    
  }

  // remap CS scores
   for (var i = 0; i < headerValuesToUpdate.length; i++){
     // index of the remapped row
     let ids = data.map(function(v){return v[12]});
     let rowInd = ids.indexOf(headerValuesToUpdate[i][2]);

     // remap the template CS scores
     if (rowInd >= "0"){
       data[rowInd][9] = headerValuesToUpdate[i][3];
     }

   }

   // orderByHeader
   data = orderByHeader(data);

   // revert back the CS scores
   for (var i = 0; i < headerValuesToUpdate.length; i++){
     // index of the remapped row
     let ids = data.map(function(v){return v[12]});
     let rowInd = ids.indexOf(headerValuesToUpdate[i][2]);
     

     // remap the template CS scores
     if (rowInd > "0"){
      data[rowInd][9] = ""; // headerValuesToUpdate[i][3];
     }

   }
  
  
  return data;

}

// sort DEs
function orderByHeader(data){
  let firstRow = data[0];
  let headers = data.map(function(v){
    return v[14];
  });

  let qOrders = data.map(function(v){return v[11];});
  let uid = data.map(function(v){return v[12];});
  let deTypes = data.map(function(v){return v[15];});
  let cScore = data.map(function(v){return v[9]});
  // custom sort the CS scores 
  let uniqueCScores = cScore.sort(function(a,b){
    var a1 = a;
    a1 = a1.toString();
    var b1 = b;
    b1 = b1.toString();

    // normalize cs scores
    if (a1.charAt(3) === "."){
      a1 = a1.slice(0,3) + a1.slice(4);
    }

    if (b1.charAt(3) === "."){
      b1 = b1.slice(0,3) + b1.slice(4);
    }

    if (a1.charAt(0) === b1.charAt(0) && a1.length >= 4 && b1.length === 3){
      return 1;
    }
    if (a1.charAt(0) === b1.charAt(0) && b1.length >= 4 && a1.length === 3){
      return -1;
    }

    //return a - b;
    return a1 - b1;
  }).filter(onlyUnique)
    .slice(1)
  //  .slice(0,-1);
  //Logger.log(uniqueCScores);

  // get a filter of the data by headers
  let dataPieces = [];
  for (var i = 0; i < uniqueCScores.length; i++){
    // Get the unique tab name to search
    let score = uniqueCScores[i];
    var dataByTab = data.filter(function(v){
      if (v[9] === score){
        return true;
      } else {
        return false;
      }
    });


    // sort by qOrders
    dataByTab = dataByTab.sort(function(v1,v2){
      return v1[11] - v2[11];
    })

    //&& v[15] !== "COMPOSITE_SCORE"

    var tabName = dataByTab.map(function(v){
      return v[14];
    });

    //Logger.log(tabName);

    tabName = tabName.filter(onlyUnique);
    

    //remove the Composite Scores
    let index = tabName.indexOf("Composite Scores");
    if (index != -1){
      tabName.splice(index,1);
    }

    

    dataPieces[i] = data.filter(function(v){
    if (v[9] === score && v[15] === "COMPOSITE_SCORE"){
      return true;
    } else if (v[9] === score){
      return true;
    //} else if (tabName.length > 0 && v[14] === tabName[0]) {
     /// return true;
    } else {
      return false;
    }
    
    
    });

    // first just sort everythign by qHeaders
    dataPieces[i] = dataPieces[i].sort(function(v1, v2){
      return v1[11] - v2[11];
    });
    // ensure CS questions are always at the front
    dataPieces[i] = dataPieces[i].sort(function(v1,v2){
      return v1[15] == "COMPOSITE_SCORE" ? -1 : v2[15] == "COMPOSITE_SCORE" ? 1 : 0;
    });
  }

   //Logger.log(dataPieces[29].map(function(v){return [v[7],v[9],v[11]];}));
   //Logger.log(dataPieces[2]);
  
  // Remove the duplicated rows

  let res = dataPieces.flat();
  // let uids = res.map(function(v){return v[12];});
  // let uniqueUids = uids.filter(onlyUnique);
  // let uniqueUidIndex = [];
  // for (var i = 0; i < uniqueUids.length; i++){
  //   uniqueUidIndex[i] = uids.indexOf(uniqueUids[i]);
  // }
  

  // //let indexOfUniqueIDs = res.map(function(v){ return v[12].indexOf()}); 
  // //Logger.log(uniqueUidIndex);

  // let result = [];
  // for (var i = 0; i < uniqueUidIndex.length; i++){
  //   result[i] = res[uniqueUidIndex[i]];
  // }

  //Logger.log(uniqueCScores);
  //Logger.log(res.length);

 
  


  // res = res.reduce((a,c) => {
  //   if (!a.find(v => v[12] === c[12])){
  //     a.push(c);
  //   }
  //   return a;
  // }, []);
  //Logger.log(res.length);
  //res = res.unshift(firstRow);
  res.unshift(firstRow);

  return res;


}



// Utils: getMandatoryDES
function getMandatoryQuestionIndex(data, qUidIndex){
 
  let allIds = data.map(function(v){return v[qUidIndex];});
  let mandatoryDes = data.filter(function(v1){
    if (v1[17] === true){
      return true;
    } else{
      return false;
    }
  });
  let mandatoryDesIds = mandatoryDes.map(function(v){return v[qUidIndex];});

  let indecies = [];

  if (mandatoryDesIds.length >= 1){
    for (var i = 0; i < mandatoryDesIds.length; i++){
      indecies[i] = allIds.indexOf(mandatoryDesIds[i]);
    }
  }


  return indecies;
}

// utils
function getAllIndexes(arr, val) {
    var indexes = [], i = -1;
    while ((i = arr.indexOf(val, i+1)) != -1){
        indexes.push(i);
    }
    return indexes;
  }
// utils2: getMandatoryQuestionIndexFromTemplate
function getMandatoryQuestionIndexFromTemplate(data, qNameIndex){
  let allDes = data.map(function(v){return v[qNameIndex];});
  let mandatoryDes = allDes.filter(function(v){
    return v.includes("*");
  });

  let indexes = [];

  if (mandatoryDes.length >= 1){
    for (var i = 0; i < mandatoryDes.length; i++){
      indexes[i] = getAllIndexes(allDes,mandatoryDes[i]);
    }
  }

  indexes = indexes.flat();
  indexes = indexes.filter(onlyUnique);

  return indexes;

}

// Highlight Mandatory DEs
function highlightMandatory(data){
  
  // get mandattory questions positions
  let indecies = getMandatoryQuestionIndex(data, 12);

  let qNames = data.map(function(v){
    return v[0];
  });

  let qNameTranformed = [];

  for (var i = 0; i < qNames.length; i++){
    // filter the matching mandatory q
    var ind = indecies.filter(function(v){
        if (v === i){
          return true;
        } else{
          return false;
        }  
      });

    if (ind.length > 0){
      qNameTranformed[i] = "*" + qNames[i];
    } else {
      qNameTranformed[i] = qNames[i];
    }
  }

  // remap the qNames in the template
  for (var i = 0; i < data.length; i++){
    data[i][0] = qNameTranformed[i];
  }
  
  return data;
}



function formatTemplate(sheetName){
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let template = ss.getSheetByName(sheetName);

  let header1 = template.getRange("A4:k4");
  let table = template.getDataRange();

  let firstCS = template.getRange("A5:K5");

  // edit header
  //header1.setFontWeight("bold");
  header1.setFontColor("white");
  header1.setBackground("#807F80");
  //header1.setFontFamily("Roboto");
  header1.setWrap(true);
  // edit row size
  template.setRowHeight(4, 100);
  template.setRowHeight(3, 10);
  template.setColumnWidth(11, 600);
  template.setColumnWidth(3,130);
  template.setFrozenRows(4);

  // edit table 
  table.setFontFamily("Calibri");
  //table.setFontFamily("Arial");
  table.setHorizontalAlignment("left");
  table.setBorder(true, true, true,true,false,true,"#D4D4D3",SpreadsheetApp.BorderStyle.SOLID);
  table.setWrap(true);

  // get the composite scores indecies
  // i.e indecies where the data range is blank
  let data = table.getValues();

  //let qTypes = data.map(function(v){return v[7]});
  //let qNames = data.map(function(v){return v[0]});

  qTypeIndex = getAllCompositeScoresPositions("", data, 7, 0);

  // color all the composite score to orange
  for (var i = 0; i < qTypeIndex.length; i++){
    cScore = qTypeIndex[i] + 1;
    template.getRange("A" + cScore + ":K" + cScore).setBackground("#F8E4D8");
  }
  
  // color the composite score lable
  // from the global variable
  let cScoreRow = template.getRange("A"+cScoreNamePosition + ":K" + cScoreNamePosition);
  cScoreRow.setFontColor("white");
  //cScoreRow.setFontFamily("Roboto");
  cScoreRow.setBackground("#2B4D74");
  //cScoreRow.setFontWeight("bold");
  //Logger.log(cScoreNamePosition);
  firstCS.setBackground("#DE8344");

  // format Mandatory questions
  // get Mandatory DES postion
  qNameIndex = getMandatoryQuestionIndexFromTemplate(data,0);
  //var bold = SpreadsheetApp.newTextStyle().setBold(true).build();
  //var redText = SpreadsheetApp.newRichTextValue().setText("*").setTextStyle(0,1,bold).build(); 

  // color the asterick
  for (var i = 0; i < qNameIndex.length; i++){
    mQIndex = qNameIndex[i] + 1;
    //template.getRange("A" + cScore + ":K" + cScore).setBackground("#F8E4D8");
    //value = template.getRange("A"+mQIndex).getValue();
    //Logger.log(value);
    //var redText = SpreadsheetApp.newRichTextValue().setText(value).setTextStyle(0,2,bold).build();
    //template.getRange("A"+mQIndex).setRichTextValue(redText);
    template.getRange("A"+mQIndex).setFontColor("#DE3021");
  }

  // Grayout cells without CS scores
  let rowIndWithoutCScore = getAllPositions("",data,9,7);

  for (var i = 0; i < rowIndWithoutCScore.length; i++){
    var rowIndCscore = rowIndWithoutCScore[i] + 1;
    template.getRange("J" + rowIndCscore).setBackground("#D4D4D3");
  }

  // Gray out cells without answer options
  let rowIndWithoutAnsOpts = getAllPositions("", data, 1, 7);

  for (var i = 0; i < rowIndWithoutAnsOpts.length; i++){
    var rowIndAnsOpts = rowIndWithoutAnsOpts[i] + 1;
    template.getRange("B" + rowIndAnsOpts).setBackground("#D4D4D3");
  }

  // Highlight the tabs; Interger CS scores

  let intCS = getAllIntegerCSPositions(data,9);

  // filter the CS scores from the summary
  let intCSClean = [];
  for (var i = 0; i < intCS.length; i++){
    intCSClean[i] = qTypeIndex.filter(function(v){
      if (v === intCS[i]){
        return true;
      } else{
        return false;
      }
    });
  }

  intCSClean = intCSClean.flat();

  // highlight the tabs
  for (var i = 0; i < intCSClean.length; i++){
    intcScore = intCSClean[i] + 1;
    template.getRange("A" + intcScore + ":K" + intcScore).setBackground("#D18751");
  }






}

function getAllPositions(elementToFind,arr,col,qNameIndex){
  var indecies = [];
  for(var i = 0; i < arr.length; i++){
    if (arr[i][col] === elementToFind && arr[i][qNameIndex] !== ""){
      indecies.push(i)
    }
  }
  return indecies;

}

function getAllIntegerCSPositions(arr,col){
  var indecies = [];
  for(var i = 0; i < arr.length; i++){
    var item = arr[i][col];
    if (item.toString().length === 1 && item !== "" ){
      indecies.push(i)
    }
  }
  return indecies;

}


function getAllCompositeScoresPositions(elementToFind, data, qTypeIndex, qNameIndex) {
  var indecies = [];
    for (var i = 0; i < data.length; i++) {
        if (data[i][qTypeIndex] === elementToFind && data[i][qNameIndex] !== "") {
            indecies.push(i);
        } 
    }

  //let qCSIndex = data.map(function(v){return v[0].indexOf("Composite Scores");});
  let qCSIndex = data.map(function(v){return v[0];});

  qCSIndex = qCSIndex.indexOf("Composite Scores");

  //write index to global var cScoreNamePosition
  cScoreNamePosition =  qCSIndex + 1;

  // indeceis not to return
  let qCSIndexToDel = indecies.filter(function(v){
    if (v >= qCSIndex){
      return true;
    } else{
      return false;
    }
  });

  indecies = indecies.splice(qCSIndex[0], qCSIndexToDel.length); 

  return indecies.slice(2);

}




function reTransformQuestionTypes(data){
  let questionTypes = data.map(function(v){
    return v[7];
  });

  questionTypes = questionTypes.map(function(v){
    if (v === "93 Compositive_Score"){
      return "";
    } else{
      return v;
   }
  });

  // remove the first element and remap the values
  questionTypes = questionTypes.slice(1);

  for (var i = 1; i < data.length; i++){
    data[i][7] = questionTypes[i-1];
  }

  return data;

}



function transformQuestionTypes(data){
  let questionTypes = data.map(function(v){
    return v[7];
  });

  // transform
  questionTypes = questionTypes.map(function(v){
    if (v === "02 INT"){
      return "93 Compositive_Score";
    } else if (v === "RADIO_GROUP_HORIZONTAL"){
      return "08 RADIO_HORIZONTAL";
    } else if (v === "RADIO_GROUP_VERTICAL"){
      return "09 RADIO_VERTICAL";
    } else {
      return v;
    }
  });

  // remove the first element and remap the values
  questionTypes = questionTypes.slice(1);

  for (var i = 1; i < data.length; i++){
    data[i][7] = questionTypes[i-1];
  }

  return data;
}


function transformAnswerOptions(data){
  let answers = data.map(function(v){
    return v[1].split("-")[1];
  });

  // remove the first element and remap the values
  answers = answers.slice(1);

  for (var i = 1; i < data.length; i++){
    data[i][1] = answers[i-1];
  }


  return data;
}



// renameParentChildRelations
function renameParentChildRelations(dt){
   // Get the des IDs
  let ids = dt.map(function(v){
    return v[12];
  });

  // Get the parent Ids to search
  let parentIds = dt.map(function(v){
    return v[2];
  }); 

  parentIds = parentIds.slice(1);
  parentIds = parentIds.filter(onlyUnique);
  parentIds = parentIds.filter(hasValue);

  // map Parent Ids
  let parentResult = new Array(dt.length);

  for (var i = 0; i < dt.length; i++){
    for (var j = 0; j < parentIds.length; j++){
      if (dt[i][2] === parentIds[j]){
        index = ids.indexOf(parentIds[j]);
        parentResult[i] = dt[index][0];
      } 
    }
  }

  // slice the first element & replace the Parent ids after the first element
  parentResult = parentResult.slice(1);
  

  for (var i = 1; i < dt.length; i++){
    dt[i][2] = parentResult[i-1];
  }

  return dt;

}


function onlyUnique(value, index, self) {
 return self.indexOf(value) === index;
}

function hasValue(item){
  if (item === ""){
      return false;
    } else{
      return true;
    }
}
  
// rename the header
function renameHeader(data){
  let header = [
    "Tab / Header / Question",
    "Answer Option",
    "Hidden logic (Parent/Child) Level I",
    "Hidden logic (Parent/Child) Level II",
    "Hidden logic (Parent/Child)",
    "Match logic (Parent/Child)",
    "Match logic (code)",
    "Question Type",
    "Score",
    "Composite Indicator",
    "Feedback Content",
    //"Image URL",
    //"Video URL",
    "Qorder",
    "UID"
  ];

  // loop through the first row and rename the objects
  for (var row = 0; row < 1; row ++){
    for (var col = 0; col < header.length; col++){
      data[row][col] = header[col];
    }
  }

  return data;


}




function extractQuestions(arr, indexArr){  
  return arr.map(obj => {
                 return obj.filter((ob, index) => {
    if (indexArr.includes(index)){return ob}
  })
})
}


function deleteRow(arr, rowIndex) {
   arr = arr.slice(0); // make copy
   arr.splice(rowIndex, 1);
   return arr;
}
