function ReimbursementForm() {
  Browser.msgBox("If a link did not open check that the pop-up is not blocked.\\nOr... alternatively click here:\\n"+"test");
  var newForm = FormApp.create('reimbursement');
  var ss = SpreadsheetApp.create('Form'); 

}

/**
 * Open a URL in a new tab.
 */
function testURL(){
  openURL('https://docs.google.com/spreadsheets/d/17avDF_B7j2tUHU7nu9ywKF3xpXFO5Jcx__rpBUQhlu8/edit?usp=sharing')
}

function openURL(url){
  var html = HtmlService.createHtmlOutput('<html><script>'
  +'window.close = function(){window.setTimeout(function(){google.script.host.close()},9)};'
  +'var a = document.createElement("a"); a.href="'+url+'"; a.target="_blank";'
  +'if(document.createEvent){'
  +'  var event=document.createEvent("MouseEvents");'
  +'  if(navigator.userAgent.toLowerCase().indexOf("firefox")>-1){window.document.body.append(a)}'                          
  +'  event.initEvent("click",true,true); a.dispatchEvent(event);'
  +'}else{ a.click() }'
  +'close();'
  +'</script>'
  // Offer URL as clickable link in case above code fails.
  +'<body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="'+url+'" target="_blank" onclick="window.close()">Click here to proceed</a>.</body>'
  +'<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script>'
  +'</html>')
  .setWidth( 90 ).setHeight( 1 );
  SpreadsheetApp.getUi().showModalDialog( html, "Opening ..." );
  var htmlManual = HtmlService
  Browser.msgBox('If the link did not open, check pop-ups are not disabled.\\nOr... Alternatively copy this link:\\n'+url)
  .setWidth(420) //optional
  .setHeight(80); //optional
}

function AddEntryForm() {
  var ActiveSheet = SpreadsheetApp.getActiveSpreadsheet();
  var FormSheet = ActiveSheet.getSheetByName("Receipt Input Form");
  var DataSheet = ActiveSheet.getSheetByName("Receipt Input Form Back-End");

  var values1 = [[FormSheet.getRange("E6").getValue(),//general info
                 FormSheet.getRange("E8").getValue(),
                 FormSheet.getRange("E10").getValue(),
                 FormSheet.getRange("E14").getValue(),//vendor info
                 FormSheet.getRange("E16").getValue(),
                 FormSheet.getRange("E18").getValue(),
                 FormSheet.getRange("E20").getValue(),
                 FormSheet.getRange("H6").getValue(),//finance info
                 FormSheet.getRange("H8").getValue(),
                 FormSheet.getRange("H10").getValue(),
                 FormSheet.getRange("H12").getValue(),
                 FormSheet.getRange("H14").getValue(),
                 FormSheet.getRange("H18").getValue(),//link/comment info
                 FormSheet.getRange("H20").getValue(),
                 FormSheet.getRange("H22").getValue(),
                 FormSheet.getRange("E30").getValue(),//project info
                 FormSheet.getRange("E32").getValue(),
                 FormSheet.getRange("E34").getValue(),
                 FormSheet.getRange("E36").getValue(),
                 FormSheet.getRange("D39").getValue(),//proj1 info
                 FormSheet.getRange("E39").getValue(),//proj1 info
                 FormSheet.getRange("F39").getValue(),//proj1 info
                 FormSheet.getRange("E44").getValue(),
                 FormSheet.getRange("E46").getValue(),
                 FormSheet.getRange("E48").getValue(),
                 FormSheet.getRange("E50").getValue(),
                 FormSheet.getRange("E52").getValue(),
                 FormSheet.getRange("E54").getValue(),
                 FormSheet.getRange("H30").getValue(),//finance info
                 FormSheet.getRange("H32").getValue(),
                 FormSheet.getRange("H34").getValue()]];
  
  var values2 = [[FormSheet.getRange("E6").getValue(),//general info
                 FormSheet.getRange("E8").getValue(),
                 FormSheet.getRange("E10").getValue(),
                 FormSheet.getRange("E14").getValue(),//vendor info
                 FormSheet.getRange("E16").getValue(),
                 FormSheet.getRange("E18").getValue(),
                 FormSheet.getRange("E20").getValue(),
                 FormSheet.getRange("H6").getValue(),//finance info
                 FormSheet.getRange("H8").getValue(),
                 FormSheet.getRange("H10").getValue(),
                 FormSheet.getRange("H12").getValue(),
                 FormSheet.getRange("H14").getValue(),
                 FormSheet.getRange("H18").getValue(),//link/comment info
                 FormSheet.getRange("H20").getValue(),
                 FormSheet.getRange("H22").getValue(),
                 FormSheet.getRange("E30").getValue(),//project info
                 FormSheet.getRange("E32").getValue(),
                 FormSheet.getRange("E34").getValue(),
                 FormSheet.getRange("E36").getValue(),
                 FormSheet.getRange("D40").getValue(),//proj2 info
                 FormSheet.getRange("E40").getValue(),//proj2 info
                 FormSheet.getRange("F40").getValue(),//proj2 info
                 FormSheet.getRange("E44").getValue(),
                 FormSheet.getRange("E46").getValue(),
                 FormSheet.getRange("E48").getValue(),
                 FormSheet.getRange("E50").getValue(),
                 FormSheet.getRange("E52").getValue(),
                 FormSheet.getRange("E54").getValue(),
                 FormSheet.getRange("H30").getValue(),//finance info
                 FormSheet.getRange("H32").getValue(),
                 FormSheet.getRange("H34").getValue()]];

  var values3 = [[FormSheet.getRange("E6").getValue(),//general info
                 FormSheet.getRange("E8").getValue(),
                 FormSheet.getRange("E10").getValue(),
                 FormSheet.getRange("E14").getValue(),//vendor info
                 FormSheet.getRange("E16").getValue(),
                 FormSheet.getRange("E18").getValue(),
                 FormSheet.getRange("E20").getValue(),
                 FormSheet.getRange("H6").getValue(),//finance info
                 FormSheet.getRange("H8").getValue(),
                 FormSheet.getRange("H10").getValue(),
                 FormSheet.getRange("H12").getValue(),
                 FormSheet.getRange("H14").getValue(),
                 FormSheet.getRange("H18").getValue(),//link/comment info
                 FormSheet.getRange("H20").getValue(),
                 FormSheet.getRange("H22").getValue(),
                 FormSheet.getRange("E30").getValue(),//project info
                 FormSheet.getRange("E32").getValue(),
                 FormSheet.getRange("E34").getValue(),
                 FormSheet.getRange("E36").getValue(),
                 FormSheet.getRange("D41").getValue(),//proj3 info
                 FormSheet.getRange("E41").getValue(),//proj3 info
                 FormSheet.getRange("F41").getValue(),//proj3 info
                 FormSheet.getRange("E44").getValue(),
                 FormSheet.getRange("E46").getValue(),
                 FormSheet.getRange("E48").getValue(),
                 FormSheet.getRange("E50").getValue(),
                 FormSheet.getRange("E52").getValue(),
                 FormSheet.getRange("E54").getValue(),
                 FormSheet.getRange("H30").getValue(),//finance info
                 FormSheet.getRange("H32").getValue(),
                 FormSheet.getRange("H34").getValue()]];


  var ActiveSheet = SpreadsheetApp.getActiveSpreadsheet() * 1;
  var FormCells = [
    "E46", "E48", "E50", "E52", "E54",
    "H30", "H32"
  ];

  var LineNumIncremented = FormSheet.getRange("E44").getValue() + 1;
  
    
  var NumProjects = FormSheet.getRange("E34").getValue();
  var Split = FormSheet.getRange("F42").getValue();

  var Sub1 = FormSheet.getRange("D39").isBlank();
  var Sub2 = FormSheet.getRange("D40").isBlank();
  var Sub3 = FormSheet.getRange("D41").isBlank();

  var Proj1 = FormSheet.getRange("E39").isBlank();
  var Proj2 = FormSheet.getRange("E40").isBlank();
  var Proj3 = FormSheet.getRange("E41").isBlank();

  var Gen1 = FormSheet.getRange("E6").isBlank(); //purchaser name
  var Gen2 = FormSheet.getRange("E8").isBlank(); //receipt number

  var Ven1 = FormSheet.getRange("E14").isBlank(); //vendor name
  var Ven2 = FormSheet.getRange("E16").isBlank(); //date purchased
  var Ven3 = FormSheet.getRange("E18").isBlank(); //date received

  var Fin1 = FormSheet.getRange("H6").isBlank(); //total entries
  var Fin2 = FormSheet.getRange("H8").isBlank(); //total shipping cost
  var Fin3 = FormSheet.getRange("H10").isBlank(); //total tax cost
  var Fin4 = FormSheet.getRange("H12").isBlank(); //total receipt cost
  var Fin5 = FormSheet.getRange("H14").isBlank(); //conversion rate

  var Rec1 = FormSheet.getRange("H18").isBlank(); //receipt link

  var Lin1 = FormSheet.getRange("E30").isBlank(); //rocket type
  var Lin2 = FormSheet.getRange("E32").isBlank(); //project year
  var Lin3 = FormSheet.getRange("E34").isBlank(); //related projs
  var Lin4 = FormSheet.getRange("E36").isBlank(); //system name
  var Lin5 = FormSheet.getRange("E44").isBlank(); //item line num
  var Lin6 = FormSheet.getRange("E46").isBlank(); //item name
  var Lin7 = FormSheet.getRange("H30").isBlank(); //unit cost
  var Lin8 = FormSheet.getRange("H32").isBlank(); //qty

  var Spl1 = FormSheet.getRange("F39").isBlank(); //split1
  var Spl2 = FormSheet.getRange("F40").isBlank(); //split2
  var Spl3 = FormSheet.getRange("F41").isBlank(); //split3


  var OthersNotFilledIn = (
    Gen1 || Gen2 ||
    Ven1 || Ven2 || Ven3 ||
    Fin1 || Fin2 || Fin3 || Fin4 || Fin5 || Rec1 ||
    Lin1 || Lin2 || Lin3 || Lin4 || Lin5 || Lin6 || Lin7 || Lin8
  );

  var OneNotFilledIn = (Sub1 || Proj1 || Spl1);
  var TwoNotFilledIn = (Sub1 || Proj1 || Sub2 || Proj2 || Spl2);
  var ThreeNotFilledIn = (Sub1 || Proj1 || Sub2 || Proj2 || Sub3 || Proj3 || Spl3);


  if (NumProjects == 1 && Split == 1 && OthersNotFilledIn == false && OneNotFilledIn == false) {
      DataSheet.getRange(DataSheet.getLastRow()+1,1,1,31).setValues(values1);
      SpreadsheetApp.getUi().alert("🎉 Data Successfully Added! 🎉", "If you want to make any changes, please edit from the Receipt Input Form Back-End tab. To check grand totals, please go to the Grand Total Checker tab.",SpreadsheetApp.getUi().ButtonSet.OK);
      SpreadsheetApp.getActiveSheet().getRange("E44").setValue(LineNumIncremented);
      resetByRangesList_(FormSheet, FormCells); 
  }

  else if (NumProjects == 2 && Split == 1 && OthersNotFilledIn == false && TwoNotFilledIn == false) {
    DataSheet.getRange(DataSheet.getLastRow()+1,1,1,31).setValues(values1);
    DataSheet.getRange(DataSheet.getLastRow()+1,1,1,31).setValues(values2);
    SpreadsheetApp.getUi().alert("🎉 Data Successfully Added! 🎉", "If you want to make any changes, please edit from the Receipt Input Form Back-End tab. To check grand totals, please go to the Grand Total Checker tab.",SpreadsheetApp.getUi().ButtonSet.OK);
    SpreadsheetApp.getActiveSheet().getRange("E44").setValue(LineNumIncremented);
    resetByRangesList_(FormSheet, FormCells); 
  }

  else if (NumProjects == 3 && Split == 1 && OthersNotFilledIn == false && ThreeNotFilledIn == false) {
    DataSheet.getRange(DataSheet.getLastRow()+1,1,1,31).setValues(values1);
    DataSheet.getRange(DataSheet.getLastRow()+1,1,1,31).setValues(values2);
    DataSheet.getRange(DataSheet.getLastRow()+1,1,1,31).setValues(values3);
    SpreadsheetApp.getUi().alert("🎉 Data Successfully Added! 🎉", "If you want to make any changes, please edit from the Receipt Input Form Back-End tab. To check grand totals, please go to the Grand Total Checker tab.",SpreadsheetApp.getUi().ButtonSet.OK);
    SpreadsheetApp.getActiveSheet().getRange("E44").setValue(LineNumIncremented);
    resetByRangesList_(FormSheet, FormCells); 
  }

  else {
    SpreadsheetApp.getUi().alert("🛑 Data Error! 🛑", "Cannot submit the form because there is not a 100% split between the related projects or all required fields are not filled in. Please edit your form entry and adjust as necessary.",SpreadsheetApp.getUi().ButtonSet.OK);
  }

                
} 


function ResetForm(){
  var ActiveSheet = SpreadsheetApp.getActiveSpreadsheet();
  var FormSheet = ActiveSheet.getSheetByName("Receipt Input Form");
  var FormCells = [
    "E6", "E8",
    "E14", "E16", "E18",
    "H6", "H8", "H10", "H12", "H14", "H18", "H20", "H22",
    "E30", "E34", "E36",
    "D39", "E39", "F39", "D40", "E40", "F40", "D41", "E41", "F41",
    "E44", "E46", "E48", "E50", "E52", "E54",
    "H30", "H32"
  ];
  resetByRangesList_(FormSheet, FormCells);
}

function resetByRangesList_(FormSheet, FormCells){
  FormSheet.getRangeList(FormCells).clearContent();
}

