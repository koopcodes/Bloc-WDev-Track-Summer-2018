// onOpen is called when a Google Document is opened by the user.
function onOpen() {
  // Install menu 'FORECAST' with menu item 'RUN REPORT'
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  spreadsheet.addMenu('FORECAST', [{ name : 'RUN REPORT', functionName : 'go2'}]);
}

// Display spreadsheet key -- Inactive -- add to FORECAST menu or run Tools > Script Manager to display ID of current spreadsheet
function displaySpreadsheetId()
{
  var id = SpreadsheetApp.getActiveSpreadsheet().getId();
  var nbsp = unescape('%A0'); // no-break space
  var space1 = '';
  for (var i=0; i<60; i++) space1 += nbsp;
  var youCanUseItLikeThis = 'You'+nbsp+'can'+nbsp+'use'+nbsp+'it'+nbsp+'like'+nbsp+'this:';
  var space2 = space1+space1+space1;
  var msg = "The Spreadsheet ID is '"+id+"'."+space1+' '+
      '';
  Browser.msgBox('Spreadsheet ID', msg, Browser.Buttons.OK);
}

function go2() {

  getInitial();
  createNewSheet();
  getEDOAdata();
  getEDODdata();
  sort();
  updateCounts();
  switchSheet();
  caseload();
  prettySheet()
  goHome()

}

//sourcefile is EDOD_DO_NOT_DELETE  destfile is Admissions Forecast
function getInitial() {
  var sourcefile = SpreadsheetApp.openById("t6rpTfC3dx78qhurAS05KDw");
  var lastRow = sourcefile.getLastRow();
  var destfile = SpreadsheetApp.openById("tjye2A4l4U9zf8Ae9fKXyLA");
  iTotalCount = (lastRow-1);
  iAdCount = 0, iYaCount = 0, iADMCount = 0, iADFCount = 0, iYAMCount = 0, iYAFCount = 0, iMautzCount = 0, iMautzADCount = 0, iMautzYACount = 0, iMosesCount = 0, iMosesADCount = 0, iMosesYACount = 0, iAdamsCount = 0, iAdamsADCount = 0, iAdamsYACount = 0, iSullivanCount = 0, iSullivanADCount = 0, iSullivanYACount = 0, iKasenchakCount = 0, iKasenchakADCount = 0, iKasenchakYACount = 0, iWeldCount = 0, iWeldADCount = 0, iWeldYACount = 0, iJensenCount=0, iJensenADCount=0, iJensenYACount=0, iBrinkCount=0, iBrinkADCount=0, iBrinkYACount=0, iJulCount=0, iJulADCount=0, iJulYACount=0; iDunnCount=0, iDunnADCount=0, iDunnYACount=0, iHinojosaCount = 0, iHinojosaADCount = 0, iHinojosaYACount = 0, iGlassCount = 0, iGlassADCount = 0, iGlassYACount = 0, iKonikCount = 0, iKonikADCount = 0, iKonikYACount = 0;


  for (i=2; i<lastRow+1; i++) {

      var edodtag = sourcefile.getRange("I"+i).getValue();
      var gendertag =sourcefile.getRange("C"+i).getValue();

      if (edodtag.match(", YA") && gendertag.match("M")){
        iYaCount++;
        iYAMCount++;
      }

      if (edodtag.match(", YA") && gendertag.match("F")){
            iYaCount++;
            iYAFCount++;
      }

      if (!edodtag.match(", YA") && gendertag.match("M")){
            	iAdCount++;
            	iADMCount++;
      }

      if (!edodtag.match(", YA") && gendertag.match("F")){
                	iAdCount++;
                	iADFCount++;
      }

    var therapRange = sourcefile.getRange("G"+i).getValues();
    if (therapRange == "Toby Mautz, MSW, LCSW") {
        iMautzCount++;
		if (edodtag.match(", YA")){
		   iMautzYACount++;
		   }  else
		   {
		   iMautzADCount++;
		   }
        }

    if (therapRange == "Hilary Moses, MSW, LCSW") {
        iMosesCount++;
		if (edodtag.match(", YA")){
		   iMosesYACount++;
		   } else
		   {
		   iMosesADCount++;
		   }
        }

    if (therapRange == "Jason R.K. Adams, Ph.D.") {
        iAdamsCount++;
		if (edodtag.match(", YA")){
		   iAdamsYACount++;
		   } else
		   {
		   iAdamsADCount++;
		   }
        }

    if (therapRange == "Mike Sullivan, MA") {
        iSullivanCount++;
		if (edodtag.match(", YA")){
		    iSullivanYACount++;
		    } else
		    {
		    iSullivanADCount++;
		    }
        }

    if (therapRange == "Kathryn Kasenchak, Psy.D.") {
        iKasenchakCount++;
		if (edodtag.match(", YA")){
		    iKasenchakYACount++;
		    } else
		    {
		    iKasenchakADCount++;
		    }
        }

    if (therapRange == "Kelly Weld, M.A., MFT") {
	iWeldCount++;
		if (edodtag.match(", YA")){
		    iWeldYACount++;
		    } else
		    {
		    iWeldADCount++;
		    }
        }
    if (therapRange == "Bridger Jensen, MA, CTRS") {
    	iJensenCount++;
    		if (edodtag.match(", YA")){
		    iJensenYACount++;
		    } else
		    {
		    iJensenADCount++;
		    }
    }
    if (therapRange == "Haley Brink, MS, LMFTA") {
       	iBrinkCount++;
		   if (edodtag.match(", YA")){
	    	   iBrinkYACount++;
	    	   } else
	    	   {
	    	   iBrinkADCount++;
            	   }
    }
    if (therapRange == "Erik Jul") {
        iJulCount++;
           	   if (edodtag.match(", YA")){
		   iJulYACount++;
		   } else
		   {
		   iJulADCount++;
		   }
    }
    if (therapRange == "Mark Dunn") {
        iDunnCount++;
           	   if (edodtag.match(", YA")){
		   iDunnYACount++;
		   } else
		   {
		   iDunnADCount++;
		   }
    }
	if (therapRange == "Yvette Hinojosa, MA") {
	    iHinojosaCount++;
	       	   if (edodtag.match(", YA")){
		   iHinojosaYACount++;
		   } else
		   {
		   iHinojosaADCount++;
		   }
	}
	if (therapRange == "Annette Glass, LMFT") {
	    iGlassCount++;
	       	   if (edodtag.match(", YA")){
		   iGlassYACount++;
		   } else
		   {
		   iGlassADCount++;
		   }
	}
	if (therapRange == "Brian Konik, Ph.D") {
	    iKonikCount++;
	       	   if (edodtag.match(", YA")){
		   iKonikYACount++;
		   } else
		   {
		   iKonikADCount++;
		   }
	}
  }

}


function createNewSheet() {
  var ss = SpreadsheetApp.openById("tjye2A4l4U9zf8Ae9fKXyLA");
  var sheet = ss.getSheets()[0];
  var today = new Date();
  var todayformat = Utilities.formatDate(today,'GMT-8',"MM/dd/yyyy' @ 'HH:mm:ss' PST'");
  var user = Session.getActiveUser().getUsername();
  var m = sheet.getMaxRows();
  sheet.insertRowsAfter(m, 50);
  sheet.clear();
  sheet.getRange("H1").setValue("Last Run on: "+todayformat+" by "+user);
  sheet.getRange("A1").setValue("Initial Counts:");
  sheet.getRange("B1").setValue("Toby Mautz: "+iMautzCount);
  sheet.getRange("C1").setValue("Hilary Moses: "+iMosesCount);
  sheet.getRange("D1").setValue("Jason Adams: "+iAdamsCount);
  sheet.getRange("E1").setValue("Mike Sullivan: "+iSullivanCount);
  sheet.getRange("A2").setValue("Kathryn Kasenchak: "+iKasenchakCount);
  sheet.getRange("B2").setValue("Kelly Weld: "+iWeldCount).setFontColor("#0022CC");
  sheet.getRange("C2").setValue("Bridger Jensen: "+iJensenCount);
  sheet.getRange("D2").setValue("Haley Brink: "+iBrinkCount);
  sheet.getRange("E2").setValue("Erik Jul: "+iJulCount).setFontColor("#0022CC");
  sheet.getRange("A3").setValue("Mark Dunn: "+iDunnCount).setFontColor("#0022CC");
  sheet.getRange("B3").setValue("Yvette Hinojosa: "+iHinojosaCount);
  sheet.getRange("C3").setValue("Annette Glass: "+iGlassCount).setFontColor("#0022CC");
  sheet.getRange("D3").setValue("Brian Konik: "+iKonikCount);
  sheet.getRange("L2").setValue("AD: "+iAdCount);
  sheet.getRange("M2").setValue("YA: "+iYaCount);
  sheet.getRange("N2").setValue("ADM: "+iADMCount);
  sheet.getRange("O2").setValue("ADF: "+iADFCount);
  sheet.getRange("P2").setValue("YAM: "+iYAMCount);
  sheet.getRange("Q2").setValue("YAF: "+iYAFCount);
  sheet.getRange("R2").setValue("Initial Total: "+iTotalCount);


  if (+iYaCount + +iAdCount != +iTotalCount) {
    Browser.msgBox("YA + AD does not equal Total");
  }
  if (+iADMCount + +iADFCount + +iYAMCount + +iYAFCount != +iTotalCount) {
    Browser.msgBox("ADM + ADF + YAM + YAF does not equal Total");
  }
   if (+iMautzCount + +iMosesCount + +iAdamsCount + +iSullivanCount + +iKasenchakCount + +iWeldCount + +iJensenCount + +iBrinkCount + +iJulCount + +iDunnCount + +iHinojosaCount + +iGlassCount + +iKonikCount != +iTotalCount) {
    Browser.msgBox("Therapist totals do not equal Total");
  }
  sheet.getRange("A4").setValue("Date");
  sheet.getRange("B4").setValue("Date Type");
  sheet.getRange("C4").setValue("First Name");
  sheet.getRange("D4").setValue("Last Name");
  sheet.getRange("E4").setValue("Adms. Rep");
  sheet.getRange("F4").setValue("Age");
  sheet.getRange("G4").setValue("Gender");
  sheet.getRange("H4").setValue("Therapist");
  sheet.getRange("I4").setValue("Th. Count");
  sheet.getRange("J4").setValue("Th. AD Count");
  sheet.getRange("K4").setValue("Th. YA Count");
  sheet.getRange("L4").setValue("AD");
  sheet.getRange("M4").setValue("YA");
  sheet.getRange("N4").setValue("ADM");
  sheet.getRange("O4").setValue("ADF");
  sheet.getRange("P4").setValue("YAM");
  sheet.getRange("Q4").setValue("YAF");
  sheet.getRange("R4").setValue("Total");
  sheet.getRange("S4").setValue("Referrer");
  sheet.getRange("T4").setValue("YA Flag");
  sheet.getRange("U4").setValue("DOA");
  sheet.getRange("V4").setValue("LOS");
}

//sourcefile is EDOA_DO_NOT_DELTE destfile is Admissions Forecast
function getEDOAdata() {
  var sourcefile = SpreadsheetApp.openById("tlGKQ7-xAAMc4XSWt58RxLw");
  var destfile = SpreadsheetApp.openById("tjye2A4l4U9zf8Ae9fKXyLA");
  var lastRow = sourcefile.getLastRow()

  var dates = sourcefile.getRange("G2:G"+lastRow).getValues();
  destfile.getRange("A5:A"+(+lastRow + 3)).setValues(dates);

  var names = sourcefile.getRange("A2:B"+lastRow).getValues();
  destfile.getRange("C5:D"+(+lastRow + 3)).setValues(names);

  var gender = sourcefile.getRange("C2:C"+lastRow).getValues();
  destfile.getRange("G5:G"+(+lastRow + 3)).setValues(gender);

  var age = sourcefile.getRange("E2:E"+lastRow).getValues();
  destfile.getRange("F5:F"+(+lastRow + 3)).setValues(age);

  var therapist = sourcefile.getRange("H2:H"+lastRow).getValues();
  destfile.getRange("H5:H"+(+lastRow + 3)).setValues(therapist);

  var ref = sourcefile.getRange("I2:I"+lastRow).getValues();
  destfile.getRange("S5:S"+(+lastRow + 3)).setValues(ref);

  destfile.getRange("B5:B"+(+lastRow + 3)).setValue('Arrival');

  for (i=2; i<(lastRow+1); i++) {

    var tags = sourcefile.getRange("J"+i).getValue();
    var newDate = sourcefile.getRange("K"+i).getValues();

    if (tags.match("Likely")) {
      destfile.getRange("A"+(i+3)).setFontColor("#ffa500");
      destfile.getRange("A"+(i+3)).setFontWeight("bold");
    }

    if (newDate != "") {
      destfile.getRange("A"+(i+3)).setValues(newDate);
      destfile.getRange("A"+(i+3)).setFontColor("#00CC00");
      destfile.getRange("A"+(i+3)).setFontWeight("bold");
    }

    if (tags.match(", YA")) {
      destfile.getRange("T"+(i+3)).setValue("YA");
    } else {
      destfile.getRange("T"+(i+3)).setValue("AD");
    }

    if (tags.match("Rob")) {
            	destfile.getRange("E"+(i+3)).setValue("Rob");
        }
    if (tags.match("Erin")) {
            	destfile.getRange("E"+(i+3)).setValue("Erin");
        }
    if (tags.match("Mark A")) {
            	destfile.getRange("E"+(i+3)).setValue("Mark A");
        }
    if (tags.match("Denise")) {
            	destfile.getRange("E"+(i+3)).setValue("Denise");
        }
    if (tags.match("Lori")) {
            	destfile.getRange("E"+(i+3)).setValue("Lori");
        }
    if (tags.match("JD")) {
                destfile.getRange("E"+(i+3)).setValue("JD");
    }
  }
}

// sourcefile is EDOD_DO_NOT_DELETE  destfile is Admissions Forecast
function getEDODdata() {
  var sourcefile = SpreadsheetApp.openById("t6rpTfC3dx78qhurAS05KDw");
  var destfile = SpreadsheetApp.openById("tjye2A4l4U9zf8Ae9fKXyLA");
  var lastRow = sourcefile.getLastRow();
  var lastRowDest = destfile.getLastRow();

  var dates = sourcefile.getRange("E2:E"+lastRow).getValues();
  destfile.getRange("A"+(lastRowDest+1)+":A"+(lastRowDest+lastRow-1)).setValues(dates);

  var doa = sourcefile.getRange("F2:F"+lastRow).getValues();
  destfile.getRange("U"+(lastRowDest+1)+":U"+(lastRowDest+lastRow-1)).setValues(doa);

  var names = sourcefile.getRange("A2:B"+lastRow).getValues();
  destfile.getRange("C"+(lastRowDest+1)+":D"+(lastRow+lastRowDest-1)).setValues(names);

  var gender = sourcefile.getRange("C2:C"+lastRow).getValues();
  destfile.getRange("G"+(lastRowDest+1)+":G"+(lastRow+lastRowDest-1)).setValues(gender);

  var age = sourcefile.getRange("D2:D"+lastRow).getValues();
  destfile.getRange("F"+(lastRowDest+1)+":F"+(lastRow+lastRowDest-1)).setValues(age);

  var therapist = sourcefile.getRange("G2:G"+lastRow).getValues();
  destfile.getRange("H"+(lastRowDest+1)+":H"+(lastRow+lastRowDest-1)).setValues(therapist);

  var ref = sourcefile.getRange("H2:H"+lastRow).getValues();
  destfile.getRange("S"+(lastRowDest+1)+":S"+(lastRow+lastRowDest-1)).setValues(ref);

  destfile.getRange("B"+(lastRowDest+1)+":B"+(lastRow+lastRowDest-1)).setValue('Departure');

  for (i=lastRowDest+1; i<(lastRow+lastRowDest); i++) {
    if (sourcefile.getRange("J"+(i-lastRowDest+1)).getValues() !="") {
      var newDate = sourcefile.getRange("J"+(i-lastRowDest+1)).getValues();
      destfile.getRange("A"+i).setValues(newDate);
      destfile.getRange("A"+i).setFontColor("#CC0000");
      destfile.getRange("A"+i).setFontWeight("bold");
    }

    var tags = sourcefile.getRange("I"+(i-lastRowDest+1)).getValue();
    if (tags.match(", YA")) {
      destfile.getRange("T"+i).setValue("YA");
    } else {
       destfile.getRange("T"+i).setValue("AD");
    }

    if (tags.match("Rob")) {
        	destfile.getRange("E"+i).setValue("Rob");
    }
    if (tags.match("Erin")) {
        	destfile.getRange("E"+i).setValue("Erin");
    }
    if (tags.match("Mark A")) {
        	destfile.getRange("E"+i).setValue("Mark A");
    }
    if (tags.match("Denise")) {
        	destfile.getRange("E"+i).setValue("Denise");
    }
    if (tags.match("Lori")) {
        	destfile.getRange("E"+i).setValue("Lori");
    }
    if (tags.match("JD")) {
                destfile.getRange("E"+i).setValue("JD");
    }
  }
}

function sort() {

    var ss = SpreadsheetApp.openById("tjye2A4l4U9zf8Ae9fKXyLA");
    var lastRow = ss.getLastRow();
    var range = ss.getRange("A5:V"+lastRow);
    range.sort([{column: 1, ascending: true}, {column: 2, ascending: true}]);

}

// sourcefile is EDOA_DO_NOT_DELETE ss is Admissions Forecast
function updateCounts() {
  var TotalCount=iTotalCount;
  var YaCount=iYaCount;
  var AdCount=iAdCount;
  var ADMCount=iADMCount;
  var ADFCount=iADFCount;
  var YAMCount=iYAMCount;
  var YAFCount=iYAFCount;
  var MautzCount=iMautzCount;
  var MautzADCount=iMautzADCount;
  var MautzYACount=iMautzYACount;
  var MosesCount=iMosesCount;
  var MosesADCount=iMosesADCount;
  var MosesYACount=iMosesYACount;
  var AdamsCount=iAdamsCount;
  var AdamsADCount=iAdamsADCount;
  var AdamsYACount=iAdamsYACount;
  var SullivanCount=iSullivanCount;
  var SullivanADCount=iSullivanADCount;
  var SullivanYACount=iSullivanYACount;
  var KasenchakCount=iKasenchakCount;
  var KasenchakADCount=iKasenchakADCount;
  var KasenchakYACount=iKasenchakYACount;
  var WeldCount=iWeldCount;
  var WeldADCount=iWeldADCount;
  var WeldYACount=iWeldYACount;
  var JensenCount=iJensenCount;
  var JensenADCount=iJensenADCount;
  var JensenYACount=iJensenYACount;
  var BrinkCount=iBrinkCount;
  var BrinkADCount=iBrinkADCount;
  var BrinkYACount=iBrinkYACount;
  var JulCount=iJulCount;
  var JulADCount=iJulADCount;
  var JulYACount=iJulYACount;
  var DunnCount=iDunnCount;
  var DunnADCount=iDunnADCount;
  var DunnYACount=iDunnYACount;
  var HinojosaCount=iHinojosaCount;
  var HinojosaADCount=iHinojosaADCount;
  var HinojosaYACount=iHinojosaYACount;
  var GlassCount=iGlassCount;
  var GlassADCount=iGlassADCount;
  var GlassYACount=iGlassYACount;
  var KonikCount=iKonikCount;
  var KonikADCount=iKonikADCount;
  var KonikYACount=iKonikYACount;
  var MautzLimit = 0, MosesLimit = 0, AdamsLimit = 0, SullivanLimit = 0, KasenchakLimit = 0, WeldLimit = 0, JensenLimit = 0, BrinkLimit = 0, JulLimit = 0, DunnLimit = 0, HinojosaLimit = 0, GlassLimit = 0, KonikLimit = 0;

  var ss = SpreadsheetApp.openById("tjye2A4l4U9zf8Ae9fKXyLA");
  var sheet = ss.getSheetByName('AdmissionsForecast');
  var lastRow = ss.getLastRow();
  var sourcefile = SpreadsheetApp.openById("t6rpTfC3dx78qhurAS05KDw");
  var srclastRow = sourcefile.getLastRow();
  var limitfile = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TherapistLimits");
  var limitlastRow = limitfile.getLastRow();


  for (i=2; i<(limitlastRow+1); i++) {
      var therapistName = limitfile.getRange("A"+i).getValue();
      var therapistLimit = limitfile.getRange("B"+i).getValue();

      switch(therapistName)
      {
      	case "Bridger Jensen, MA, CTRS":
      		JensenLimit= therapistLimit;
      	break;

      	case "Erik Jul":
      		JulLimit= therapistLimit;
		break;

      	case "Haley Brink, MS, LMFTA":
      		BrinkLimit= therapistLimit;
 		break;

      	case "Hilary Moses, MSW, LCSW":
      		MosesLimit= therapistLimit;
      	break;

      	case "Jason R.K. Adams, Ph.D.":
      		AdamsLimit= therapistLimit;
      	break;

      	case "Kathryn Kasenchak, Psy.D.":
      		KasenchakLimit= therapistLimit;
      	break;

      	case "Kelly Weld, M.A., MFT":
      		WeldLimit= therapistLimit;
      	break;

      	case "Mike Sullivan, MA":
      		SullivanLimit= therapistLimit;
      	break;

      	case "Toby Mautz, MSW, LCSW":
      		MautzLimit= therapistLimit;
      	break;

      	case "Mark Dunn":
      		DunnLimit= therapistLimit;
      	break;

      	case "Yvette Hinojosa, MA":
      		HinojosaLimit= therapistLimit;
      	break;

      	case "Annette Glass, LMFT":
      		GlassLimit= therapistLimit;
      	break;

      	case "Brian Konik, Ph.D":
      		KonikLimit= therapistLimit;
      	break;

      }
  }


  for (i=5; i<(lastRow+1); i++) {
      var age = ss.getRange("T"+i).getValue();
      var gender = ss.getRange("G"+i).getValue();
      var status = ss.getRange("B"+i).getValue();
      var therapist = ss.getRange("H"+i).getValue();
      var doa = ss.getRange("U"+i).getValue();
      var dod = ss.getRange("A"+i).getValue();

      if (doa != "") {
      	var los = ((dod - doa)/86400000);
      	var losr = los.toPrecision(2);
        ss.getRange("V"+i).setValue(losr);
      }

      if (status=="Arrival") {

        sheet.getRange(i,1,1,18).setFontWeight("bold")

        if (age == "AD") {
          AdCount++;
        } else {
          YaCount++;
          ss.getRange("C"+i).setFontColor("#0022cc");
          ss.getRange("D"+i).setFontColor("#0022cc");
        }

        ss.getRange("L"+i).setValue(AdCount);
        ss.getRange("M"+i).setValue(YaCount);

        if (gender == "M" && age == "AD") {
	              ADMCount++;
	    }
	if (gender == "F" && age == "AD") {
	              ADFCount++;
	    }
	if (gender == "M" && age == "YA") {
	              YAMCount++;
	    }
	if (gender == "F" && age == "YA") {
	              YAFCount++;
    	    }

        ss.getRange("N"+i).setValue(ADMCount);
        ss.getRange("O"+i).setValue(ADFCount);
        ss.getRange("P"+i).setValue(YAMCount);
        ss.getRange("Q"+i).setValue(YAFCount);

      switch(therapist)
      {
        case "Toby Mautz, MSW, LCSW":
          MautzCount++;
          if (age == "YA"){
          	MautzYACount++;
          	} else
          	{
          	MautzADCount++;
          }
          ss.getRange("I"+i).setValue(MautzCount);
          ss.getRange("J"+i).setValue(MautzADCount);
          ss.getRange("K"+i).setValue(MautzYACount);
          if (MautzCount > MautzLimit) {ss.getRange("I"+i).setBackground("#CC0000"); ss.getRange("I"+i).setFontWeight("bold")};
          if (MautzCount == MautzLimit) {ss.getRange("I"+i).setBackground("#CC0000"); ss.getRange("I"+i).setFontWeight("bold")};
        break;

        case "Hilary Moses, MSW, LCSW":
          MosesCount++;
          if (age == "YA"){
	         MosesYACount++;
	         } else
	         {
	         MosesADCount++;
          }
          ss.getRange("I"+i).setValue(MosesCount);
          ss.getRange("J"+i).setValue(MosesADCount);
          ss.getRange("K"+i).setValue(MosesYACount);
          if (MosesCount > MosesLimit) {ss.getRange("I"+i).setBackground("#CC0000"); ss.getRange("I"+i).setFontWeight("bold")};
          if (MosesCount == MosesLimit) {ss.getRange("I"+i).setBackground("#CC0000"); ss.getRange("I"+i).setFontWeight("bold")};
        break;

        case "Jason R.K. Adams, Ph.D.":
          AdamsCount++;
          if (age == "YA"){
	          AdamsYACount++;
	          } else
	          {
	          AdamsADCount++;
          }
          ss.getRange("I"+i).setValue(AdamsCount);
          ss.getRange("J"+i).setValue(AdamsADCount);
          ss.getRange("K"+i).setValue(AdamsYACount);
          if (AdamsCount > AdamsLimit) {ss.getRange("I"+i).setBackground("#CC0000"); ss.getRange("I"+i).setFontWeight("bold")};
          if (AdamsCount == AdamsLimit) {ss.getRange("I"+i).setBackground("#CC0000"); ss.getRange("I"+i).setFontWeight("bold")};
        break;

        case "Mike Sullivan, MA":
          SullivanCount++;
          if (age == "YA"){
		  SullivanYACount++;
		  } else
		  {
		  SullivanADCount++;
          }
          ss.getRange("I"+i).setValue(SullivanCount);
          ss.getRange("J"+i).setValue(SullivanADCount);
          ss.getRange("K"+i).setValue(SullivanYACount);
          if (SullivanCount > SullivanLimit) {ss.getRange("I"+i).setBackground("#CC0000"); ss.getRange("I"+i).setFontWeight("bold")};
          if (SullivanCount == SullivanLimit) {ss.getRange("I"+i).setBackground("#CC0000"); ss.getRange("I"+i).setFontWeight("bold")};
        break;

        case "Kathryn Kasenchak, Psy.D.":
          KasenchakCount++;
          if (age == "YA"){
		  KasenchakYACount++;
		  } else
		  {
		  KasenchakADCount++;
          }
          ss.getRange("I"+i).setValue(KasenchakCount);
          ss.getRange("J"+i).setValue(KasenchakADCount);
          ss.getRange("K"+i).setValue(KasenchakYACount);
          if (KasenchakCount > KasenchakLimit) {ss.getRange("I"+i).setBackground("#CC0000"); ss.getRange("I"+i).setFontWeight("bold")};
          if (KasenchakCount == KasenchakLimit) {ss.getRange("I"+i).setBackground("#CC0000"); ss.getRange("I"+i).setFontWeight("bold")};
        break;

        case "Kelly Weld, M.A., MFT":
          WeldCount++;
          if (age == "YA"){
		  WeldYACount++;
		  } else
		  {
		  WeldADCount++;
          }
          ss.getRange("I"+i).setValue(WeldCount);
          ss.getRange("J"+i).setValue(WeldADCount);
          ss.getRange("K"+i).setValue(WeldYACount);
          if (WeldCount > WeldLimit) {ss.getRange("I"+i).setBackground("#CC0000"); ss.getRange("I"+i).setFontWeight("bold")};
          if (WeldCount == WeldLimit) {ss.getRange("I"+i).setBackground("#CC0000"); ss.getRange("I"+i).setFontWeight("bold")};
        break;

        case "Bridger Jensen, MA, CTRS":
          JensenCount++;
          if (age == "YA"){
		  JensenYACount++;
		  } else
		  {
		  JensenADCount++;
          }
          ss.getRange("I"+i).setValue(JensenCount);
          ss.getRange("J"+i).setValue(JensenADCount);
          ss.getRange("K"+i).setValue(JensenYACount);
          if (JensenCount > JensenLimit) {ss.getRange("I"+i).setBackground("#CC0000"); ss.getRange("I"+i).setFontWeight("bold")};
          if (JensenCount == JensenLimit) {ss.getRange("I"+i).setBackground("#CC0000"); ss.getRange("I"+i).setFontWeight("bold")};
        break;

        case "Haley Brink, MS, LMFTA":
	  BrinkCount++;
	  if (age == "YA"){
		  BrinkYACount++;
		  } else
		  {
		  BrinkADCount++;
          }
	  ss.getRange("I"+i).setValue(BrinkCount);
	  ss.getRange("J"+i).setValue(BrinkADCount);
	  ss.getRange("K"+i).setValue(BrinkYACount);
	  if (BrinkCount > BrinkLimit) {ss.getRange("I"+i).setBackground("#CC0000"); ss.getRange("I"+i).setFontWeight("bold")};
	  if (BrinkCount == BrinkLimit) {ss.getRange("I"+i).setBackground("#CC0000"); ss.getRange("I"+i).setFontWeight("bold")};
        break;

	case "Erik Jul":
	  JulCount++;
	  if (age == "YA"){
		  JulYACount++;
		  } else
		  {
		  JulADCount++;
          }
	  ss.getRange("I"+i).setValue(JulCount);
	  ss.getRange("J"+i).setValue(JulADCount);
	  ss.getRange("K"+i).setValue(JulYACount);
	  if (JulCount > JulLimit) {ss.getRange("I"+i).setBackground("#CC0000"); ss.getRange("I"+i).setFontWeight("bold")};
	  if (JulCount == JulLimit) {ss.getRange("I"+i).setBackground("#CC0000"); ss.getRange("I"+i).setFontWeight("bold")};
        break;

	case "Mark Dunn":
	  DunnCount++;
	  if (age == "YA"){
		  DunnYACount++;
		  } else
		  {
		  DunnADCount++;
	      }
	  ss.getRange("I"+i).setValue(DunnCount);
	  ss.getRange("J"+i).setValue(DunnADCount);
	  ss.getRange("K"+i).setValue(DunnYACount);
	  if (DunnCount > DunnLimit) {ss.getRange("I"+i).setBackground("#CC0000"); ss.getRange("I"+i).setFontWeight("bold")};
	  if (DunnCount == DunnLimit) {ss.getRange("I"+i).setBackground("#CC0000"); ss.getRange("I"+i).setFontWeight("bold")};
	    break;

	case "Yvette Hinojosa, MA":
	  HinojosaCount++;
	  if (age == "YA"){
		  HinojosaYACount++;
		  } else
		  {
		  HinojosaADCount++;
	      }
	  ss.getRange("I"+i).setValue(HinojosaCount);
	  ss.getRange("J"+i).setValue(HinojosaADCount);
	  ss.getRange("K"+i).setValue(HinojosaYACount);
	  if (HinojosaCount > HinojosaLimit) {ss.getRange("I"+i).setBackground("#CC0000"); ss.getRange("I"+i).setFontWeight("bold")};
	  if (HinojosaCount == HinojosaLimit) {ss.getRange("I"+i).setBackground("#CC0000"); ss.getRange("I"+i).setFontWeight("bold")};
	    break;

	case "Annette Glass, LMFT":
	  GlassCount++;
	  if (age == "YA"){
		  GlassYACount++;
		  } else
		  {
		  GlassADCount++;
	      }
	  ss.getRange("I"+i).setValue(GlassCount);
	  ss.getRange("J"+i).setValue(GlassADCount);
	  ss.getRange("K"+i).setValue(GlassYACount);
	  if (GlassCount > GlassLimit) {ss.getRange("I"+i).setBackground("#CC0000"); ss.getRange("I"+i).setFontWeight("bold")};
	  if (GlassCount == GlassLimit) {ss.getRange("I"+i).setBackground("#CC0000"); ss.getRange("I"+i).setFontWeight("bold")};
	    break;

	case "Brian Konik, Ph.D":
	  KonikCount++;
	  if (age == "YA"){
		  KonikYACount++;
		  } else
		  {
		  KonikADCount++;
	      }
	  ss.getRange("I"+i).setValue(KonikCount);
	  ss.getRange("J"+i).setValue(KonikADCount);
	  ss.getRange("K"+i).setValue(KonikYACount);
	  if (KonikCount > KonikLimit) {ss.getRange("I"+i).setBackground("#CC0000"); ss.getRange("I"+i).setFontWeight("bold")};
	  if (KonikCount == KonikLimit) {ss.getRange("I"+i).setBackground("#CC0000"); ss.getRange("I"+i).setFontWeight("bold")};
	    break;

      }

      TotalCount++;
      ss.getRange("R"+i).setValue(TotalCount);



    } else {
   if (age == "AD") {
       AdCount--;
      } else {
        YaCount--;
        ss.getRange("C"+i).setFontColor("#0022cc");
        ss.getRange("D"+i).setFontColor("#0022cc");
      }

      ss.getRange("L"+i).setValue(AdCount);
      ss.getRange("M"+i).setValue(YaCount);

      if (gender == "M" && age == "AD") {
	              ADMCount--;
	    }
	if (gender == "F" && age == "AD") {
	              ADFCount--;
	    }
	if (gender == "M" && age == "YA") {
	              YAMCount--;
	    }
	if (gender == "F" && age == "YA") {
	              YAFCount--;
    	    }

        ss.getRange("N"+i).setValue(ADMCount);
        ss.getRange("O"+i).setValue(ADFCount);
        ss.getRange("P"+i).setValue(YAMCount);
        ss.getRange("Q"+i).setValue(YAFCount);

      switch(therapist)
      {
        case "Toby Mautz, MSW, LCSW":
          MautzCount--;if (age == "YA"){
          	MautzYACount--;
          	} else
          	{
          	MautzADCount--;
          }
          ss.getRange("I"+i).setValue(MautzCount);
          ss.getRange("J"+i).setValue(MautzADCount);
          ss.getRange("K"+i).setValue(MautzYACount);
          if (MautzCount > (MautzLimit-1)) {ss.getRange("I"+i).setBackground("#CC0000"); ss.getRange("I"+i).setFontWeight("bold")};
          if (MautzCount == (MautzLimit-1)) {ss.getRange("I"+i).setBackground("#00CC00"); ss.getRange("I"+i).setFontWeight("bold")};
        break;

        case "Hilary Moses, MSW, LCSW":
          MosesCount--;
          if (age == "YA"){
		 MosesYACount--;
		 } else
		 {
		 MosesADCount--;
	  }
	  ss.getRange("I"+i).setValue(MosesCount);
	  ss.getRange("J"+i).setValue(MosesADCount);
          ss.getRange("K"+i).setValue(MosesYACount);
          if (MosesCount > (MosesLimit-1)) {ss.getRange("I"+i).setBackground("#CC0000"); ss.getRange("I"+i).setFontWeight("bold")};
          if (MosesCount == (MosesLimit-1)) {ss.getRange("I"+i).setBackground("#00CC00"); ss.getRange("I"+i).setFontWeight("bold")};
        break;

        case "Jason R.K. Adams, Ph.D.":
          AdamsCount--;
          if (age == "YA"){
	          AdamsYACount--;
	          } else
	          {
	          AdamsADCount--;
          }
          ss.getRange("I"+i).setValue(AdamsCount);
          ss.getRange("J"+i).setValue(AdamsADCount);
          ss.getRange("K"+i).setValue(AdamsYACount);
          if (AdamsCount > (AdamsLimit-1)) {ss.getRange("I"+i).setBackground("#CC0000"); ss.getRange("I"+i).setFontWeight("bold")};
          if (AdamsCount == (AdamsLimit-1)) {ss.getRange("I"+i).setBackground("#00CC00"); ss.getRange("I"+i).setFontWeight("bold")};
        break;

        case "Mike Sullivan, MA":
          SullivanCount--;
          if (age == "YA"){
		  SullivanYACount--;
		  } else
		  {
		  SullivanADCount--;
	  }
	  ss.getRange("I"+i).setValue(SullivanCount);
	  ss.getRange("J"+i).setValue(SullivanADCount);
          ss.getRange("K"+i).setValue(SullivanYACount);
          if (SullivanCount > (SullivanLimit-1)) {ss.getRange("I"+i).setBackground("#CC0000"); ss.getRange("I"+i).setFontWeight("bold")};
          if (SullivanCount == (SullivanLimit-1)) {ss.getRange("I"+i).setBackground("#00CC00"); ss.getRange("I"+i).setFontWeight("bold")};
        break;

        case "Kathryn Kasenchak, Psy.D.":
          KasenchakCount--;
          if (age == "YA"){
		  KasenchakYACount--;
		  } else
		  {
		  KasenchakADCount--;
          }
          ss.getRange("I"+i).setValue(KasenchakCount);
          ss.getRange("J"+i).setValue(KasenchakADCount);
          ss.getRange("K"+i).setValue(KasenchakYACount);
          if (KasenchakCount > (KasenchakLimit-1)) {ss.getRange("I"+i).setBackground("#CC0000"); ss.getRange("I"+i).setFontWeight("bold")};
          if (KasenchakCount == (KasenchakLimit-1)) {ss.getRange("I"+i).setBackground("#00CC00"); ss.getRange("I"+i).setFontWeight("bold")};
        break;

        case "Kelly Weld, M.A., MFT":
          WeldCount--;
          if (age == "YA"){
		  WeldYACount--;
		  } else
		  {
		  WeldADCount--;
	  }
	  ss.getRange("I"+i).setValue(WeldCount);
	  ss.getRange("J"+i).setValue(WeldADCount);
          ss.getRange("K"+i).setValue(WeldYACount);
          if (WeldCount > (WeldLimit-1)) {ss.getRange("I"+i).setBackground("#CC0000"); ss.getRange("I"+i).setFontWeight("bold")};
          if (WeldCount == (WeldLimit-1)) {ss.getRange("I"+i).setBackground("#00CC00"); ss.getRange("I"+i).setFontWeight("bold")};
        break;

        case "Bridger Jensen, MA, CTRS":
           JensenCount--;
           if (age == "YA"){
		  JensenYACount--;
		  } else
		  {
		  JensenADCount--;
	   }
	   ss.getRange("I"+i).setValue(JensenCount);
	   ss.getRange("J"+i).setValue(JensenADCount);
           ss.getRange("K"+i).setValue(JensenYACount);
           ss.getRange("I"+i).setValue(JensenCount);
           if (JensenCount > (JensenLimit-1)) {ss.getRange("I"+i).setBackground("#CC0000"); ss.getRange("I"+i).setFontWeight("bold")};
           if (JensenCount == (JensenLimit-1)) {ss.getRange("I"+i).setBackground("#00CC00"); ss.getRange("I"+i).setFontWeight("bold")};
        break;

        case "Haley Brink, MS, LMFTA":
           BrinkCount--;
           if (age == "YA"){
	  	  BrinkYACount--;
	  	  } else
	  	  {
	  	  BrinkADCount--;
	  }
	  ss.getRange("I"+i).setValue(BrinkCount);
	  ss.getRange("J"+i).setValue(BrinkADCount);
	  ss.getRange("K"+i).setValue(BrinkYACount);
	  if (BrinkCount > (BrinkLimit-1)) {ss.getRange("I"+i).setBackground("#CC0000"); ss.getRange("I"+i).setFontWeight("bold")};
          if (BrinkCount == (BrinkLimit-1)) {ss.getRange("I"+i).setBackground("#00CC00"); ss.getRange("I"+i).setFontWeight("bold")};
        break;

	case "Erik Jul":
	   JulCount--;
	   if (age == "YA"){
	   	  JulYACount--;
	   	  } else
	   	  {
	   	  JulADCount--;
	   }
	   ss.getRange("I"+i).setValue(JulCount);
	   ss.getRange("J"+i).setValue(JulADCount);
	   ss.getRange("K"+i).setValue(JulYACount);
	   if (JulCount > (JulLimit-1)) {ss.getRange("I"+i).setBackground("#CC0000"); ss.getRange("I"+i).setFontWeight("bold")};
	   if (JulCount == (JulLimit-1)) {ss.getRange("I"+i).setBackground("#00CC00"); ss.getRange("I"+i).setFontWeight("bold")};
          break;

	case "Mark Dunn":
	   DunnCount--;
	   if (age == "YA"){
	   	  DunnYACount--;
	   	  } else
	   	  {
	   	  DunnADCount--;
	   }
	   ss.getRange("I"+i).setValue(DunnCount);
	   ss.getRange("J"+i).setValue(DunnADCount);
	   ss.getRange("K"+i).setValue(DunnYACount);
	   if (DunnCount > (DunnLimit-1)) {ss.getRange("I"+i).setBackground("#CC0000"); ss.getRange("I"+i).setFontWeight("bold")};
	   if (DunnCount == (DunnLimit-1)) {ss.getRange("I"+i).setBackground("#00CC00"); ss.getRange("I"+i).setFontWeight("bold")};
	      break;

	case "Yvette Hinojosa, MA":
	   HinojosaCount--;
	   if (age == "YA"){
	   	  HinojosaYACount--;
	   	  } else
	   	  {
	   	  HinojosaADCount--;
	   }
	   ss.getRange("I"+i).setValue(HinojosaCount);
	   ss.getRange("J"+i).setValue(HinojosaADCount);
	   ss.getRange("K"+i).setValue(HinojosaYACount);
	   if (HinojosaCount > (HinojosaLimit-1)) {ss.getRange("I"+i).setBackground("#CC0000"); ss.getRange("I"+i).setFontWeight("bold")};
	   if (HinojosaCount == (HinojosaLimit-1)) {ss.getRange("I"+i).setBackground("#00CC00"); ss.getRange("I"+i).setFontWeight("bold")};
	      break;

		case "Annette Glass, LMFT":
		   GlassCount--;
		   if (age == "YA"){
		   	  GlassYACount--;
		   	  } else
		   	  {
		   	  GlassADCount--;
		   }
		   ss.getRange("I"+i).setValue(GlassCount);
		   ss.getRange("J"+i).setValue(GlassADCount);
		   ss.getRange("K"+i).setValue(GlassYACount);
		   if (GlassCount > (GlassLimit-1)) {ss.getRange("I"+i).setBackground("#CC0000"); ss.getRange("I"+i).setFontWeight("bold")};
		   if (GlassCount == (GlassLimit-1)) {ss.getRange("I"+i).setBackground("#00CC00"); ss.getRange("I"+i).setFontWeight("bold")};
		      break;

		case "Brian Konik, Ph.D":
		   KonikCount--;
		   if (age == "YA"){
		   	  KonikYACount--;
		   	  } else
		   	  {
		   	  KonikADCount--;
		   }
		   ss.getRange("I"+i).setValue(KonikCount);
		   ss.getRange("J"+i).setValue(KonikADCount);
		   ss.getRange("K"+i).setValue(KonikYACount);
		   if (KonikCount > (KonikLimit-1)) {ss.getRange("I"+i).setBackground("#CC0000"); ss.getRange("I"+i).setFontWeight("bold")};
		   if (KonikCount == (KonikLimit-1)) {ss.getRange("I"+i).setBackground("#00CC00"); ss.getRange("I"+i).setFontWeight("bold")};
		      break;

      }

      TotalCount--;
      ss.getRange("R"+i).setValue(TotalCount);

  }
    var haveage = ss.getRange("F"+i).getValue();
    if (haveage == "") {
         ss.getRange("F"+i).setValue(age);
      }
 }
  ss.deleteColumn(20);
  ss.deleteColumn(20);
  ss.setColumnWidth(18, 40);
  ss.getRange("A4:T4").setFontWeight("bold");
  ss.getRange("A4:T4").setBackground("#bebebe");

  var l = sheet.getLastRow();
  var m = sheet.getMaxRows();
  sheet.deleteRows(l+1,m-l);//delete all the empty rows after last filled row
  var range = sheet.getRange("N2:N").setBackground("#00ffff");
  var range = sheet.getRange("O2:O").setBackground("#ead1dc");
  var range = sheet.getRange("P2:P").setBackground("#6d9eeb");
  var range = sheet.getRange("Q2:Q").setBackground("#ff00ff");

}

//switch to CaseloadForecast
function switchSheet() {
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.setActiveSheet(ss.getSheetByName("CaseloadForecast"));
var m = sheet.getMaxRows();
sheet.insertRowsAfter(m, 10);
var l = sheet.getMaxColumns();
sheet.insertColumnsAfter(l, 10);
}

function caseload (){
//Get Therapist Names and Limits from TherapistLimits tab and write to Caseload Forecast tab
var limitSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TherapistLimits");
var limitLastRow = limitSheet.getLastRow();
var therapistNames = limitSheet.getRange("A2:A"+limitLastRow).getValues();
var therapistLimits = limitSheet.getRange("B2:B"+limitLastRow).getValues();
var caseloadSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CaseloadForecast");
caseloadSheet.clear();
caseloadSheet.getRange("A1").setValue("Therapist");
caseloadSheet.getRange("B1").setValue("Limit");
caseloadSheet.getRange("C1").setValue("Current");
caseloadSheet.getRange("D1").setValue("Peak");
caseloadSheet.getRange("E1").setValue("DOL");
caseloadSheet.getRange("F1").setValue("Next");
caseloadSheet.getRange("A2:A"+(limitLastRow)).setValues(therapistNames);
caseloadSheet.getRange("B2:B"+(limitLastRow)).setValues(therapistLimits);

//Get Initial Counts
var edodFile = SpreadsheetApp.openById("t6rpTfC3dx78qhurAS05KDw");
var edodLastRow = edodFile.getLastRow();
iMautzCount = 0, iMosesCount = 0, iAdamsCount = 0, iSullivanCount = 0, iKasenchakCount = 0, iWeldCount = 0, iJensenCount = 0, iBrinkCount = 0, iJulCount = 0, iDunnCount = 0, iHinojosaCount = 0, iGlassCount = 0, iKonikCount = 0;
for (i=2; i<edodLastRow+1; i++) {
    var therapRange = edodFile.getRange("G"+i).getValues();
    if (therapRange == "Toby Mautz, MSW, LCSW") {iMautzCount++;}
    if (therapRange == "Hilary Moses, MSW, LCSW") {iMosesCount++;}
    if (therapRange == "Jason R.K. Adams, Ph.D.") {iAdamsCount++;}
    if (therapRange == "Mike Sullivan, MA") {iSullivanCount++;}
    if (therapRange == "Kathryn Kasenchak, Psy.D.") {iKasenchakCount++;}
    if (therapRange == "Kelly Weld, M.A., MFT") {iWeldCount++;}
    if (therapRange == "Bridger Jensen, MA, CTRS") {iJensenCount++;}
    if (therapRange == "Haley Brink, MS, LMFTA") {iBrinkCount++;}
    if (therapRange == "Erik Jul") {iJulCount++;}
    if (therapRange == "Mark Dunn") {iDunnCount++;}
    if (therapRange == "Yvette Hinojosa, MA") {iHinojosaCount++;}
    if (therapRange == "Annette Glass, LMFT") {iGlassCount++;}
    if (therapRange == "Brian Konik, Ph.D") {iKonikCount++;}
  }

//Get Max/Peak Therapist Counts
var forecastSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("AdmissionsForecast");
var forecastLastRow = forecastSheet.getLastRow();
MautzCounts = [], MosesCounts = [], AdamsCounts = [], SullivanCounts = [], KasenchakCounts = [], WeldCounts = [], JensenCounts = [], BrinkCounts = [], JulCounts = [], DunnCounts = [], HinojosaCounts = [], GlassCounts = [], KonikCounts = [];
for (i=5; i<forecastLastRow+1; i++) {
    var therapistName = forecastSheet.getRange("H"+i).getValue();
    var therapistCount = forecastSheet.getRange("I"+i).getValue();
    var forecastDate = forecastSheet.getRange("A"+i).getValue();
    if (therapistName == "Toby Mautz, MSW, LCSW") {MautzCounts.push(therapistCount)}
    if (therapistName == "Hilary Moses, MSW, LCSW") {MosesCounts.push(therapistCount)}
    if (therapistName == "Jason R.K. Adams, Ph.D.") {AdamsCounts.push(therapistCount)}
    if (therapistName == "Mike Sullivan, MA") {SullivanCounts.push(therapistCount)}
    if (therapistName == "Kathryn Kasenchak, Psy.D.") {KasenchakCounts.push(therapistCount)}
    if (therapistName == "Kelly Weld, M.A., MFT") {WeldCounts.push(therapistCount)}
    if (therapistName == "Bridger Jensen, MA, CTRS") {JensenCounts.push(therapistCount)}
    if (therapistName == "Haley Brink, MS, LMFTA") {BrinkCounts.push(therapistCount)}
    if (therapistName == "Erik Jul") {JulCounts.push(therapistCount)}
    if (therapistName == "Mark Dunn") {DunnCounts.push(therapistCount)}
    if (therapistName == "Yvette Hinojosa, MA") {HinojosaCounts.push(therapistCount)}
    if (therapistName == "Annette Glass, LMFT") {GlassCounts.push(therapistCount)}
    if (therapistName == "Brian Konik, Ph.D") {KonikCounts.push(therapistCount)}
  }

//Get Next Available Date
// var caseloadSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CaseloadForecast");
var caseloadLastRow = caseloadSheet.getLastRow();
// var forecastSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("AdmissionsForecast");
var forecastLastRow = forecastSheet.getLastRow();
MautzDates = [], MosesDates = [], AdamsDates = [], SullivanDates = [], KasenchakDates = [], WeldDates = [], JensenDates = [], BrinkDates = [], JulDates = [], DunnDates = [], HinojosaDates = [], GlassDates = [], KonikDates = [];
  for (i=5; i<forecastLastRow+1; i++) {
    var therapistName = forecastSheet.getRange("H"+i).getValue();
    var countColor = forecastSheet.getRange("I"+i).getBackground();
    var forecastDate = forecastSheet.getRange("A"+i).getValue();
    if ((therapistName == "Toby Mautz, MSW, LCSW") && (countColor == "#00CC00")) {MautzDates.push(forecastDate)}
    if ((therapistName == "Hilary Moses, MSW, LCSW") && (countColor == "#00CC00")) {MosesDates.push(forecastDate)}
    if ((therapistName == "Jason R.K. Adams, Ph.D.") && (countColor == "#00CC00")) {AdamsDates.push(forecastDate)}
    if ((therapistName == "Mike Sullivan, MA") && (countColor == "#00CC00")) {SullivanDates.push(forecastDate)}
    if ((therapistName == "Kathryn Kasenchak, Psy.D.") && (countColor == "#00CC00")) {KasenchakDates.push(forecastDate)}
    if ((therapistName == "Kelly Weld, M.A., MFT") && (countColor == "#00CC00")) {WeldDates.push(forecastDate)}
    if ((therapistName == "Bridger Jensen, MA, CTRS") && (countColor == "#00CC00")) {JensenDates.push(forecastDate)}
    if ((therapistName == "Haley Brink, MS, LMFTA") && (countColor == "#00CC00")) {BrinkDates.push(forecastDate)}
    if ((therapistName == "Erik Jul") && (countColor == "#00CC00")) {JulDates.push(forecastDate)}
    if ((therapistName == "Mark Dunn") && (countColor == "#00CC00")) {DunnDates.push(forecastDate)}
    if ((therapistName == "Yvette Hinojosa, MA") && (countColor == "#00CC00")) {HinojosaDates.push(forecastDate)}
    if ((therapistName == "Annette Glass, LMFT") && (countColor == "#00CC00")) {GlassDates.push(forecastDate)}
    if ((therapistName == "Brian Konik, Ph.D") && (countColor == "#00CC00")) {KonikDates.push(forecastDate)}
  }

//Write Initial Counts, Max/Peak Therapist Count, and Next Available
  for (i=2; i<caseloadLastRow+1; i++){
    var therapistName = caseloadSheet.getRange("A"+i).getValue();
    if (therapistName == "Toby Mautz, MSW, LCSW") {caseloadSheet.getRange("C"+i).setValue(iMautzCount);
                                                   caseloadSheet.getRange("D"+i).setValue(Math.max.apply(Math, MautzCounts));
                                                   caseloadSheet.getRange("F"+i).setValue(Math.max.apply(null, MautzDates))}
    if (therapistName == "Hilary Moses, MSW, LCSW") {caseloadSheet.getRange("C"+i).setValue(iMosesCount);
                                                   caseloadSheet.getRange("D"+i).setValue(Math.max.apply(Math, MosesCounts));
                                                   caseloadSheet.getRange("F"+i).setValue(Math.max.apply(null, MosesDates))}
    if (therapistName == "Jason R.K. Adams, Ph.D.") {caseloadSheet.getRange("C"+i).setValue(iAdamsCount);
                                                   caseloadSheet.getRange("D"+i).setValue(Math.max.apply(Math, AdamsCounts));
                                                   caseloadSheet.getRange("F"+i).setValue(Math.max.apply(Math, AdamsDates))}
    if (therapistName == "Mike Sullivan, MA") {caseloadSheet.getRange("C"+i).setValue(iSullivanCount);
                                                   caseloadSheet.getRange("D"+i).setValue(Math.max.apply(Math, SullivanCounts));
                                                   caseloadSheet.getRange("F"+i).setValue(Math.max.apply(Math, SullivanDates))}
    if (therapistName == "Kathryn Kasenchak, Psy.D.") {caseloadSheet.getRange("C"+i).setValue(iKasenchakCount);
                                                   caseloadSheet.getRange("D"+i).setValue(Math.max.apply(Math, KasenchakCounts));
                                                   caseloadSheet.getRange("F"+i).setValue(Math.max.apply(Math, KasenchakDates))}
    if (therapistName == "Kelly Weld, M.A., MFTi") {caseloadSheet.getRange("C"+i).setValue(iWeldCount);
                                                   caseloadSheet.getRange("D"+i).setValue(Math.max.apply(Math, WeldCounts));
                                                   caseloadSheet.getRange("F"+i).setValue(Math.max.apply(Math, WeldDates))}
    if (therapistName == "Bridger Jensen, MA, CTRS") {caseloadSheet.getRange("C"+i).setValue(iJensenCount);
                                                   caseloadSheet.getRange("D"+i).setValue(Math.max.apply(Math, JensenCounts));
                                                   caseloadSheet.getRange("F"+i).setValue(Math.max.apply(Math, JensenDates))}
    if (therapistName == "Haley Brink, MS, LMFTA") {caseloadSheet.getRange("C"+i).setValue(iBrinkCount);
                                                   caseloadSheet.getRange("D"+i).setValue(Math.max.apply(Math, BrinkCounts));
                                                   caseloadSheet.getRange("F"+i).setValue(Math.max.apply(Math, BrinkDates))}
    if (therapistName == "Erik Jul") {caseloadSheet.getRange("C"+i).setValue(iJulCount);
                                                   caseloadSheet.getRange("D"+i).setValue(Math.max.apply(Math, JulCounts));
                                                   caseloadSheet.getRange("F"+i).setValue(Math.max.apply(Math, JulDates))}
    if (therapistName == "Mark Dunn") {caseloadSheet.getRange("C"+i).setValue(iDunnCount);
                                                   caseloadSheet.getRange("D"+i).setValue(Math.max.apply(Math, DunnCounts));
                                                   caseloadSheet.getRange("F"+i).setValue(Math.max.apply(Math, DunnDates))}
    if (therapistName == "Yvette Hinojosa, MA") {caseloadSheet.getRange("C"+i).setValue(iHinojosaCount);
                                                   caseloadSheet.getRange("D"+i).setValue(Math.max.apply(Math, HinojosaCounts));
                                                   caseloadSheet.getRange("F"+i).setValue(Math.max.apply(Math, HinojosaDates))}
    if (therapistName == "Annette Glass, LMFT") {caseloadSheet.getRange("C"+i).setValue(iGlassCount);
                                                   caseloadSheet.getRange("D"+i).setValue(Math.max.apply(Math, GlassCounts));
                                                   caseloadSheet.getRange("F"+i).setValue(Math.max.apply(Math, GlassDates))}
    if (therapistName == "Brian Konik, Ph.D") {caseloadSheet.getRange("C"+i).setValue(iKonikCount);
                                                   caseloadSheet.getRange("D"+i).setValue(Math.max.apply(Math, KonikCounts));
                                                   caseloadSheet.getRange("F"+i).setValue(Math.max.apply(Math, KonikDates))}
  }

  for (i=2; i<caseloadLastRow+1; i++) {
    var numberDaysSinceNixon = caseloadSheet.getRange("F"+i).getValue();
    var newDate = new Date(numberDaysSinceNixon);
    var todayFormat = Utilities.formatDate(newDate,'GMT-8',"MM/dd/yyyy");
    var current = caseloadSheet.getRange("C"+i).getValue();
    var peak = caseloadSheet.getRange("D"+i).getValue();
    var therapistLimit = caseloadSheet.getRange("B"+i).getValue():
    caseloadSheet.getRange("F"+i).setValue(todayFormat);
    if (numberDaysSinceNixon == "-Infinity") {caseloadSheet.getRange("F"+i).setValue("TBD");}
    if (current > peak) {caseloadSheet.getRange("D"+i).setValue(current);}
    if (therapistLimit > peak) {caseloadSheet.getRange("F"+i).setValue("Now");}
  }



}

//make the CaseloadForecast pretty and delete empty space / make YA "Therapist" rows blue based on name in Therapist colum.
function prettySheet() {
 var caseloadSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CaseloadForecast");

 caseloadSheet.setRowHeight(1,40);
 caseloadSheet.setColumnWidth(1,300);
 caseloadSheet.getRange("A1:F1").setFontSize(12).setFontWeight("bold").setBackgroundColor("#d9d9d9");


var lc = caseloadSheet.getLastColumn();
var mc = caseloadSheet.getMaxColumns();
caseloadSheet.deleteColumns(lc+1,mc-lc);//delete all the empty columns after last filled column
var l = caseloadSheet.getLastRow();
var m = caseloadSheet.getMaxRows();
caseloadSheet.deleteRows(l+1,m-l);//delete all the empty rows after last filled row

var range1 = caseloadSheet.getRange("E1");
caseloadSheet.hideColumn(range1); //hide unused DOL column


  var range = caseloadSheet.getDataRange();
  var statusColumnOffset = getStatusColumnOffset();

  for (var i = range.getRow(); i < range.getLastRow(); i++) {
    rowRange = range.offset(i, 0, 1);
    status = rowRange.offset(0, statusColumnOffset).getValue();
    if (status == 'Mark Dunn') {
      rowRange.setBackground("#d0dbed");
    } else if (status == 'Kelly Weld, M.A., MFTi') {
      rowRange.setBackground("#d0dbed");
    } else if (status == 'Annette Glass, LMFT') {
      rowRange.setBackground("#d0dbed");
    } else if (status == 'Erik Jul') {
      rowRange.setBackground("#d0dbed");
    }
  }
}

//Returns the offset value of the column titled "Therapist" used by prettySheet function
function getStatusColumnOffset() {
  var caseloadSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CaseloadForecast");
  lastColumn = caseloadSheet.getLastColumn();
  var range = caseloadSheet.getRange(1,1,1,lastColumn);

  for (var i = 0; i < range.getLastColumn(); i++) {
    if (range.offset(0, i, 1, 1).getValue() == "Therapist") {
      return i;
    }
  }
}

//Returns view to sheet AdmissionsForecast
function goHome () {
var ss = SpreadsheetApp.getActiveSpreadsheet();
ss.setActiveSheet(ss.getSheetByName("AdmissionsForecast"));
}
