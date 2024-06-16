/******************************
Create Reports for PPR
Version 1.0

Daniel Terol
******************************/


//////////////////////////////////////////////////////////////////////////
// Gets the data from the CSV and parses it into a multidimensional array
var thisFile = new File($.fileName);
var basePath = thisFile.path;

// Opens the CSV file and puts the contents on a string
var csvFile = File(basePath + '/PPR_data.csv');

csvFile.encoding = 'UTF8'; // set to 'UTF8' or 'UTF-8'

csvFile.open("r");

var csvString = csvFile.read();

csvFile.close();

// This function takes the csvString and converts it into an array that we can use
function parseCSV(str) {
    var arr = [];
    var quote = false;  // 'true' means we're inside a quoted field

    // Iterate over each character, keep track of current row and column (of the returned array)
    for (var row = 0, col = 0, c = 0; c < str.length; c++) {
        var cc = str[c], nc = str[c+1];        // Current character, next character
        arr[row] = arr[row] || [];             // Create a new row if necessary
        arr[row][col] = arr[row][col] || '';   // Create a new column (start with empty string) if necessary

        // If the current character is a quotation mark, and we're inside a
        // quoted field, and the next character is also a quotation mark,
        // add a quotation mark to the current column and skip the next character
        if (cc == '"' && quote && nc == '"') { arr[row][col] += cc; ++c; continue; }

        // If it's just one quotation mark, begin/end quoted field
        if (cc == '"') { quote = !quote; continue; }

        // If it's a comma and we're not in a quoted field, move on to the next column
        if (cc == ',' && !quote) { ++col; continue; }

        // If it's a newline (CRLF) and we're not in a quoted field, skip the next character
        // and move on to the next row and move to column 0 of that new row
        if (cc == '\r' && nc == '\n' && !quote) { ++row; col = 0; ++c; continue; }

        // If it's a newline (LF or CR) and we're not in a quoted field,
        // move on to the next row and move to column 0 of that new row
        if (cc == '\n' && !quote) { ++row; col = 0; continue; }
        if (cc == '\r' && !quote) { ++row; col = 0; continue; }

        // Otherwise, append the current character to the current column
        arr[row][col] += cc;
    }
    return arr;
}

// This is the multidimensional array that we'll use to make all the updates on the document
var myData = parseCSV(csvString);

// Removes the % percentage signs if any
for (y=2; y<myData[0].length; y++){
    myData[1][y] = myData[1][y].replace('%', '');
    myData[2][y] = myData[2][y].replace('%', '');
    myData[3][y] = myData[3][y].replace('%', '');
    myData[4][y] = myData[4][y].replace('%', '');
}


///////////////////////
// Here starts the loop
for (y=2; y<myData[0].length; y++){
    
    
    var fileRef = File (basePath + '/PPR_Template.ait');
    if (fileRef != null & y>0) {
        var openOptions = new OpenOptions();
        openOptions.updateLegacyText = true;
        var docRef = open(fileRef, DocumentColorSpace.RGB, openOptions);
    }
    
    var myDoc = app.activeDocument;
    var editsLayer = myDoc.layers['Edits'];
    var readmeLayer = myDoc.layers['READ ME'];
    readmeLayer.remove(); //Removes the READ ME layer that contains information for designers


    var legendCompany = editsLayer.textFrames['companyName'];
    myData[0][y] = myData[0][y].toUpperCase();
    legendCompany.contents = myData[0][y];
    var legendCompetitor = editsLayer.groupItems['legend2'];
    
    
    var bar_1_1 = editsLayer.pathItems['bar_1_1'];
    var bar_1_2 = editsLayer.pathItems['bar_1_2'];
    var bar_2_1 = editsLayer.pathItems['bar_2_1'];
    var bar_2_2 = editsLayer.pathItems['bar_2_2'];
    var bar_3_1 = editsLayer.pathItems['bar_3_1'];
    var bar_3_2 = editsLayer.pathItems['bar_3_2'];
    var bar_4_1 = editsLayer.pathItems['bar_4_1'];
    var bar_4_2 = editsLayer.pathItems['bar_4_2'];
    
    var label_1_1 = editsLayer.textFrames['label_1_1'];
    var label_1_2 = editsLayer.textFrames['label_1_2'];
    var label_2_1 = editsLayer.textFrames['label_2_1'];
    var label_2_2 = editsLayer.textFrames['label_2_2'];
    var label_3_1 = editsLayer.textFrames['label_3_1'];
    var label_3_2 = editsLayer.textFrames['label_3_2'];
    var label_4_1 = editsLayer.textFrames['label_4_1'];
    var label_4_2 = editsLayer.textFrames['label_4_2'];
    
    var graph_1 = editsLayer.groupItems['graph_1'];
    var graph_2 = editsLayer.groupItems['graph_2'];
    var graph_3 = editsLayer.groupItems['graph_3'];
    var graph_4 = editsLayer.groupItems['graph_4'];

    var topRanges = [0, 0, 0, 0];

    
    function updateLabels(a,b,x,al,bl){ // a = current data value, b = benchmark data, x = the graph to update, al = bar1 label, bl = benchmark label
        var aK = a / 1000;
        aK = aK.toFixed(1);
        var aM = a / 1000000;
        aM = aM.toFixed(2);
        var bK = b / 1000;
        bK = bK.toFixed(1);
        var bM = b / 1000000;
        bM = bM.toFixed(2);
        var x1 = x.textFrames[1];
        var x2 = x.textFrames[2];
        var no;
        var hi;
        var ho = b - a;
        var per = '';

        if(ho < 0){
            hi = a;
        } else {
            hi = b;
        }

        if (x == '[GroupItem graph_1]') {
            no = 0;
        } else if (x == '[GroupItem graph_2]'){
            no = 1;
        } else if (x == '[GroupItem graph_3]'){
            no = 2;
        } else if (x == '[GroupItem graph_4]'){
            no = 3;
        }
                 
        if ((myData[1][0] === '%' && no === 0) || 
        (myData[2][0] === '%' && no === 1) || 
        (myData[3][0] === '%' && no === 2) || 
        (myData[4][0] === '%' && no === 3)
        ){
            per = '%';
        } 
        
        if (hi <= 1000){ // Deals with the under 1K data
            al.contents = a + per;
            bl.contents = b + per;
            if (hi <= 100){
                x1.contents = 50;
                x2.contents = 100;
                topRanges[no] = 100;
            } else if (hi > 100 && hi <= 200){
                x1.contents = 100;
                x2.contents = 200;
                topRanges[no] = 200;
            } else if (hi > 200 && hi <= 500){
                x1.contents = 250;
                x2.contents = 500;
                topRanges[no] = 500;
            } else if (hi > 500 && hi <= 800){
                x1.contents = 400;
                x2.contents = 800;
                topRanges[no] = 800;
            } else if (hi > 800 && hi <= 1000){
                x1.contents = 800;
                x2.contents = 1000;
                topRanges[no] = 1000;
            }
        } else if (hi > 1000 && hi < 999950){ // Deals with the 1K data
            al.contents = aK + 'K' + per;
            bl.contents = bK + 'K' + per;
            
            if (hi > 1000 && hi <= 1600){
                x1.contents = '0.8K';
                x2.contents = '1.6K';
                topRanges[no] = 1600;
            } else if (hi > 1600 && hi <= 2000){
                x1.contents = '1K';
                x2.contents = '2K';
                topRanges[no] = 2000;
            } else if (hi > 2000 && hi <= 4000){
                x1.contents = '2K';
                x2.contents = '4K';
                topRanges[no] = 4000;
            } else if (hi > 4000 && hi <= 6000){
                x1.contents = '3K';
                x2.contents = '6K';
                topRanges[no] = 6000;
            } else if (hi > 6000 && hi <= 10000){
                x1.contents = '5K';
                x2.contents = '10K';
                topRanges[no] = 10000;
            } else if (hi > 10000 && hi <= 20000){
                x1.contents = '10K';
                x2.contents = '20K';
                topRanges[no] = 20000;
            } else if (hi > 20000 && hi <= 50000){
                x1.contents = '25K';
                x2.contents = '50K';
                topRanges[no] = 50000;
            } else if (hi > 50000 && hi <= 100000){
                x1.contents = '50K';
                x2.contents = '100K';
                topRanges[no] = 100000;
            } else if (hi > 100000 && hi <= 200000){
                x1.contents = '100K';
                x2.contents = '200K';
                topRanges[no] = 200000;
            } else if (hi > 200000 && hi <= 500000){
                x1.contents = '250K';
                x2.contents = '500K';
                topRanges[no] = 500000;
            } else if (hi > 500000 && hi < 999950){
                x1.contents = '500K';
                x2.contents = '1M';
                topRanges[no] = 1000000;
            }
        } else { // Deals with the 1M data
            al.contents = aM + 'M' + per;
            bl.contents = bM + 'M' + per;
            if (hi >= 999950 && hi <= 2000000){
                x1.contents = '1M';
                x2.contents = '2M';
                topRanges[no] = 2000000;
            } else if (hi > 2000000 && hi <= 5000000){
                x1.contents = '2M';
                x2.contents = '5M';
                topRanges[no] = 5000000;
            } else if (hi > 5000000 && hi <= 10000000){
                x1.contents = '5M';
                x2.contents = '10M';
                topRanges[no] = 10000000;
            } else if (hi > 10000000 && hi <= 20000000){
                x1.contents = '10M';
                x2.contents = '20M';
                topRanges[no] = 20000000;
            } else if (hi > 20000000 && hi <= 50000000){
                x1.contents = '25M';
                x2.contents = '50M';
                topRanges[no] = 50000000;
            } else if (hi > 50000000 && hi <= 100000000){
                x1.contents = '50M';
                x2.contents = '100M';
                topRanges[no] = 100000000;
            }
        }
        
    }
    updateLabels(myData[1][y], myData[1][1], graph_1, label_1_1, label_1_2);
    updateLabels(myData[2][y], myData[2][1], graph_2, label_2_1, label_2_2);
    updateLabels(myData[3][y], myData[3][1], graph_3, label_3_1, label_3_2);
    updateLabels(myData[4][y], myData[4][1], graph_4, label_4_1, label_4_2);
        

    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    // Graph bar sizes. The numbers are the full size of the chart plot, height for the vertical charts, width for the horizontal
    var height_1_1 = 110.5 / topRanges[0] * myData[1][y];
    height_1_1 = height_1_1.toFixed(2);
    var height_1_2 = 110.5 / topRanges[0] * myData[1][1];
    height_1_2 = height_1_2.toFixed(2);
    var height_2_1 = 110.5 / topRanges[1] * myData[2][y];
    height_2_1 = height_2_1.toFixed(2);
    var height_2_2 = 110.5 / topRanges[1] * myData[2][1];
    height_2_2 = height_2_2.toFixed(2);
    var height_3_1 = 110.5 / topRanges[2] * myData[3][y];
    height_3_1 = height_3_1.toFixed(2);
    var height_3_2 = 110.5 / topRanges[2] * myData[3][1];
    height_3_2 = height_3_2.toFixed(2);
    var width_4_1 = 261.5 / topRanges[3] * myData[4][y];
    width_4_1 = width_4_1.toFixed(2);
    var width_4_2 = 261.5 / topRanges[3] * myData[4][1];
    width_4_2 = width_4_2.toFixed(2);
    
    
    /////////////////////////////////////////////////////////////////////////////
    // This is the function to adjust the size of the different bars on the graphs
    function adjustVertBar (a, h){ // a is the bar, h is the new height
        var barHeight = a.height;
        var diff = h - barHeight;
        a.height = h;
        a.top += diff;
    }
    
    adjustVertBar (bar_1_1, height_1_1);
    adjustVertBar (bar_1_2, height_1_2);
    adjustVertBar (bar_2_1, height_2_1);
    adjustVertBar (bar_2_2, height_2_2);
    adjustVertBar (bar_3_1, height_3_1);
    adjustVertBar (bar_3_2, height_3_2);
    
    bar_4_1.width = width_4_1;
    bar_4_2.width = width_4_2;
    

    ////////////////////////////////////////////////////////////////////////////
    // Adjust the positions of the items next to objects which size gets updated
    function adjustPosition (a, x){ // a is the object to move, x the object of reference, f the object of reference old size
        if (a==legendCompetitor){
            a.left = x.left + x.width + 20;
        } else if (a==label_4_1 || a==label_4_2){
            a.left = x.left + x.width + 4;
        } else {
            a.top = x.top + 10;
        }
        
    }
    
    adjustPosition (legendCompetitor, legendCompany);
    adjustPosition (label_4_1, bar_4_1);
    adjustPosition (label_4_2, bar_4_2);
    adjustPosition (label_1_1, bar_1_1);
    adjustPosition (label_1_2, bar_1_2);
    adjustPosition (label_2_1, bar_2_1);
    adjustPosition (label_2_2, bar_2_2);
    adjustPosition (label_3_1, bar_3_1);
    adjustPosition (label_3_2, bar_3_2);
    

    //////////////////////////////////////
    // Places and resizes the company logo
    var newLogo = File (basePath + '/Logos/' + myData[5][y]);
    var companyLogo = editsLayer.groupItems.createFromFile(newLogo);
    var logoWidth = companyLogo.width;
    var logoHeith = companyLogo.height;
    
    var ratio = logoWidth / logoHeith;
    if (ratio > 0 & ratio <= 2){
        companyLogo.width = 35  ;
    } else if (ratio > 2 & ratio <= 3.5){
        companyLogo.width = 45;
    } else if (ratio > 3.5 & ratio <= 4){
        companyLogo.width = 50;
    } else if (ratio > 4 & ratio <= 4.5){
        companyLogo.width = 50;
    } else if (ratio > 4.5 & ratio <= 5.5){
        companyLogo.width = 60;
    } else if (ratio > 5 & ratio <= 10){
        companyLogo.width = 95;
    } else {
        companyLogo.width = 110;
    }
    companyLogo.height = companyLogo.width / ratio;

    var LogoRef = editsLayer.textFrames['LogoRef'];
    
    //companyLogo.top = 1126.5 + companyLogo.height / 2;

    companyLogo.top = LogoRef.top - LogoRef.height / 2 + companyLogo.height / 2;

    companyLogo.left = LogoRef.left + LogoRef.width + 11;


    ////////////////////////////////
    // Saves and closes the document
    var newPDF = basePath + '/PDF/' + myData[0][y];
    
    function saveFileToPDF(dest) {
        var doc = app.activeDocument;
        
        if (app.documents.length > 0) {
            var saveName = new File(dest);
            saveOpts = new PDFSaveOptions();
            
            saveOpts.compatibility = PDFCompatibility.ACROBAT5;
            saveOpts.generateThumbnails = true;
            saveOpts.preserveEditability = true;
            
            doc.saveAs(saveName, saveOpts);
        }
    }
    saveFileToPDF(newPDF);
    
    if ( app.documents.length > 0 ) {
        var aiDocument = myDoc;
        aiDocument.close( SaveOptions.DONOTSAVECHANGES );
        aiDocument = null;
    } 
}