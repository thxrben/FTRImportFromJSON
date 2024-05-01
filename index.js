const fs = require('fs');
var HTMLParser = require('node-html-parser');
const excel = require('excel4node');

const searchForType = "FTR"; // This is the Type-Filter. Can be TSR, FTR, TRL, PSA, etc.
var globalOutIterator = 0; // Counts all DCP entries (Number of total DCPs found)

const workbook = new excel.Workbook();
//Create the new worksheets
const worksheetNAS = workbook.addWorksheet("NAS");
const worksheetS1 = workbook.addWorksheet("S1");
const worksheetS2 = workbook.addWorksheet("S2");
const worksheetS3 = workbook.addWorksheet("S3");
const worksheetS4 = workbook.addWorksheet("S4");
const worksheetS5 = workbook.addWorksheet("S5");


const headerCellDesign = workbook.createStyle({ // this is the design for the headers ("Nummer, Filmname, Soll gelöscht werden? etc.")
    alignment: {
        wrapText: false,
        horizontal: 'center'
    },
    border: {
        left: {
            style: 'medium',
            color: '000000'
        }, 
        right: {
            style: 'medium',
            color: '000000'
        }, 
        top: {
            style: 'medium',
            color: '000000'
        }, 
        bottom: {
            style: 'medium',
            color: '000000'
        }
    }
});

const featureCellDesign = workbook.createStyle({
    border: {
        top: {
            style: 'thin',
            color: '000000'
        }
    },
})

const leftTextDesign = workbook.createStyle({
    alignment: {
        horizontal: 'left'
    }
})
const centerTextDesign = workbook.createStyle({
    alignment: {
        horizontal: 'center'
    }
})

const generalCellDesign = workbook.createStyle({
      border: {
        left: {
            style: 'thin',
            color: '000000'
        }, 
        right: {
            style: 'thin',
            color: '000000'
        }, 
        top: {
            style: 'thin',
            color: '000000'
        }, 
        bottom: {
            style: 'thin',
            color: '000000'
        }
    }
});

var counter = { // Counter for different features in a given storage space. (Number of different Titles per Storage)
    "NAS": {
        count: 1
    },
    "S1": {
        count: 1
    },
    "S2": {
        count: 1
    },
    "S3": {
        count: 1
    },
    "S4": {
        count: 1
    },
    "S5": {
        count: 1
    },
}

var cellArray = { // Temporary vars for the current absolute writing positions. An Offset will be used. This defines the starting position of the table.
    "NAS": {
        x: 1,
        y: 8
    },
    "S1": {
        x: 1,
        y: 8
    },
    "S2": {
        x: 1,
        y: 8
    },
    "S3": {
        x: 1,
        y: 8
    },
    "S4": {
        x: 1,
        y: 8
    },
    "S5": {
        x: 1,
        y: 8
    },
    
}

//Headings:
const headingColumnsNAS = [
    "Nummer",
    "Filmname",
    "CTID",
    "Soll gelöscht werden?",
    "Zusatz:",
];

const headingColumnsSAAL = [
    "Nummer",
    "Filmname",
    "CTID",
    "Existiert auf NAS?",
    "Soll gelöscht werden?",
    "Zusatz:"
];

//Create the headings in the worksheets...
var headingColumnIndex = 1;

headingColumnsNAS.forEach(heading => {
    worksheetNAS.cell(7, headingColumnIndex++)
        .string(heading).style(headerCellDesign);
});

headingColumnIndex = 1;


headingColumnsSAAL.forEach(heading => {
    worksheetS1.cell(7, headingColumnIndex)
    .string(heading).style(headerCellDesign);
    worksheetS2.cell(7, headingColumnIndex)
        .string(heading).style(headerCellDesign);
    worksheetS3.cell(7, headingColumnIndex)
        .string(heading).style(headerCellDesign);
    worksheetS4.cell(7, headingColumnIndex)
        .string(heading).style(headerCellDesign);
    worksheetS5.cell(7, headingColumnIndex)
        .string(heading).style(headerCellDesign);

    headingColumnIndex++;
});


//Format each worksheet according to these statements...
worksheetNAS.cell(1,1).string("Filmliste");
worksheetS1.cell(1,1).string("Filmliste");
worksheetS2.cell(1,1).string("Filmliste");
worksheetS3.cell(1,1).string("Filmliste");
worksheetS4.cell(1,1).string("Filmliste");
worksheetS5.cell(1,1).string("Filmliste");

worksheetNAS.cell(3,1).string("Stand vom:");
worksheetS1.cell(3,1).string("Stand vom:");
worksheetS2.cell(3,1).string("Stand vom:");
worksheetS3.cell(3,1).string("Stand vom:");
worksheetS4.cell(3,1).string("Stand vom:");
worksheetS5.cell(3,1).string("Stand vom:");

worksheetNAS.cell(3,2).date(new Date());
worksheetS1.cell(3,2).date(new Date());
worksheetS2.cell(3,2).date(new Date());
worksheetS3.cell(3,2).date(new Date());
worksheetS4.cell(3,2).date(new Date());
worksheetS5.cell(3,2).date(new Date());



worksheetNAS.cell(5,1).string("Speicherort:");
worksheetS1.cell(5,1).string("Speicherort:");
worksheetS2.cell(5,1).string("Speicherort:");
worksheetS3.cell(5,1).string("Speicherort:");
worksheetS4.cell(5,1).string("Speicherort:");
worksheetS5.cell(5,1).string("Speicherort:");

worksheetNAS.cell(5,2).string("NAS");
worksheetS1.cell(5,2).string("Filmserver 1");
worksheetS2.cell(5,2).string("Filmserver 2");
worksheetS3.cell(5,2).string("Filmserver 3");
worksheetS4.cell(5,2).string("Filmserver 4");
worksheetS5.cell(5,2).string("Filmserver 5");



worksheetNAS.cell(1, 8);
worksheetS1.cell(1,8);
worksheetS2.cell(1,8);
worksheetS3.cell(1,8);
worksheetS4.cell(1,8);
worksheetS5.cell(1,8);

worksheetNAS.column(2).setWidth(15);
worksheetS1.column(2).setWidth(15);
worksheetS2.column(2).setWidth(15);
worksheetS3.column(2).setWidth(15);
worksheetS4.column(2).setWidth(15);
worksheetS5.column(2).setWidth(15);


worksheetNAS.column(3).setWidth(19);
worksheetS1.column(3).setWidth(19);
worksheetS2.column(3).setWidth(19);
worksheetS3.column(3).setWidth(19);
worksheetS4.column(3).setWidth(19);
worksheetS5.column(3).setWidth(19);


worksheetNAS.column(4).setWidth(18);
worksheetS1.column(4).setWidth(18);
worksheetS2.column(4).setWidth(18);
worksheetS3.column(4).setWidth(18);
worksheetS4.column(4).setWidth(18);
worksheetS5.column(4).setWidth(18);



// Here we are starting to retrieve, sort and filter the features from our json file.
// This here starts the actual work.
// The jsons file name in content/ should be "storageExport.txt"

fs.readdir('content/', function(err, filenames) {
    if (err) {
      onError(err);
      return;
    }

        var rawdata = fs.readFileSync('content/storageExport.txt', 'utf-8', (err,data) => {
            console.log("Error: "+err);
            return;
        });

        // Parse the data and create an array:
        const jsonData = JSON.parse(rawdata);
        const data = Object.entries(jsonData);


        //console.log(data);
        
        data.forEach(feature => { // For each feature found:
            const storages = Object.entries(feature[1].storages); // Get the storage entry (on which storage is this feature currently saved?)
            console.log("IS NULL: "+(storages == undefined));
            storages.forEach(storage => { // For each storage found:
                if(storage[1].size > 0) { // If any DCP is stored there (This could be redundant? Unsure.)
                    const ctids = storage[1].ctids; // get the ctids
                    //Now we are going to write the feature to a worksheet. For this, we are going to pass the feature and everything here. (storage[0] and storage[1] could be replaced by the storage array, and the access could happen inside the function.)
                    // (The function call can be optimized by removing necessary variables that are redundant here. Might clean this up later.)
                    writeFeatureToWorksheet(feature, ctids, storage[0].toLocaleUpperCase(), storage[1]); // Writes a feature including all of its CTIDs into the spreadsheets.
                }
            });
        });
        // After all the work, the Worksheet shall be written out
        workbook.write('Export.xlsx');
        // And here, we are done.
        console.log("IM DONE!")
});

function writeFeatureToWorksheet(feature, ctids, storageName, storage) {
    var storageCell = undefined;
    var worksheetToUse = undefined;
    var counterToUse = undefined;

    //Retrieve the appropriate arrays for each storage. 
    //Havent found a way yet to replace this switch statement.
    switch(storageName) {
        case "NAS":
            storageCell = cellArray.NAS;
            worksheetToUse = worksheetNAS;
            counterToUse = counter.NAS;
            break;
        case "S1":
            storageCell = cellArray.S1;
            worksheetToUse = worksheetS1;
            counterToUse = counter.S1;
            break;
        case "S2":
            storageCell = cellArray.S2;
            worksheetToUse = worksheetS2;
            counterToUse = counter.S2;
            break;
        case "S3":
            storageCell = cellArray.S3;
            worksheetToUse = worksheetS3;
            counterToUse = counter.S3;
            break;
        case "S4":
            storageCell = cellArray.S4;
            worksheetToUse = worksheetS4;
            counterToUse = counter.S4;
            break;
        case "S5":
            storageCell = cellArray.S5;
            worksheetToUse = worksheetS5;
            counterToUse = counter.S5;
            break;
        default:
            console.err("Storage not found: \""+storageName+"\"");
            return;
    }

    //console.log(feature[1]);
    //console.log(storageCell);
    //
    const xBegin = storageCell.x; // Before playing with the vars, lets save them for later. (See the for-loop in line 380)
    const yBegin = storageCell.y;

    worksheetToUse.cell(storageCell.y, storageCell.x).number(counterToUse.count++); // Write the current title count
    worksheetToUse.cell(storageCell.y, storageCell.x+1).string(feature[0]).style(generalCellDesign); // Write the title of the feature
    
    //Größe der DCPs (alle zusammen)
    //worksheetToUse.cell(storageCell.y, storageCell.x+3).string((Math.round(storage.size * 100) / 100).toFixed(2)+" GB").style(generalCellDesign); //Größe

    const isCurrentlyPlaying = feature[1].isCurrentlyRunning; // Is currently running as a feature?s
    
    if(storageName != "NAS") { // SAAL-WORKSHEET:
        const isOnNas = feature[1].storages.nas.ctids.length > 0;

        const design = workbook.createStyle({
            alignment: {
                horizontal: 'center'
            },
            fill: {
                type: 'pattern', // the only one implemented so far.
                patternType: 'solid', // most common.
                fgColor: (!isOnNas ? 'ff7171' : '8affd7'), // you can add two extra characters to serve as alpha, i.e. '2172d7aa'.
                // bgColor: 'ffffff' // bgColor only applies on patternTypes other than solid.
            }
        });

        // Ist auf NAS:
        worksheetToUse.cell(storageCell.y, storageCell.x+3).string((isOnNas ? "JA" : "NEIN")).style(generalCellDesign).style(design); //Ist auf NAS?
        // Spielt noch?
        worksheetToUse.cell(storageCell.y, storageCell.x+4).string((isCurrentlyPlaying ? "SPIELT NOCH!" : "")).style(generalCellDesign);
        
    } else { // NAS-WORKSHEET:
        worksheetToUse.cell(storageCell.y, storageCell.x+3).string((isCurrentlyPlaying ? "SPIELT NOCH!" : "")).style(generalCellDesign);
    }

    console.log(feature[1]);
    var ctid = ctids[0].split("<b>ContentTitleId: </b>");
    //Style the cells:
    worksheetToUse.cell(storageCell.y, storageCell.x+2).style(generalCellDesign);
    worksheetToUse.cell(storageCell.y, storageCell.x+3).style(generalCellDesign);
    worksheetToUse.cell(storageCell.y++, storageCell.x+2).string(ctid[1]).style(generalCellDesign).style(leftTextDesign); //Finally, put the CTID into the cell. Limited to 20 chars (0-20)

    
/*
    ctids.forEach(ctid => {
        worksheetToUse.cell(storageCell.y, storageCell.x+2).style(generalCellDesign);
        worksheetToUse.cell(storageCell.y, storageCell.x+3).style(generalCellDesign);
        worksheetToUse.cell(storageCell.y++, storageCell.x+2).string(ctid).style(generalCellDesign).style(leftTextDesign); 
    });
*/

    for(var i = 1; i < xBegin+ (storageName == "NAS" ? 5 : 6); i++) {
        worksheetToUse.cell(yBegin, i).style(featureCellDesign);
    }
}


// #######################
// ABANDONED CODE SNIPPETS
// #######################

        /*
        if(feature[1].isKDMfree) {

            const kdmCellDesign = workbook.createStyle({
                alignment: {
                    horizontal: 'center'
                },
                fill: {
                    type: 'pattern', // the only one implemented so far.
                    patternType: 'solid', // most common.
                    fgColor: '8affd7', // fg (foreground) is the "foreground" of the "background.", e.g. the first layer of the background color of a cell. (2nd layer is used for e.g. gradients etc.)
                    // bgColor: 'ffffff' // bgColor only applies on patternTypes other than solid.
                }
            });
        
            //worksheetToUse.cell(storageCell.y, storageCell.x+2).string("JA").style(generalCellDesign).style(kdmCellDesign).style(centerTextDesign); //KDM frei Spalte
        } else {
            worksheetToUse.cell(storageCell.y, storageCell.x+2).string(" ").style(generalCellDesign);
        }
*/