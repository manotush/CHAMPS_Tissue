let data = [];
let serialNumberCounter = 1; // Initialize the serial number counter
let boxNumberCounter = 1; // Initialize the box number counter

function addData() {
    const Sl_No = serialNumberCounter++; // Automatically increment the serial number
    const CHAMPS_ID = document.getElementById('idNumber1').value;
    const Kit_ID = document.getElementById('idNumber2').value;
    const Cass = document.getElementById('idNumber2').value;
    const levels = Array.from({ length: 9 }, (_, i) => document.getElementById(`level${i + 1}`).value);
    
    levels.forEach((Specimens, index) => {
        if (Specimens) {
            let Tissue_Type;
            let Cassette_ID;
            switch (index + 1) {
                case 1:
                    Tissue_Type = "Cassette A (Liver and Abdominal Organs)";
                    Cassette_ID = `M0${Cass}.042`;
                    break;
                case 2:
                    Tissue_Type = "Cassette B (Right Lung)";
                    Cassette_ID = `M0${Cass}.044`;
                    break;
                case 3:
                    Tissue_Type = "Cassette C (Left Thoracic Organs, Left Lung and Heart)";
                    Cassette_ID = `M0${Cass}.046`;
                    break;
                case 4:
                    Tissue_Type = "Cassette D (Central Nervous System, Posterior Fossa and Fontanelle)";
                    Cassette_ID = `M0${Cass}.048`;
                    break;
                case 5:
                    Tissue_Type = "Cassette E (Central Nervous System, Trans-nasal)";
                    Cassette_ID = `M0${Cass}.050`;
                    break;
                case 6:
                    Tissue_Type = "Membrane";
                    Cassette_ID = `M0${Cass}.080`;
                    break;
                case 7:
                    Tissue_Type = "Cord";
                    Cassette_ID = `M0${Cass}.082`;
                    break;
                case 8:
                    Tissue_Type = "PL Parench";
                    Cassette_ID = `M0${Cass}.084`;
                    break;
                case 9:
                    Tissue_Type = "PL Parench";
                    Cassette_ID = `M0${Cass}.086`;
                    break;
            }

            const newData = {
                Sl_No,
                CHAMPS_ID: `BDAA0${CHAMPS_ID}`,
                Kit_ID: `M0${Kit_ID}`,
                Cassette_ID,
                No_of_Specimens: `${Specimens} ${Specimens > 1 ? 'pieces' : 'piece'}`,
                Tissue_Type,
                Box_No: Math.ceil(Sl_No / 8)
            };
            data.push(newData);
        }
    });

    console.log('Data added:', data);

    // Clear the form fields after adding the data
    document.getElementById('dataForm').reset();
}

function generateSpreadsheet() {
    // Define the column headings
    const headerRow = {
        Sl_No: 'Sl. No.',
        CHAMPS_ID: 'CHAMPS ID',
        Kit_ID: 'Kit ID',
        Cassette_ID: 'Cassette ID',
        No_of_Specimens: 'No. of Specimens',
        Tissue_Type: 'Tissue Type',
        Box_No: 'Box No.'
    };

    // Start by creating an empty row and then add the header row
    const spreadsheetData = [{}, headerRow, ...data];

    // Create the worksheet from the modified data
    const ws = XLSX.utils.json_to_sheet(spreadsheetData, { skipHeader: true });

    // Ensure the worksheet has the merge property initialized
    if (!ws['!merges']) ws['!merges'] = [];

    // Merge columns A to last column in the first row (index 0)
    const mergeCell = {
        s: { r: 0, c: 0 }, // start cell (first row, first column)
        e: { r: 0, c: Object.keys(headerRow).length - 1 } // end cell (first row, last column)
    };
    ws['!merges'].push(mergeCell);

    // Add the text to the merged cell and apply formatting
    ws['A1'] = { v: "Formalin fixed non-infectious Tissue Cassettes for Central Pathology Laboratory (CPL-CDC)/ CHAMPS Study", s: { alignment: { horizontal: "center" }, font: { bold: true, name: "Times New Roman", sz: 14 } } };

    // Merging cells with the same Sl_No, CHAMPS_ID, Kit_ID, and Box_No
    let mergeRanges = [];
    let startRow = 2; // Start after the intentionally empty row and header row
    let endRow = startRow;

    for (let i = 2; i < spreadsheetData.length; i++) {
        if (i === 2 || spreadsheetData[i].Sl_No !== spreadsheetData[i - 1].Sl_No) {
            if (i > 2 && startRow !== endRow) {
                mergeRanges.push({
                    s: { r: startRow, c: 0 }, // Sl_No
                    e: { r: endRow - 1, c: 0 } // merge Sl_No
                });
                mergeRanges.push({
                    s: { r: startRow, c: 1 }, // CHAMPS_ID
                    e: { r: endRow - 1, c: 1 } // merge CHAMPS_ID
                });
                mergeRanges.push({
                    s: { r: startRow, c: 2 }, // Kit_ID
                    e: { r: endRow - 1, c: 2 } // merge Kit_ID
                });
                mergeRanges.push({
                    s: { r: startRow, c: 6 }, // Box_No
                    e: { r: endRow - 1, c: 6 } // merge Box_No
                });
            }
            startRow = i;
        }
        endRow = i + 1;
    }
    if (startRow !== endRow) {
        mergeRanges.push({
            s: { r: startRow, c: 0 },
            e: { r: endRow - 1, c: 0 }
        });
        mergeRanges.push({
            s: { r: startRow, c: 1 },
            e: { r: endRow - 1, c: 1 }
        });
        mergeRanges.push({
            s: { r: startRow, c: 2 },
            e: { r: endRow - 1, c: 2 }
        });
        mergeRanges.push({
            s: { r: startRow, c: 6 },
            e: { r: endRow - 1, c: 6 }
        });
    }

    ws['!merges'] = ws['!merges'].concat(mergeRanges);

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Data');

    XLSX.writeFile(wb, 'MITS_()_Tissue_Shipment2024.xlsx');
    alert('Spreadsheet generated and downloaded.');
}
