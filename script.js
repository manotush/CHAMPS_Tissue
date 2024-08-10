let data = [];

function addData() {
    const Sl_No = document.getElementById('serialNumber').value;
    const CHAMPS_ID = document.getElementById('idNumber1').value;
    const Kit_ID = document.getElementById('idNumber2').value;
    const Cass = document.getElementById('idNumber2').value;
    const levels = Array.from({ length: 9 }, (_, i) => document.getElementById(`level${i + 1}`).value);
    
    levels.forEach((Specimens, index) => {
        if (Specimens && (index + 1 === 1)) {
            const newData = {
                Sl_No,
                CHAMPS_ID,
                Kit_ID,
                Cassette_ID: `${Cass}.042`,
                No_of_Specimens:`${Specimens} pieces`,
                Tissue_Type: `Cassette A (Liver and Abdominal Organs)`,
                Box_No: Math.ceil((data.length + 1) / 10)
            };
            data.push(newData);
        } else if (Specimens && (index + 1 === 2)) {
            const newData = {
                Sl_No,
                CHAMPS_ID,
                Kit_ID,
                Cassette_ID: `${Cass}.044`,
                No_of_Specimens:`${Specimens} pieces`,
                Tissue_Type: `Cassette B (Right Lung)`,
                Box_No: Math.ceil((data.length + 1) / 10)
            };
            data.push(newData);
        } else if (Specimens && (index + 1 === 3)) {
            const newData = {
                Sl_No,
                CHAMPS_ID,
                Kit_ID,
                Cassette_ID: `${Cass}.046`,
                No_of_Specimens:`${Specimens} pieces`,
                Tissue_Type: `Cassette C (Left Thoracic Organs, Left Lung and Heart)`,
                Box_No: Math.ceil((data.length + 1) / 10)
            };
            data.push(newData);
        } else if (Specimens && (index + 1 === 4)) {
            const newData = {
                Sl_No,
                CHAMPS_ID,
                Kit_ID,
                Cassette_ID: `${Cass}.048`,
                No_of_Specimens:`${Specimens} pieces`,
                Tissue_Type: `Cassette D (Central Nervous System, Posterior Fossa and Fontanelle)`,
                Box_No: Math.ceil((data.length + 1) / 10)
            };
            data.push(newData);
        } else if (Specimens && (index + 1 === 5)) {
            const newData = {
                Sl_No,
                CHAMPS_ID,
                Kit_ID,
                Cassette_ID: `${Cass}.050`,
                No_of_Specimens:`${Specimens} pieces`,
                Tissue_Type: `Cassette E (Central Nervous System, Trans-nasal)`,
                Box_No: Math.ceil((data.length + 1) / 10)
            };
            data.push(newData);
        } else if (Specimens && (index + 1 === 6)) {
            const newData = {
                Sl_No,
                CHAMPS_ID,
                Kit_ID,
                Cassette_ID: `${Cass}.080`,
                No_of_Specimens:`${Specimens} piece`,
                Tissue_Type: `Membrane`,
                Box_No: Math.ceil((data.length + 1) / 10)
            };
            data.push(newData);
        } else if (Specimens && (index + 1 === 7)) {
            const newData = {
                Sl_No,
                CHAMPS_ID,
                Kit_ID,
                Cassette_ID: `${Cass}.082`,
                No_of_Specimens:`${Specimens} piece`,
                Tissue_Type: `Cord`,
                Box_No: Math.ceil((data.length + 1) / 10)
            };
            data.push(newData);
        } else if (Specimens && (index + 1 === 8)) {
            const newData = {
                Sl_No,
                CHAMPS_ID,
                Kit_ID,
                Cassette_ID: `${Cass}.084`,
                No_of_Specimens:`${Specimens} piece`,
                Tissue_Type: `PL Parench`,
                Box_No: Math.ceil((data.length + 1) / 10)
            };
            data.push(newData);
        } else if (Specimens && (index + 1 === 9)) {
            const newData = {
                Sl_No,
                CHAMPS_ID,
                Kit_ID,
                Cassette_ID: `${Cass}.086`,
                No_of_Specimens:`${Specimens} piece`,
                Tissue_Type: `PL Parench`,
                Box_No: Math.ceil((data.length + 1) / 10)
            };
            data.push(newData);
        }
    });

    console.log('Data added:', data);

    // Clear the form fields after adding the data
    document.getElementById('dataForm').reset();
}

function generateSpreadsheet() {
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Data');

    XLSX.writeFile(wb, 'MITS_Tissue_Shipment.xlsx');
    alert('Spreadsheet generated and downloaded.');
}
