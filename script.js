document.getElementById('processFile').addEventListener('click', async () => {
    const fileInput = document.getElementById('fileInput');
    if (!fileInput.files.length) {
        alert("Kies een bestand");
        return;
    }

    const file = fileInput.files[0];
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(await file.arrayBuffer());

    const mainSheet = workbook.getWorksheet("Product");
    const variantSheet = workbook.getWorksheet("Variant");

    if (!mainSheet || !variantSheet) {
        alert("We kunnen de werkkaarten niet vinden");
        return;
    }

    const variantData = [];
    variantSheet.eachRow((row, rowIndex) => {
        if (rowIndex > 0) {
            variantData.push(row.getCell(1).value);
        }
    });

    let insertPosition = mainSheet.rowCount + 1; 

  
    for (let i = 0; i < variantData.length; i++) {
        const variant = variantData[i];
        const currentHandle = `Handle ${i + 1}`;

     
        for (let rowIndex = 2; rowIndex <= mainSheet.rowCount; rowIndex++) {
            const row = mainSheet.getRow(rowIndex);
            const rowValues = row.values;


            let hasVariantPlaceholder = false;
            rowValues.forEach((value) => {
                if (typeof value === 'string' && value.includes('[variant]')) {
                    hasVariantPlaceholder = true;
                }
            });

            if (hasVariantPlaceholder) {

                let newRow = mainSheet.insertRow(insertPosition);
                newRow.getCell(1).value = currentHandle; 
                newRow.getCell(2).value = mainSheet.getCell('B2').value; 

                rowValues.forEach((value, colIndex) => {
                    if (typeof value === 'string' && value.includes('[variant]')) {
                        newRow.getCell(colIndex).value = value.replace('[variant]', variant);
                    } else {
                        newRow.getCell(colIndex).value = value;
                    }
                });
                insertPosition++;
            }
        }
    }

    for (let rowIndex = mainSheet.rowCount; rowIndex >= 2; rowIndex--) {
        const row = mainSheet.getRow(rowIndex);
        const rowValues = row.values;

        let hasVariantPlaceholder = false;
        rowValues.forEach((value) => {
            if (typeof value === 'string' && value.includes('[variant]')) {
                hasVariantPlaceholder = true;
            }
        });

        if (hasVariantPlaceholder) {
            mainSheet.spliceRows(rowIndex, 1);
        }
    }

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);

    const originalFileName = file.name.split('.').slice(0, -1).join('.');
    const finalFileName = `${originalFileName} Final.xlsx`;
    link.download = finalFileName;
    link.click();
});
