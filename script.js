document.getElementById('process-file').addEventListener('click', async () => {
    const fileInput = document.getElementById('file-input');
    if (!fileInput.files.length) {
        alert("Please select the Excel file first.");
        return;
    }

    const file = fileInput.files[0];
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(await file.arrayBuffer());

    const mainSheet = workbook.getWorksheet("Product");
    const variantSheet = workbook.getWorksheet("Variant");

    if (!mainSheet || !variantSheet) {
        alert("Required sheets are not found in the Excel file.");
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
        const titleRowIndex = 2 + Math.floor(i / 2);
        const title = mainSheet.getCell(`B${titleRowIndex}`).value;

        let rowValues = mainSheet.getRow(2).values;
        let newRow = mainSheet.insertRow(insertPosition);

        const currentHandle = `Handle ${i + 1}`;
        newRow.getCell(1).value = currentHandle;
        newRow.getCell(2).value = title;

        rowValues.forEach((value, index) => {
            if (typeof value === 'string' && value.includes('[variant]')) {
                newRow.getCell(index).value = value.replace('[variant]', variant);
            } else {
                newRow.getCell(index).value = value;
            }
        });

        insertPosition++;

        rowValues = mainSheet.getRow(3).values;
        newRow = mainSheet.insertRow(insertPosition);

        newRow.getCell(1).value = currentHandle;
        newRow.getCell(2).value = title;

        rowValues.forEach((value, index) => {
            if (typeof value === 'string' && value.includes('[variant]')) {
                newRow.getCell(index).value = value.replace('[variant]', variant);
            } else {
                newRow.getCell(index).value = value;
            }
        });

        insertPosition++;
    }

    mainSheet.spliceRows(2, 2);

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "Lak Matt Flamant Final.xlsx";
    link.click();
});

document.getElementById('sort-titles').addEventListener('click', async () => {
    const fileInput = document.getElementById('file-input');
    if (!fileInput.files.length) {
        alert("Please select the Excel file first.");
        return;
    }

    const file = fileInput.files[0];
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(await file.arrayBuffer());

    const mainSheet = workbook.getWorksheet("Product");

    if (!mainSheet) {
        alert("Product sheet not found.");
        return;
    }

    let previousTitle = mainSheet.getCell('B2').value;
    let previousHandle = mainSheet.getCell('A2').value;

    let rowIndex = 2;
    while (rowIndex <= mainSheet.rowCount) {
        if (rowIndex % 2 === 0) {
            mainSheet.getRow(rowIndex).getCell(1).value = previousHandle;
            mainSheet.getRow(rowIndex).getCell(2).value = previousTitle;
        } else {
            previousHandle = mainSheet.getRow(rowIndex).getCell(1).value;
            previousTitle = mainSheet.getRow(rowIndex).getCell(2).value;
        }
        rowIndex++;
    }

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "Lak Matt Flamant Sorted Titles and Handles.xlsx";
    link.click();
});
