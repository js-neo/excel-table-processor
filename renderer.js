let mainFilePath = '';
let avrFilePath = '';

document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('selectMainFileButton').addEventListener('click', async () => {
        console.log('Кнопка выбора основного файла нажата');
        mainFilePath = await window.electron.ipcRenderer.invoke('dialog:openFile');
        if (mainFilePath) {
            console.log('Выбран основной файл:', mainFilePath);
        }
    });

    document.getElementById('selectAvrFileButton').addEventListener('click', async () => {
        console.log('Кнопка выбора файла АВР нажата');
        avrFilePath = await window.electron.ipcRenderer.invoke('dialog:openFile');
        if (avrFilePath) {
            console.log('Выбран файл АВР:', avrFilePath);
        }
    });

    document.getElementById('processFilesButton').addEventListener('click', async () => {
        if (mainFilePath && avrFilePath) {
            await processExcelFiles(mainFilePath, avrFilePath);
        } else {
            alert('Пожалуйста, выберите оба файла.');
        }
    });
});

async function processExcelFiles(mainFilePath, avrFilePath) {
    const mainWorkbook = new window.electron.ExcelJS.Workbook();
    await mainWorkbook.xlsx.readFile(mainFilePath);
    const mainSheet = mainWorkbook.worksheets[0];

    const avrWorkbook = new window.electron.ExcelJS.Workbook();
    await avrWorkbook.xlsx.readFile(avrFilePath);
    const avrSheet = avrWorkbook.worksheets[0];

    const mainData = [];
    mainSheet.eachRow((row) => {
        const rowData = {};
        row.eachCell((cell, colNumber) => {
            rowData[`Column${colNumber}`] = cell.value;
        });
        mainData.push(rowData);
    });

    const avrData = [];
    avrSheet.eachRow((row) => {
        const rowData = {};
        row.eachCell((cell, colNumber) => {
            rowData[`Column${colNumber}`] = cell.value;
        });
        avrData.push(rowData);
    });

    function checkUniqueness(data, columnIndex) {
        const uniqueValues = new Set();
        const duplicates = [];

        data.forEach(row => {
            const value = row[`Column${columnIndex}`];
            if (value) {
                if (uniqueValues.has(value)) {
                    duplicates.push(value);
                } else {
                    uniqueValues.add(value);
                }
            }
        });

        return duplicates;
    }

    const workNamesDuplicates = checkUniqueness(mainData, 1);

    if (workNamesDuplicates.length > 0) {
        console.log('Найдены дубликаты:', workNamesDuplicates);
    } else {
        console.log('Все наименования уникальны.');
    }

    mainData.forEach(row => {
        const avrRow = avrData.find(avr => avr[`Column1`] === row[`Column1`]);

        if (avrRow) {
            row[`Column2`] = avrRow[`Column2`] || 0; // Количество АВР 1
            row[`Column3`] = row[`Column2`] * row[`Column4`]; // Стоимость АВР 1

            row[`Column5`] += row[`Column2`]; // Кол-во выполнено
            row[`Column6`] = row[`Column5`] * row[`Column4`]; // Стоимость выполнено

            row[`Column7`] = row[`Column8`] - row[`Column5`]; // Кол-во остаток
            row[`Column9`] = row[`Column7`] * row[`Column4`]; // Стоимость остаток

            row[`Column10`] = row[`Column9`] < 0 ? Math.abs(row[`Column9`]) : 0; // Перерасход
        } else {
            row[`Column2`] = 0; // Количество АВР 1
            row[`Column3`] = 0; // Стоимость АВР 1
        }
    });

    const newWorkbook = new window.electron.ExcelJS.Workbook();
    const newSheet = newWorkbook.addWorksheet('Обновленная таблица');
    newSheet.addRows(mainData);

    const outputFilePath = 'updated_table.xlsx';
    await newWorkbook.xlsx.writeFile(outputFilePath);
    alert(`Файл сохранен как ${outputFilePath}`);
}
