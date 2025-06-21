// js/excelReader.js
export const readFile = (file) => {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                resolve({ fileName: file.name, workbook: workbook });
            } catch (error) {
                reject(new Error(`Erro ao ler o arquivo ${file.name}: ${error.message}`));
            }
        };
        reader.onerror = (error) => {
            reject(new Error(`Erro ao carregar o arquivo ${file.name}: ${error.message}`));
        };
        reader.readAsArrayBuffer(file);
    });
};

export const parseEquipmentSheet = (worksheet) => {
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false, defval: '' });
    if (jsonData.length === 0) return [];

    const headers = jsonData[0].map(h => String(h).trim());
    const dataRows = jsonData.slice(1);

    return dataRows.map(row => {
        let obj = {};
        headers.forEach((header, index) => {
            obj[header] = row[index] !== undefined ? String(row[index]).trim() : '';
        });
        obj.calibrationStatus = 'Desconhecido';
        obj.calibrations = [];
        obj.nextCalibrationDate = 'N/A';
        return obj;
    });
};

export const parseCalibrationSheet = (worksheet) => {
    const jsonDataRaw = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: true, defval: '' });
    if (jsonDataRaw.length === 0) return [];

    const headers = jsonDataRaw[0].map(h => String(h).trim());
    if (!headers.includes('SN') || !headers.includes('DATA VAL')) {
        return []; // Não é uma planilha de calibração válida ou não tem as colunas chave
    }
    const dataRows = jsonDataRaw.slice(1);

    return dataRows.map(row => {
        let obj = {};
        headers.forEach((header, index) => {
            const value = row[index]; // Pega o valor bruto (número para datas, string para texto)

            if (header === 'DATA VAL' && typeof value === 'number') {
                // Formata o número da data do Excel para a string "MM/YYYY"
                obj[header] = XLSX.SSF.format('mm/yyyy', value);
            } else {
                obj[header] = value !== undefined ? String(value).trim() : ''; // Converte para string e limpa
            }
        });
        return obj;
    });
};