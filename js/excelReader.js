// js/excelReader.js

// Mapeamentos de nomes de colunas alternativos para as chaves padronizadas

// Para o Número de Série (SN)
// MUDANÇA AQUI: Adicionado 'NÚMERO DE SÉRIE'
const snColumnNames = ['SN', 'NUMERO_SERIE', 'NUMERO DE SERIE', 'SERIAL_NUMBER', 'SERIAL NO', 'NÚMERO DE SÉRIE']; 

// Para a Data de Validade da Calibração (DATA VAL) - Sciencetech não tem, mas outras podem ter
const dataValColumnNames = ['DATA VAL', 'DATA_VALIDADE', 'DATA VALIDADE', 'VALIDADE', 'VALIDITY_DATE', 'VENCIMENTO'];

// Para a Data de Calibração (DATA CAL)
const dataCalColumnNames = ['DATA CAL', 'DATA_CALIBRACAO', 'DATA_CAL', 'DATA DE SAIDA', 'DATA DE CRIACAO'];

// Para o nome do Equipamento (EQUIPAMENTO)
const equipamentoColumnNames = ['EQUIPAMENTO', 'TIPO DE EQUIPAMENTO', 'NOME EQUIPAMENTO'];

// Para o Fabricante/Marca (FABRICANTE)
const fabricanteColumnNames = ['FABRICANTE', 'MARCA', 'MANUFACTURER'];

// Para o Modelo (MODELO)
const modeloColumnNames = ['MODELO', 'MODEL'];

// Para o Patrimônio (PATRIM)
const patrimonioColumnNames = ['PATRIM', 'PATRIMONIO', 'ASSET TAG'];

// Para o Tipo de Serviço (TIPO SERVICO)
const tipoServicoColumnNames = ['TIPO SERVICO', 'TIPO_SERVICO', 'SERVICE TYPE'];


// Função auxiliar para encontrar o nome da coluna correto (case-insensitive e trim)
const findHeaderName = (headers, possibleNames) => {
    const lowerCaseHeaders = headers.map(h => h.toLowerCase()); 
    for (const name of possibleNames) {
        // Compara com nomes possíveis em minúsculas (incluindo o original com acento)
        if (lowerCaseHeaders.includes(name.toLowerCase())) { 
            return headers[lowerCaseHeaders.indexOf(name.toLowerCase())];
        }
    }
    return null; 
};


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
    console.log('CABEÇALHOS LIDOS DA PLANILHA DE CALIBRAÇÃO:', headers);

    // Encontrar o nome correto das colunas usando os mapeamentos
    const snHeader = findHeaderName(headers, snColumnNames);
    
    const dataValHeader = findHeaderName(headers, dataValColumnNames);
    const dataCalHeader = findHeaderName(headers, dataCalColumnNames);
    console.log('SN_HEADER ENCONTRADO PELA LÓGICA:', snHeader);
    const equipamentoHeader = findHeaderName(headers, equipamentoColumnNames);
    const fabricanteHeader = findHeaderName(headers, fabricanteColumnNames);
    const modeloHeader = findHeaderName(headers, modeloColumnNames);
    const patrimonioHeader = findHeaderName(headers, patrimonioColumnNames);
    const tipoServicoHeader = findHeaderName(headers, tipoServicoColumnNames);


    // Verifica se a coluna de Número de Série (SN) é essencial para identificar o item de calibração
    if (!snHeader) {
        console.warn("Planilha ignorada por não conter coluna de Número de Série essencial para calibração.");
        return [];
    }

    const dataRows = jsonDataRaw.slice(1);

    return dataRows.map(row => {
        let obj = {};
        headers.forEach((header, index) => {
            const value = row[index];

            if (header === dataValHeader && typeof value === 'number') {
                obj['DATA VAL'] = XLSX.SSF.format('mm/yyyy', value); 
            } else if (header === dataCalHeader && typeof value === 'number') {
                obj['DATA CAL'] = XLSX.SSF.format('dd/mm/yyyy', value); 
            }
            obj[header] = value !== undefined ? String(value).trim() : '';
        });

        obj['SN'] = obj[snHeader] || ''; 
        obj['DATA VAL'] = obj['DATA VAL'] || (dataValHeader ? obj[dataValHeader] : 'N/A'); 
        obj['DATA CAL'] = obj['DATA CAL'] || (dataCalHeader ? obj[dataCalHeader] : 'N/A'); 
        
        obj['EQUIPAMENTO'] = obj[equipamentoHeader] || ''; 
        obj['FABRICANTE'] = obj[fabricanteHeader] || ''; 
        obj['MODELO'] = obj[modeloHeader] || ''; 
        obj['PATRIM'] = obj[patrimonioHeader] || ''; 
        obj['TIPO SERVICO'] = obj[tipoServicoHeader] || ''; 

        obj['SN'] = String(obj['SN']).trim().replace(/^0+/, ''); 

        return obj;
    });
};
