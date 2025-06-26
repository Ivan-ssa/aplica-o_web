// js/excelReader.js

// Mapeamentos de nomes de colunas alternativos para as chaves padronizadas
const snColumnNames = ['SN', 'NUMERO_SERIE', 'NUMERO DE SERIE', 'SERIAL_NUMBER', 'SERIAL NO', 'NÚMERO DE SÉRIE']; 
const dataValColumnNames = ['DATA VAL', 'DATA_VALIDADE', 'DATA VALIDADE', 'VALIDADE', 'VALIDITY_DATE', 'VENCIMENTO'];
const dataCalColumnNames = ['DATA CAL', 'DATA_CALIBRACAO', 'DATA_CAL', 'DATA DE SAIDA', 'DATA DE CRIACAO'];

const equipamentoColumnNames = ['EQUIPAMENTO', 'TIPO DE EQUIPAMENTO', 'NOME EQUIPAMENTO'];
const fabricanteColumnNames = ['FABRICANTE', 'MARCA', 'MANUFACTURER'];
const modeloColumnNames = ['MODELO', 'MODEL'];
const patrimonioColumnNames = ['PATRIM', 'PATRIMONIO', 'ASSET TAG'];
const tipoServicoColumnNames = ['TIPO SERVICO', 'TIPO_SERVICO', 'SERVICE TYPE'];

// NOVOS: Mapeamentos para Manutenção Externa
// ATENÇÃO: Adicionado 'Nº de Série' (exatamente como você informou)
const maintenanceSnPatrimColumnNames = [
    'Nº Série', 'NUMERO_SERIE', 'NUMERO DE SERIE', 'SN', 'PATRIMONIO', 'PATRIM', 'ASSET TAG', 'SERIAL',
    'NÚMERO DE SÉRIE',      
    'N. DE SERIE',          
    'N. DE SÉRIE',          
    'N DE SERIE',           
    'Nº SERIE',             
    'N° DE SÉRIE',          
    'N° DE SERIE',
    'Nº de Série'           // <--- ESTE É O NOVO TERMO EXATO QUE FALTAVA
]; 
const maintenanceStatusColumnNames = ['STATUS', 'STATUS_MANUTENCAO', 'SITUACAO', 'STATE', 'SITUATION']; 


// Função auxiliar para encontrar o nome da coluna correto (case-insensitive e trim)
const findHeaderName = (headers, possibleNames) => {
    const lowerCaseHeaders = headers.map(h => h.toLowerCase()); 
    for (const name of possibleNames) {
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
        obj.maintenanceStatus = 'Não Aplicável'; 
        return obj;
    });
};

export const parseCalibrationSheet = (worksheet) => {
    const jsonDataRaw = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: true, defval: '' });
    if (jsonDataRaw.length === 0) return [];

    const headers = jsonDataRaw[0].map(h => String(h).trim());

    const snHeader = findHeaderName(headers, snColumnNames);
    const dataValHeader = findHeaderName(headers, dataValColumnNames);
    const dataCalHeader = findHeaderName(headers, dataCalColumnNames);

    const equipamentoHeader = findHeaderName(headers, equipamentoColumnNames);
    const fabricanteHeader = findHeaderName(headers, fabricanteColumnNames);
    const modeloHeader = findHeaderName(headers, modeloColumnNames);
    const patrimonioHeader = findHeaderName(headers, patrimonioColumnNames);
    const tipoServicoHeader = findHeaderName(headers, tipoServicoColumnNames);

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

// NOVA FUNÇÃO: Parser para a Planilha de Manutenção Externa
export const parseMaintenanceSheet = (worksheet) => {
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false, defval: '' });
    if (jsonData.length === 0) { 
        console.warn("DEBUG: Planilha de Manutenção vazia ou sem dados após cabeçalho.");
        return [];
    }

    const headers = jsonData[0].map(h => String(h).trim());
    console.log('DEBUG: Cabeçalhos da planilha de Manutenção:', headers); 
    
    const idHeader = findHeaderName(headers, maintenanceSnPatrimColumnNames);
    const statusHeader = findHeaderName(headers, maintenanceStatusColumnNames);

    console.log('DEBUG: idHeader encontrado para Manutenção:', idHeader); 
    console.log('DEBUG: statusHeader encontrado para Manutenção:', statusHeader); 

    if (!idHeader || !statusHeader) {
        console.warn("DEBUG: Planilha de Manutenção ignorada por não conter colunas essenciais (SN/Patrimônio e Status).");
        return [];
    }

    const dataRows = jsonData.slice(1);
    console.log('DEBUG: Linhas de dados de Manutenção (dataRows):', dataRows); 

    return dataRows.map(row => {
        let obj = {};
        headers.forEach((header, index) => {
            obj[header] = row[index] !== undefined ? String(row[index]).trim() : '';
        });
        
        obj['SN_PATRIM_MANUTENCAO'] = (obj[idHeader] ? String(obj[idHeader]).replace(/^0+/, '').trim() : ''); 
        obj['STATUS_MANUTENCAO_EXTERNA'] = obj[statusHeader] || 'Desconhecido'; 

        return obj;
    });
};
