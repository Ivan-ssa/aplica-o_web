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
// maintenanceSnPatrimColumnNames permanece igual, já que busca o SN
const maintenanceSnPatrimColumnNames = [
    'Nº Série', 'NUMERO_SERIE', 'NUMERO DE SERIE', 'SN', 'PATRIMONIO', 'PATRIM', 'ASSET TAG', 'SERIAL',
    'NÚMERO DE SÉRIE',      
    'N. DE SERIE',          
    'N. DE SÉRIE',          
    'N DE SERIE',           
    'Nº SERIE',             
    'N° DE SÉRIE',          
    'N° DE SERIE',
    'Nº de Série'           
]; 
// maintenanceStatusColumnNames não é mais usado por parseMaintenanceSheet
const maintenanceStatusColumnNames = ['STATUS', 'STATUS_MANUTENCAO', 'SITUACAO', 'STATE', 'SITUATION']; 


// Função auxiliar para encontrar o nome da coluna correto (case-insensitive e trim)
const findHeaderName = (headers, possibleNames) => {
    const normalizeString = (str) => {
        if (!str) return '';
        return String(str).trim()
            .normalize("NFD")
            .replace(/[\u0300-\u036f]/g, "")
            .toLowerCase()
            .replace(/[^a-z0-9 ]/g, ''); 
    };

    const normalizedHeaders = headers.map(h => normalizeString(h));

    for (const name of possibleNames) {
        const normalizedName = normalizeString(name);
        
        if (normalizedHeaders.includes(normalizedName)) {
            return headers[normalizedHeaders.indexOf(normalizedName)];
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

// NOVA FUNÇÃO: Parser para a Planilha de Manutenção Externa (APENAS SN/PATRIMÔNIO)
export const parseMaintenanceSheet = (worksheet) => {
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false, defval: '' });
    if (jsonData.length === 0) { 
        console.warn("DEBUG: Planilha de Manutenção vazia ou sem dados após cabeçalho.");
        return [];
    }

    const headers = jsonData[0].map(h => String(h).trim());
    console.log('DEBUG: Cabeçalhos da planilha de Manutenção:', headers); 
    
    // Agora, apenas idHeader é necessário
    const idHeader = findHeaderName(headers, maintenanceSnPatrimColumnNames);
    // statusHeader não é mais necessário aqui, mas o log mostra que ele está sendo encontrado

    console.log('DEBUG: idHeader encontrado para Manutenção (APENAS ID NECESSÁRIO):', idHeader); 

    // A validação agora é APENAS para o ID
    if (!idHeader) {
        console.warn("DEBUG: Planilha de Manutenção ignorada por não conter coluna de ID (SN/Patrimônio) essencial.");
        return [];
    }

    const dataRows = jsonData.slice(1);
    console.log('DEBUG: Linhas de dados de Manutenção (dataRows):', dataRows); 

    return dataRows.map(row => {
        let obj = {};
        headers.forEach((header, index) => {
            obj[header] = row[index] !== undefined ? String(row[index]).trim() : '';
        });
        
        // Retorna apenas o SN/Patrimônio normalizado
        return (obj[idHeader] ? String(obj[idHeader]).replace(/^0+/, '').trim() : ''); 
    }).filter(id => id !== ''); // Filtra IDs vazios
};
