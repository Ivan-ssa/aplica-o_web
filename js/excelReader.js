// js/excelReader.js

// Mapeamentos de nomes de colunas alternativos para as chaves padronizadas
const snColumnNames = ['SN', 'NUMERO_SERIE', 'NUMERO DE SERIE', 'SERIAL_NUMBER', 'SERIAL NO', 'NÚMERO DE SÉRIE', 'Nº Série', 'Nº DE SÉRIE'];
const dataValColumnNames = ['DATA VAL', 'DATA_VALIDADE', 'DATA VALIDADE', 'VALIDADE', 'VALIDITY_DATE', 'VENCIMENTO'];
const dataCalColumnNames = ['DATA CAL', 'DATA_CALIBRACAO', 'DATA_CAL', 'DATA DE SAIDA', 'DATA DE CRIACAO'];

const equipamentoColumnNames = ['EQUIPAMENTO', 'TIPO DE EQUIPAMENTO', 'NOME EQUIPAMENTO'];
const fabricanteColumnNames = ['FABRICANTE', 'MARCA', 'MANUFACTURER'];
const modeloColumnNames = ['MODELO', 'MODEL'];
const patrimonioColumnNames = ['PATRIM', 'PATRIMONIO', 'ASSET TAG'];
const tipoServicoColumnNames = ['TIPO SERVICO', 'TIPO_SERVICO', 'SERVICE TYPE'];

// NOVOS: Mapeamentos para Manutenção Externa - Foco na coluna Q
const maintenanceSnPatrimColumnNames = [
    'Nº Série', 'NUMERO_SERIE', 'NUMERO DE SERIE', 'SN', 'PATRIMONIO', 'PATRIM', 'ASSET TAG', 'SERIAL',
    'NÚMERO DE SÉRIE', 'N. DE SERIE', 'N. DE SÉRIE', 'N DE SERIE', 'Nº SERIE', 'N° DE SÉRIE', 'N° DE SERIE',
    'ID', 'IDENTIFICADOR', 'CODIGO', 'CÓDIGO', 'ITEM', 'TAG' // Adicionei mais opções genéricas para a coluna Q
];
const maintenanceStatusColumnNames = ['STATUS', 'STATUS_MANUTENCAO', 'SITUACAO', 'STATE', 'SITUATION'];


// Função para normalizar IDs (removida do main.js, agora aqui para uso geral)
const normalizeIdForComparison = (id) => {
    if (!id) return '';
    // Remove zeros à esquerda, espaços, converte para minúsculas e REMOVE TUDO QUE NÃO FOR ALFANUMÉRICO (letras e números)
    let normalized = String(id).replace(/^0+/, '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');
    // console.log(`DEBUG: Normalizando '${id}' para '${normalized}'`); // DEBUG: Ver o processo de normalização
    return normalized;
};

// Função auxiliar para encontrar o nome da coluna correto (case-insensitive, trim, e normalizado para acentos/especiais)
const findHeaderName = (headers, possibleNames) => {
    const normalizeString = (str) => {
        if (!str) return '';
        return String(str).trim()
            .normalize("NFD")
            .replace(/[\u0300-\u036f]/g, "")
            .toLowerCase()
            .replace(/[^a-z0-9 ]/g, ''); // Mantém espaços para nomes de colunas com espaço
    };

    const normalizedHeaders = headers.map(h => normalizeString(h));
    // console.log('DEBUG: Cabeçalhos normalizados:', normalizedHeaders); // DEBUG: Ver cabeçalhos normalizados

    for (const name of possibleNames) {
        const normalizedName = normalizeString(name);
        // console.log(`DEBUG: Tentando encontrar '${normalizedName}'`); // DEBUG: Ver nomes sendo procurados

        if (normalizedHeaders.includes(normalizedName)) {
            const originalHeader = headers[normalizedHeaders.indexOf(normalizedName)];
            // console.log(`DEBUG: Cabeçalho encontrado: '${originalHeader}' para nome normalizado '${normalizedName}'`); // DEBUG: Confirmação
            return originalHeader;
        }
    }
    // console.log(`DEBUG: Nenhum cabeçalho encontrado para os nomes possíveis:`, possibleNames); // DEBUG: Se não encontrar
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

// Modificado: parseEquipmentSheet agora normaliza IDs para o formato de comparação
export const parseEquipmentSheet = (worksheet) => {
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false, defval: '' });
    if (jsonData.length === 0) return [];

    const headers = jsonData[0].map(h => String(h).trim());
    // console.log('DEBUG: Headers da planilha Equipamentos:', headers); // DEBUG: Ver cabeçalhos originais
    const dataRows = jsonData.slice(1);

    return dataRows.map((row, rowIndex) => {
        let obj = {};
        headers.forEach((header, index) => {
            const value = row[index] !== undefined ? String(row[index]).trim() : '';
            // NOVO: Armazenar o valor original do Nº Série, mas normalizar o valor principal para busca
            if (header === 'Nº Série') {
                obj['Nº Série Original'] = value; // Guarda o valor original para exibição
                obj[header] = normalizeIdForComparison(value); // Normaliza para a busca e processamento
            } else if (header === 'Patrimônio') { // Patrimônio SÓ é normalizado, sem guardar o original
                obj[header] = normalizeIdForComparison(value);
            } else {
                obj[header] = value;
            }
        });
        obj.calibrationStatus = 'Desconhecido';
        obj.calibrations = [];
        obj.nextCalibrationDate = 'N/A';
        obj.maintenanceStatus = 'Não Aplicável';

        // DEBUG: Log do Nº Série e Patrimônio normalizados para cada equipamento
        // if (rowIndex < 5) { // Limitar para não inundar o console
        //     console.log(`DEBUG Equipamento[${rowIndex}]: Nº Série original: '${row[headers.indexOf('Nº Série')] || ''}' -> Normalizado: '${obj['Nº Série']}'`);
        //     console.log(`DEBUG Equipamento[${rowIndex}]: Patrimônio original: '${row[headers.indexOf('Patrimônio')] || ''}' -> Normalizado: '${obj['Patrimônio']}'`);
        // }
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

        // NOVO: Normalizar SN de calibração aqui também para consistência
        obj['SN'] = normalizeIdForComparison(obj[snHeader] || '');

        obj['DATA VAL'] = obj['DATA VAL'] || (dataValHeader ? obj[dataValHeader] : 'N/A');
        obj['DATA CAL'] = obj['DATA CAL'] || (dataCalHeader ? obj[dataCalHeader] : 'N/A');

        obj['EQUIPAMENTO'] = obj[equipamentoHeader] || '';
        obj['FABRICANTE'] = obj[fabricanteHeader] || '';
        obj['MODELO'] = obj[modeloHeader] || '';
        obj['PATRIM'] = obj[patrimonioHeader] || '';
        obj['TIPO SERVICO'] = obj[tipoServicoHeader] || '';

        return obj;
    });
};

// Parser para a Planilha de Manutenção Externa (retorna SNs normalizados)
export const parseMaintenanceSheet = (worksheet) => {
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false, defval: '' });
    if (jsonData.length === 0) {
        console.warn("DEBUG: Planilha de Manutenção vazia ou sem dados após cabeçalho.");
        return [];
    }

    const headers = jsonData[0].map(h => String(h).trim());
    // console.log('DEBUG: Headers da planilha de Manutenção:', headers); // DEBUG: Ver cabeçalhos originais

    const idHeader = findHeaderName(headers, maintenanceSnPatrimColumnNames);
    // const statusHeader = findHeaderName(headers, maintenanceStatusColumnNames); // Manter statusHeader para logs

    // console.log('DEBUG: idHeader encontrado para Manutenção:', idHeader);
    // console.log('DEBUG: statusHeader encontrado para Manutenção:', statusHeader);

    if (!idHeader) { // Validação agora é APENAS para idHeader
        console.warn("DEBUG: Planilha de Manutenção ignorada por não conter coluna de ID (SN/Patrimônio) essencial.");
        return [];
    }

    const dataRows = jsonData.slice(1);
    // console.log('DEBUG: Linhas de dados de Manutenção (dataRows):', dataRows);

    return dataRows.map((row, rowIndex) => {
        let obj = {};
        headers.forEach((header, index) => {
            obj[header] = row[index] !== undefined ? String(row[index]).trim() : '';
        });

        // Retorna o SN/Patrimônio normalizado
        const originalSN = obj[idHeader] || '';
        const normalizedSN = normalizeIdForComparison(originalSN);
        // DEBUG: Log do SN normalizado para cada item de manutenção
        // if (rowIndex < 5) { // Limitar para não inundar o console
        //     console.log(`DEBUG Manutenção[${rowIndex}]: SN original: '${originalSN}' -> Normalizado: '${normalizedSN}'`);
        // }
        return normalizedSN;
    }).filter(id => id !== '');
};
