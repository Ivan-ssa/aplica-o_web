// js/excelReader.js

// Mapeamentos de nomes de colunas alternativos para as chaves padronizadas
// Adicione aqui outros nomes de colunas que você encontrar em outras planilhas.

// Para o Número de Série (SN)
const snColumnNames = ['SN', 'NUMERO_SERIE', 'NUMERO DE SERIE', 'SERIAL_NUMBER', 'SERIAL NO'];

// Para a Data de Validade da Calibração (DATA VAL) - Sciencetech não tem, mas outras podem ter
const dataValColumnNames = ['DATA VAL', 'DATA_VALIDADE', 'DATA VALIDADE', 'VALIDADE', 'VALIDITY_DATE', 'VENCIMENTO'];

// Para a Data de Calibração (DATA CAL) - Inclui 'DATA DE CRIACAO'
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
    const lowerCaseHeaders = headers.map(h => h.toLowerCase()); // Converte todos os cabeçalhos para minúsculas
    for (const name of possibleNames) {
        if (lowerCaseHeaders.includes(name.toLowerCase())) { // Compara com nomes possíveis em minúsculas
            // Retorna o nome original da coluna como encontrado nos headers
            return headers[lowerCaseHeaders.indexOf(name.toLowerCase())];
        }
    }
    return null; // Retorna null se não encontrar nenhum dos nomes possíveis
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

    // Encontrar o nome correto das colunas usando os mapeamentos
    const snHeader = findHeaderName(headers, snColumnNames);
    const dataValHeader = findHeaderName(headers, dataValColumnNames);
    const dataCalHeader = findHeaderName(headers, dataCalColumnNames);

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
            const value = row[index]; // Pega o valor bruto

            // Converte valores numéricos de data do Excel para strings formatadas
            if (header === dataValHeader && typeof value === 'number') {
                obj['DATA VAL'] = XLSX.SSF.format('mm/yyyy', value); // Padroniza para 'DATA VAL'
            } else if (header === dataCalHeader && typeof value === 'number') {
                obj['DATA CAL'] = XLSX.SSF.format('dd/mm/yyyy', value); // Padroniza para 'DATA CAL'
            }
            // Para outros campos ou se não for data, apenas copia o valor
            obj[header] = value !== undefined ? String(value).trim() : '';
        });

        // Garante que as chaves padronizadas existam no objeto 'obj',
        // usando os valores encontrados pelos cabeçalhos mapeados
        obj['SN'] = obj[snHeader] || ''; // Padroniza para 'SN'
        obj['DATA VAL'] = obj['DATA VAL'] || (dataValHeader ? obj[dataValHeader] : 'N/A'); // Padroniza
        obj['DATA CAL'] = obj['DATA CAL'] || (dataCalHeader ? obj[dataCalHeader] : 'N/A'); // Padroniza
        
        obj['EQUIPAMENTO'] = obj[equipamentoHeader] || ''; // Padroniza
        obj['FABRICANTE'] = obj[fabricanteHeader] || ''; // Padroniza
        obj['MODELO'] = obj[modeloHeader] || ''; // Padroniza
        obj['PATRIM'] = obj[patrimonioHeader] || ''; // Padroniza
        obj['TIPO SERVICO'] = obj[tipoServicoHeader] || ''; // Padroniza


        // Limpeza adicional e tratamento de dados importantes
        obj['SN'] = String(obj['SN']).trim().replace(/^0+/, ''); // Garante SN limpo para cruzamento
        // Se DATA VAL for 'N/A' e DATA CAL existir, pode-se tentar estimar, mas por enquanto, manter N/A
        // ou adicionar uma flag de "Data Validade Indefinida"

        return obj;
    });
};
