// js/excelReader.js

// ... (código existente, incluindo as definições de snColumnNames, etc.) ...

// Modificado: parseEquipmentSheet agora normaliza IDs para o formato de comparação
export const parseEquipmentSheet = (worksheet) => {
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false, defval: '' });
    if (jsonData.length === 0) return [];

    const headers = jsonData[0].map(h => String(h).trim());
    // console.log('DEBUG: Headers da planilha Equipamentos:', headers); // DEBUG: Ver cabeçalhos originais
    const dataRows = jsonData.slice(1);

    // NOVO: Encontrar o cabeçalho do Número de Série de forma flexível para a planilha de equipamentos
    const snHeaderForEquipment = findHeaderName(headers, snColumnNames);
    const patrimonioHeaderForEquipment = findHeaderName(headers, patrimonioColumnNames); // Manter para Patrimônio também

    return dataRows.map((row, rowIndex) => {
        let obj = {};
        headers.forEach((header, index) => {
            const value = row[index] !== undefined ? String(row[index]).trim() : '';

            // Agora, usamos snHeaderForEquipment para identificar a coluna de SN
            if (header === snHeaderForEquipment) { // Usa o nome do cabeçalho encontrado
                obj['Nº Série Original'] = value; // Guarda o valor original para exibição
                obj['Nº Série'] = normalizeIdForComparison(value); // Normaliza para a busca e processamento
            } else if (header === patrimonioHeaderForEquipment) { // Para Patrimônio, continua apenas normalizando
                obj['Patrimônio'] = normalizeIdForComparison(value);
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
        //     console.log(`DEBUG Equipamento[${rowIndex}]: Nº Série original: '${obj['Nº Série Original'] || ''}' -> Normalizado: '${obj['Nº Série']}'`);
        //     console.log(`DEBUG Equipamento[${rowIndex}]: Patrimônio original: '${row[headers.indexOf(patrimonioHeaderForEquipment)] || ''}' -> Normalizado: '${obj['Patrimônio']}'`);
        // }
        return obj;
    });
};

// ... (Restante do código do arquivo excelReader.js) ...
