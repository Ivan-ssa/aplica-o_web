// js/tableRenderer.js

/**
 * Renderiza a tabela de equipamentos com os dados fornecidos.
 * @param {Array<Object>} equipments - O array de objetos de equipamentos a serem exibidos.
 * @param {HTMLElement} tableBodyElement - O elemento <tbody> da tabela.
 * @param {Map<string, Object>} consolidatedCalibratedMap - Mapa SN -> { fornecedor, dataCalibracao }.
 * @param {Set<string>} externalMaintenanceSNs - Um conjunto de Números de Série em manutenção externa.
 */
export function renderTable(equipments, tableBodyElement, consolidatedCalibratedMap, externalMaintenanceSNs) {
    tableBodyElement.innerHTML = '';

    if (equipments.length === 0) {
        const row = tableBodyElement.insertRow();
        const cell = row.insertCell();
        cell.colSpan = 10;
        cell.textContent = 'Nenhum equipamento encontrado com os filtros aplicados.';
        cell.style.textAlign = 'center';
        return;
    }

    equipments.forEach(equipment => {
        const row = tableBodyElement.insertRow();

        // É CRUCIAL que esta normalização seja idêntica à de main.js
        const equipmentSN = String(equipment?.NumeroSerie || '').trim();
        let normalizedEquipmentSN;

        // Reproduzir a lógica de normalizeId aqui para depuração
        if (/^\d+$/.test(equipmentSN)) {
            normalizedEquipmentSN = String(parseInt(equipmentSN, 10));
        } else {
            normalizedEquipmentSN = equipmentSN.toLowerCase();
        }

        let displayCalibrationStatus = equipment?.StatusCalibacao || '';
        let displayMaintenanceStatus = equipment?.StatusManutencao || '';
        let displayDataVencimento = equipment?.DataVencimentoCalibacao || ''; // Data original do cadastro principal

        // LÓGICA DE COLORAÇÃO E ATUALIZAÇÃO DE STATUS DE CALIBRAÇÃO
        const calibInfo = consolidatedCalibratedMap.get(normalizedEquipmentSN);

        // --- INÍCIO DO CÓDIGO DE DEBUG ADICIONAL NO RENDER (para data) ---
        // Manter estes logs para depurar a conversão da data serial do Excel.
        if (calibInfo && (normalizedEquipmentSN.toLowerCase().includes('rts') || calibInfo.fornecedor.toLowerCase().includes('rts'))) { // Filtrar logs apenas para RTS
            console.log(`DEBUG DATA - Equipamento: ${equipment.TAG} (SN: ${normalizedEquipmentSN})`);
            console.log(`  calibInfo.dataCalibacao (valor bruto da consolidação):`, calibInfo?.dataCalibacao, `(Tipo: ${typeof calibInfo?.dataCalibacao})`);
            const testDate = parseExcelDate(calibInfo?.dataCalibacao);
            console.log(`  parseExcelDate resultado:`, testDate, `(É Date válido? ${testDate instanceof Date && !isNaN(testDate)})`);
            console.log(`  formatDate resultado:`, formatDate(testDate));
        }
        // --- FIM DO CÓDIGO DE DEBUG ADICIONAL NO RENDER ---

        if (calibInfo) {
            row.classList.add('calibrated-dhme');
            displayCalibrationStatus = `Calibrado (${calibInfo.fornecedor})`;

            // *** ALTERAÇÃO AQUI: REATIVA A CONVERSÃO E FORMATAÇÃO DA DATA ***
            const dataOrigem = parseExcelDate(calibInfo.dataCalibacao);
            if (dataOrigem instanceof Date && !isNaN(dataOrigem)) {
                displayDataVencimento = formatDate(dataOrigem);
            } else {
                // Se a conversão falhar, exibe o valor bruto para depuração (pode ser o número serial)
                displayDataVencimento = calibInfo.dataCalibacao;
            }
            // *****************************************************************
        }
        else {
            // Se não foi calibrado por nenhum fornecedor da consolidação
            const originalCalibStatusLower = String(equipment?.StatusCalibacao || '').toLowerCase();
            if (originalCalibStatusLower.includes('não calibrado') || originalCalibStatusLower.includes('não cadastrado')) {
                 row.classList.add('not-calibrated');
                 displayCalibrationStatus = 'Não Calibrado/Não Encontrado (Seu Cadastro)';
            } else if (originalCalibStatusLower.includes('calibrado (total)')) {
                displayCalibrationStatus = 'Calibrado (Total)';
                displayDataVencimento = equipment?.DataVencimentoCalibacao || '';
            } else {
                displayCalibrationStatus = String(equipment?.StatusCalibacao || '');
                if (displayCalibrationStatus.trim() === '') {
                    displayCalibrationStatus = 'Não Calibrado/Não Encontrado (Seu Cadastro)';
                    row.classList.add('not-calibrated');
                }
                displayDataVencimento = equipment?.DataVencimentoCalibacao || '';
            }
        }

        // LÓGICA DE COLORAÇÃO E ATUALIZAÇÃO DE STATUS DE MANUTENÇÃO EXTERNA
        if (externalMaintenanceSNs.has(normalizedEquipmentSN)) {
            row.classList.add('in-external-maintenance');
            displayMaintenanceStatus = 'Em Manutenção Externa';
        }

        // PREENCHIMENTO DAS CÉLULAS
        row.insertCell().textContent = equipment.TAG ?? '';
        row.insertCell().textContent = equipment.Equipamento ?? '';
        row.insertCell().textContent = equipment.Modelo ?? '';
        row.insertCell().textContent = equipment.Fabricante ?? '';
        row.insertCell().textContent = equipment.Setor ?? '';
        row.insertCell().textContent = equipment.NumeroSerie ?? '';
        row.insertCell().textContent = equipment.Patrimonio ?? '';
        row.insertCell().textContent = displayCalibrationStatus;
        row.insertCell().textContent = displayDataVencimento; // Esta linha usará o valor formatado ou bruto
        // row.insertCell().textContent = displayMaintenanceStatus;
    });
}

// Helper para converter número de data do Excel para data JS
function parseExcelDate(excelDate) {
    if (typeof excelDate === 'number' && excelDate > 0) {
        // Excel baseia suas datas em 1 de janeiro de 1900. JavaScript em 1 de janeiro de 1970.
        // 25569 é a diferença de dias entre 1900-01-01 e 1970-01-01 (mais um dia devido a bug do Excel com 1900-02-29).
        // A data 45814 representa 05/07/2025 no Excel, que é 45814 dias após 1900-01-00.
        // Subtraímos 25569 para ajustar ao epoch do JS (dias desde 1970-01-01), depois multiplicamos por ms/dia.
        const date = new Date(Math.round((excelDate - 25569) * 86400 * 1000));
        date.setUTCHours(0, 0, 0, 0); // Define para meia-noite UTC para evitar problemas de fuso horário.
        return date;
    }
    // Adicionar log para quando o valor não é um número ou é inválido
    console.warn("parseExcelDate: Valor não numérico ou inválido recebido:", excelDate);
    return null;
}

// Helper para formatar data para exibição (dd/mm/yyyy)
function formatDate(date) {
    if (date instanceof Date && !isNaN(date)) {
        const day = String(date.getDate()).padStart(2, '0');
        const month = String(date.getMonth() + 1).padStart(2, '0'); // Mês é 0-indexed
        const year = date.getFullYear();
        return `${day}/${month}/${year}`;
    }
    return '';
}

/**
 * Popula o dropdown de setores com os setores únicos dos equipamentos.
 * @param {Array<Object>} equipments - O array de objetos de equipamentos.
 * @param {HTMLElement} sectorFilterElement - O elemento <select> do filtro de setor.
 */
export function populateSectorFilter(equipments, sectorFilterElement) {
    sectorFilterElement.innerHTML = '<option value="">Todos os Setores</option>';

    const sectors = new Set();
    equipments.forEach(eq => {
        if (eq.Setor && String(eq.Setor).trim() !== '') {
            sectors.add(String(eq.Setor).trim());
        }
    });

    Array.from(sectors).sort().forEach(sector => {
        const option = document.createElement('option');
        option.value = sector;
        option.textContent = sector;
        sectorFilterElement.appendChild(option);
    });
}

/**
 * Atualiza o contador de equipamentos exibidos.
 * @param {number} count - O número de equipamentos a ser exibido.
 */
export function updateEquipmentCount(count) {
    document.getElementById('equipmentCount').textContent = `Total: ${count} equipamentos`;
}