// js/uiRenderer.js
export const renderEquipmentTable = (dataToRender, equipmentTableBody, equipmentCountSpan) => {
    equipmentTableBody.innerHTML = '';
    if (!dataToRender || dataToRender.length === 0) {
        equipmentTableBody.innerHTML = '<tr><td colspan="9">Nenhum equipamento encontrado ou carregado.</td></tr>';
        equipmentCountSpan.textContent = `Total: 0 equipamentos`;
        return;
    }

    dataToRender.forEach(equipment => {
        const row = equipmentTableBody.insertRow();
        if (equipment.calibrationStatus === 'Não Calibrado') {
            row.classList.add('not-calibrated');
        } else if (equipment.calibrationStatus === 'Calibrado') {
            row.classList.add('calibrated');
        }

        row.insertCell().textContent = equipment.TAG || '';
        row.insertCell().textContent = equipment.Equipamento || '';
        row.insertCell().textContent = equipment.Modelo || '';
        row.insertCell().textContent = equipment.Fabricante || '';
        row.insertCell().textContent = equipment.Setor || '';
        row.insertCell().textContent = equipment['Nº Série'] || '';
        row.insertCell().textContent = equipment.Patrimônio || '';

        const statusCell = row.insertCell();
        statusCell.textContent = equipment.calibrationStatus || 'Desconhecido';
        if (equipment.calibrationStatus === 'Calibrado' && equipment.calibrations && equipment.calibrations.length > 0) {
             statusCell.title = equipment.calibrations.map(cal => `Data Cal: ${cal['DATA CAL'] || 'N/A'}, Vencimento: ${cal['DATA VAL'] || 'N/A'}, Tipo: ${cal['TIPO SERVIÇO'] || 'N/A'}`).join('\n');
        }

        const vencimentoCell = row.insertCell();
        vencimentoCell.textContent = equipment.nextCalibrationDate || 'N/A';
    });

    equipmentCountSpan.textContent = `Total: ${dataToRender.length} equipamentos`;
};

export const populateSectorFilter = (equipmentData, sectorFilter) => {
    const sectors = new Set();
    equipmentData.forEach(eq => {
        if (eq.Setor) {
            sectors.add(eq.Setor.trim());
        }
    });

    sectorFilter.innerHTML = '<option value="">Todos os Setores</option>';
    Array.from(sectors).sort().forEach(sector => {
        const option = document.createElement('option');
        option.value = sector;
        option.textContent = sector;
        sectorFilter.appendChild(option);
    });
};

// NOVA FUNÇÃO: Renderizar a tabela de calibrações divergentes
export const renderDivergentCalibrationsTable = (divergentData, divergentTableBody) => {
    divergentTableBody.innerHTML = ''; // Limpa a tabela

    if (!divergentData || divergentData.length === 0) {
        divergentTableBody.innerHTML = '<tr><td colspan="6">Nenhuma calibração com divergência encontrada.</td></tr>';
        return;
    }

    divergentData.forEach(cal => {
        const row = divergentTableBody.insertRow();
        row.insertCell().textContent = cal.SN || '';
        row.insertCell().textContent = cal.EQUIPAMENTO || '';
        row.insertCell().textContent = cal.MARCA || '';
        row.insertCell().textContent = cal.MODELO || '';
        row.insertCell().textContent = cal['DATA VAL'] || '';
        row.insertCell().textContent = cal['TIPO SERVIÇO'] || '';
        // Adicione mais células conforme as colunas que você quer exibir da planilha DHME
    });
};
