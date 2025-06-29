// js/uiRenderer.js
export const renderEquipmentTable = (dataToRender, equipmentTableBody, equipmentCountSpan) => {
    equipmentTableBody.innerHTML = '';
    if (!dataToRender || dataToRender.length === 0) {
        equipmentTableBody.innerHTML = '<tr><td colspan="10">Nenhum equipamento encontrado ou carregado.</td></tr>';
        equipmentCountSpan.textContent = `Total: 0 equipamentos`;
        return;
    }

    dataToRender.forEach(equipment => {
        const row = equipmentTableBody.insertRow();

        // Lógica de aplicação de classe CSS por status de Calibração
        if (equipment.calibrationStatus === 'Não Calibrado') {
            row.classList.add('not-calibrated');
        } else if (equipment.calibrationStatus === 'Calibrado (DHMED)') {
            row.classList.add('calibrated-dhme');
        } else if (equipment.calibrationStatus === 'Calibrado (Sciencetech)') {
            row.classList.add('calibrated-sciencetech');
        } else if (equipment.calibrationStatus === 'Não Cadastrado (DHMED)') {
            row.classList.add('divergent-calibrated');
        } else if (equipment.calibrationStatus === 'Não Cadastrado (Sciencetech)') {
            row.classList.add('divergent-calibrated-sciencetech');
        }

        // Aplicar classe CSS para Status de Manutenção
        if (equipment.maintenanceStatus && equipment.maintenanceStatus !== 'Não Aplicável') {
            row.classList.add('in-external-maintenance');
        }


        row.insertCell().textContent = equipment.TAG || '';
        row.insertCell().textContent = equipment.Equipamento || '';
        row.insertCell().textContent = equipment.Modelo || '';
        row.insertCell().textContent = equipment.Fabricante || '';
        row.insertCell().textContent = equipment.Setor || '';
        // Usar 'Nº Série Original' para exibição do Nº Série
        row.insertCell().textContent = equipment['Nº Série Original'] || equipment['Nº Série'] || '';
        // Voltar a usar o 'Patrimônio' normalizado para exibição do Patrimônio
        row.insertCell().textContent = equipment.Patrimônio || '';

        const statusCell = row.insertCell();
        statusCell.textContent = equipment.calibrationStatus || 'Desconhecido';
        if (equipment.calibrations && equipment.calibrations.length > 0 && equipment.calibrationStatus !== 'Não Calibrado') {
             statusCell.title = equipment.calibrations.map(cal => `Data Cal: ${cal['DATA CAL'] || 'N/A'}, Vencimento: ${cal['DATA VAL'] || 'N/A'}, Tipo: ${cal['TIPO SERVICO'] || 'N/A'}, Origem: ${cal._source || 'N/A'}`).join('\n');
        }

        const vencimentoCell = row.insertCell();
        vencimentoCell.textContent = equipment.nextCalibrationDate || 'N/A';

        const maintenanceCell = row.insertCell();
        maintenanceCell.textContent = equipment.maintenanceStatus || '';
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
