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
        
        // Adiciona classe CSS para destaque baseado no novo status
        if (equipment.calibrationStatus === 'Não Calibrado') {
            row.classList.add('not-calibrated');
        } else if (equipment.calibrationStatus === 'Calibrado') {
            row.classList.add('calibrated');
        } else if (equipment.calibrationStatus === 'Não Cadastrado (DHME)') { // NOVO STATUS
            row.classList.add('divergent-calibrated'); // Nova classe CSS para divergências
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
        // Tooltip para calibrados e também para os "Não Cadastrados (DHME)" que terão detalhes de calibração
        if (equipment.calibrations && equipment.calibrations.length > 0 && equipment.calibrationStatus !== 'Não Calibrado') {
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

// REMOVER ESTA FUNÇÃO se você removeu a tabela do HTML
// export const renderDivergentCalibrationsTable = (divergentData, divergentTableBody) => {
//     divergentTableBody.innerHTML = '';
//     if (!divergentData || divergentData.length === 0) {
//         divergentTableBody.innerHTML = '<tr><td colspan="6">Nenhuma calibração com divergência encontrada.</td></tr>';
//         return;
//     }
//     divergentData.forEach(cal => {
//         const row = divergentTableBody.insertRow();
//         row.insertCell().textContent = cal.SN || '';
//         row.insertCell().textContent = cal.EQUIPAMENTO || '';
//         row.insertCell().textContent = cal.MARCA || '';
//         row.insertCell().textContent = cal.MODELO || '';
//         row.insertCell().textContent = cal['DATA VAL'] || '';
//         row.insertCell().textContent = cal['TIPO SERVIÇO'] || '';
//     });
// };
