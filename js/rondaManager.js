// js/rondaManager.js

/**
 * Inicializa a tabela de Ronda com equipamentos de um setor específico ou vazia.
 * @param {Array<Object>} allEquipments - Todos os equipamentos cadastrados.
 * @param {HTMLElement} rondaTableBody - O tbody da tabela de ronda.
 * @param {HTMLElement} rondaCountSpan - O span para exibir a contagem.
 * @param {string} selectedSector - O setor selecionado para iniciar a ronda.
 * @param {Function} normalizeId - Função para normalizar IDs.
 */
export function initRonda(allEquipments, rondaTableBody, rondaCountSpan, selectedSector = '', normalizeId) {
    rondaTableBody.innerHTML = ''; // Limpa a tabela

    let equipmentsForRonda = [];
    if (selectedSector && allEquipments.length > 0) {
        equipmentsForRonda = allEquipments.filter(eq => 
            String(eq.Setor || '').trim() === selectedSector
        );
    } else {
        equipmentsForRonda = []; 
    }

    // Mapeia para os dados da ronda, garantindo que SN/Patrimônio são normalizados
    window.rondaData = equipmentsForRonda.map(eq => ({
        TAG: eq.TAG ?? '',
        Equipamento: eq.Equipamento ?? '',
        Setor: eq.Setor ?? '',
        NumeroSerie: normalizeId(eq.NumeroSerie), 
        Patrimonio: normalizeId(eq.Patrimonio),   
        Disponibilidade: '', 
        Localizacao: '',     
        Observacoes: ''      
    }));

    renderRondaTable(window.rondaData, rondaTableBody, rondaCountSpan);
}

/**
 * Renderiza os dados da ronda na tabela.
 * @param {Array<Object>} data - Os dados da ronda.
 * @param {HTMLElement} rondaTableBody - O tbody da tabela de ronda.
 * @param {HTMLElement} rondaCountSpan - O span para exibir a contagem.
 */
function renderRondaTable(data, rondaTableBody, rondaCountSpan) {
    rondaTableBody.innerHTML = '';

    if (data.length === 0) {
        const row = rondaTableBody.insertRow();
        const cell = row.insertCell();
        cell.colSpan = 6;
        cell.textContent = 'Nenhum equipamento para ronda no setor selecionado ou ronda não carregada.';
        cell.style.textAlign = 'center';
        rondaCountSpan.textContent = `Total: 0 Equipamentos na Ronda`;
        return;
    }

    data.forEach((item, index) => {
        const row = rondaTableBody.insertRow();
        row.dataset.rowIndex = index; // Armazena o índice para fácil referência ao salvar

        row.insertCell().textContent = item.TAG;
        row.insertCell().textContent = item.Equipamento;
        row.insertCell().textContent = item.Setor;

        const dispCell = row.insertCell();
        const dispSelect = document.createElement('select');
        dispSelect.dataset.property = 'Disponibilidade';
        const dispOptions = ['Disponível', 'Em Uso', 'Em Manutenção', 'Desativado', 'Perdido', 'Outro'];

        const defaultDispOption = document.createElement('option');
        defaultDispOption.value = '';
        defaultDispOption.textContent = 'Selecione...';
        dispSelect.appendChild(defaultDispOption);

        dispOptions.forEach(opt => {
            const option = document.createElement('option');
            option.value = opt;
            option.textContent = opt;
            dispSelect.appendChild(option);
        });
        dispSelect.value = item.Disponibilidade || ''; 
        dispSelect.addEventListener('change', (e) => updateRondaItem(row.dataset.rowIndex, 'Disponibilidade', e.target.value));
        dispCell.appendChild(dispSelect);

        const locCell = row.insertCell();
        const locInput = document.createElement('input');
        locInput.type = 'text';
        locInput.value = item.Localizacao || '';
        locInput.dataset.property = 'Localizacao';
        locInput.addEventListener('change', (e) => updateRondaItem(row.dataset.rowIndex, 'Localizacao', e.target.value));
        locCell.appendChild(locInput);

        const obsCell = row.insertCell();
        const obsInput = document.createElement('input');
        obsInput.type = 'text';
        obsInput.value = item.Observacoes || '';
        obsInput.dataset.property = 'Observacoes';
        obsInput.addEventListener('change', (e) => updateRondaItem(row.dataset.rowIndex, 'Observacoes', e.target.value));
        obsCell.appendChild(obsInput);
    });

    rondaCountSpan.textContent = `Total: ${data.length} Equipamentos na Ronda`;
}

/**
 * Atualiza um item de ronda na memória quando um input/select é alterado.
 * @param {number} index - O índice do item na window.rondaData.
 * @param {string} property - A propriedade a ser atualizada ('Disponibilidade', 'Localizacao', 'Observacoes').
 * @param {*} value - O novo valor.
 */
function updateRondaItem(index, property, value) {
    if (window.rondaData[index]) {
        window.rondaData[index][property] = value;
    }
}

/**
 * Popula o select de setores na seção de Ronda.
 * @param {Array<Object>} allEquipments - Todos os equipamentos para obter os setores.
 * @param {HTMLElement} selectElement - O elemento <select> a ser populado.
 */
export function populateRondaSectorSelect(allEquipments, selectElement) {
    selectElement.innerHTML = '<option value="">Selecione um Setor</option>';
    const sectors = new Set();
    allEquipments.forEach(eq => {
        if (eq.Setor && String(eq.Setor).trim() !== '') {
            sectors.add(String(eq.Setor).trim());
        }
    });
    Array.from(sectors).sort().forEach(sector => {
        const option = document.createElement('option');
        option.value = sector;
        option.textContent = sector;
        selectElement.appendChild(option);
    });
}

/**
 * Carrega dados de uma ronda existente de um arquivo Excel e renderiza na tabela.
 * @param {Array<Object>} existingRondaData - Dados lidos do arquivo de ronda.
 * @param {HTMLElement} rondaTableBody - O tbody da tabela de ronda.
 * @param {HTMLElement} rondaCountSpan - O span para exibir a contagem.
 */
export function loadExistingRonda(existingRondaData, rondaTableBody, rondaCountSpan) {
    window.rondaData = existingRondaData; 
    renderRondaTable(window.rondaData, rondaTableBody, rondaCountSpan);
}

/**
 * Salva os dados da ronda (incluindo inputs do usuário) para um arquivo Excel.
 * @param {HTMLElement} rondaTableBody - O tbody da tabela de ronda para coletar os dados visuais.
 */
export function saveRonda(rondaTableBody) {
    if (window.rondaData.length === 0) {
        alert("Não há dados na tabela de ronda para salvar.");
        return;
    }

    const dataToExport = window.rondaData.map(item => ({
        'TAG': item.TAG,
        'Equipamento': item.Equipamento,
        'Setor': item.Setor,
        'Nº de Série': item.NumeroSerie, 
        'Patrimônio': item.Patrimonio,   
        'Disponibilidade': item.Disponibilidade,
        'Localização': item.Localizacao,
        'Observações': item.Observacoes
    }));

    const ws = XLSX.utils.json_to_sheet(dataToExport);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Ronda_Equipamentos");
    XLSX.writeFile(wb, `Ronda_Equipamentos_${new Date().toISOString().slice(0, 10)}.xlsx`);
}