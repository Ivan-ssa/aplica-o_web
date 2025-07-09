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

    // Mapeia para os dados da ronda
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
        row.dataset.rowIndex = index;

        // Adiciona as células com `data-label` para a interface móvel
        let cell;
        
        cell = row.insertCell();
        cell.textContent = item.TAG;
        cell.dataset.label = 'TAG';

        cell = row.insertCell();
        cell.textContent = item.Equipamento;
        cell.dataset.label = 'Equipamento';

        cell = row.insertCell();
        cell.textContent = item.Setor;
        cell.dataset.label = 'Setor';

        const dispCell = row.insertCell();
        dispCell.dataset.label = 'Disponibilidade';
        const dispSelect = document.createElement('select');
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
        locCell.dataset.label = 'Localização (Sala/Quarto)';
        const locInput = document.createElement('input');
        locInput.type = 'text';
        locInput.value = item.Localizacao || '';
        locInput.placeholder = 'Digite a localização...';
        locInput.addEventListener('change', (e) => updateRondaItem(row.dataset.rowIndex, 'Localizacao', e.target.value));
        locCell.appendChild(locInput);

        const obsCell = row.insertCell();
        obsCell.dataset.label = 'Observações da Ronda';
        const obsInput = document.createElement('input');
        obsInput.type = 'text';
        obsInput.value = item.Observacoes || '';
        obsInput.placeholder = 'Digite as observações...';
        obsInput.addEventListener('change', (e) => updateRondaItem(row.dataset.rowIndex, 'Observacoes', e.target.value));
        obsCell.appendChild(obsInput);
    });

    rondaCountSpan.textContent = `Total: ${data.length} Equipamentos na Ronda`;
}

/**
 * Atualiza um item de ronda na memória quando um input/select é alterado.
 */
function updateRondaItem(index, property, value) {
    if (window.rondaData[index]) {
        window.rondaData[index][property] = value;
    }
}

/**
 * Popula o select de setores na seção de Ronda.
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
 */
export function loadExistingRonda(existingRondaData, rondaTableBody, rondaCountSpan) {
    window.rondaData = existingRondaData; 
    renderRondaTable(window.rondaData, rondaTableBody, rondaCountSpan);
}

/**
 * Salva os dados da ronda para um arquivo Excel.
 */
export function saveRonda() {
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