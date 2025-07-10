// js/rondaManager.js

let html5QrCodeScanner;
let rondaChecklist = []; // Guarda o checklist de localizações para o setor atual
let currentEditingLocation = null; // Guarda a localização que está a ser editada

/**
 * Preenche o dropdown de seleção de setor na página da ronda.
 */
export function populateRondaSectorSelect(locations, selectElement) {
    selectElement.innerHTML = '<option value="">Selecione um Setor</option>';
    const sectors = new Set(locations.map(loc => loc.Setor));
    Array.from(sectors).sort().forEach(sector => {
        const option = document.createElement('option');
        option.value = sector;
        option.textContent = sector;
        selectElement.appendChild(option);
    });
}

/**
 * Inicia uma ronda guiada para o setor selecionado, criando o checklist de localizações.
 */
export function startGuidedRonda(selectedSector, allLocations) {
    if (!selectedSector) {
        alert("Por favor, selecione um setor para iniciar a ronda.");
        return;
    }

    // Filtra as localizações para o setor selecionado e cria o checklist
    rondaChecklist = allLocations
        .filter(loc => loc.Setor === selectedSector)
        .map(loc => ({
            setor: loc.Setor,
            subLocalizacao: loc.SubLocalizacao,
            status: 'Pendente', // Pode ser 'Pendente', 'Verificado'
            equipamentosEncontrados: [],
            observacoes: ''
        }));
    
    renderRondaChecklist();
}

/**
 * Renderiza o checklist de localizações na tela.
 */
function renderRondaChecklist() {
    const container = document.getElementById('rondaItemsContainer');
    container.innerHTML = '';

    if (rondaChecklist.length === 0) {
        container.innerHTML = '<p class="ronda-placeholder">Nenhuma localização cadastrada para este setor.</p>';
        updateRondaProgress();
        return;
    }

    rondaChecklist.forEach((locationItem, index) => {
        const card = document.createElement('div');
        card.className = 'location-card';
        card.classList.add(locationItem.status === 'Verificado' ? 'status-verified' : 'status-pending');
        card.dataset.index = index;

        let equipamentosHTML = locationItem.equipamentosEncontrados.map(eq => 
            `<div class="found-equipment">
                <span>✅ ${eq.Equipamento} (SN: ${eq.NumeroSerie})</span>
                ${eq.isOutOfSector ? '<span class="warning">⚠️ FORA DO SETOR</span>' : ''}
            </div>`
        ).join('');

        card.innerHTML = `
            <div class="card-header">
                <h3>${locationItem.subLocalizacao}</h3>
                <span class="status-indicator">${locationItem.status}</span>
            </div>
            <div class="card-body">
                <div class="found-equipments-list">${equipamentosHTML || '<p>Nenhum equipamento registado aqui.</p>'}</div>
                <div class="card-actions">
                    <button class="btn-scan" data-index="${index}">Escanear Equipamento</button>
                    <input type="text" class="manual-add-input" placeholder="Ou digite TAG/SN">
                    <button class="btn-manual-add" data-index="${index}">Adicionar</button>
                </div>
            </div>
        `;
        container.appendChild(card);
    });

    // Adiciona os event listeners aos novos botões
    container.querySelectorAll('.btn-scan').forEach(btn => btn.addEventListener('click', handleScanButtonClick));
    container.querySelectorAll('.btn-manual-add').forEach(btn => btn.addEventListener('click', handleManualAddButtonClick));

    updateRondaProgress();
}

/**
 * Atualiza o contador de progresso da ronda.
 */
function updateRondaProgress() {
    const verificados = rondaChecklist.filter(item => item.status === 'Verificado').length;
    const total = rondaChecklist.length;
    document.getElementById('rondaProgress').textContent = `Progresso: ${verificados} / ${total} localizações verificadas`;
}

/**
 * Lida com o clique no botão "Escanear Equipamento" de um cartão de localização.
 */
function handleScanButtonClick(event) {
    const locationIndex = event.target.dataset.index;
    currentEditingLocation = rondaChecklist[locationIndex];
    startScanner();
}

/**
 * Lida com o clique no botão "Adicionar" manual de um cartão de localização.
 */
function handleManualAddButtonClick(event) {
    const locationIndex = event.target.dataset.index;
    const input = event.target.previousElementSibling;
    const id = input.value;

    if (!id.trim()) {
        alert("Por favor, digite uma TAG ou SN.");
        return;
    }
    
    currentEditingLocation = rondaChecklist[locationIndex];
    addEquipmentToLocation(id);
    input.value = ''; // Limpa o campo
}

/**
 * Adiciona um equipamento à localização que está a ser editada atualmente.
 */
function addEquipmentToLocation(equipmentId) {
    if (!currentEditingLocation) return;
    
    const normalizedId = String(equipmentId).trim().toLowerCase();
    const equipmentFound = window.allEquipments.find(eq => 
        String(eq.NumeroSerie).trim().toLowerCase() === normalizedId || 
        String(eq.TAG).trim().toLowerCase() === normalizedId
    );

    if (!equipmentFound) {
        alert(`Equipamento com ID "${equipmentId}" não encontrado na base de dados.`);
        return;
    }

    if (currentEditingLocation.equipamentosEncontrados.some(eq => eq.NumeroSerie === equipmentFound.NumeroSerie)) {
        alert(`Equipamento "${equipmentFound.Equipamento}" já foi adicionado a esta localização.`);
        return;
    }

    const isOutOfSector = equipmentFound.Setor !== currentEditingLocation.setor;

    currentEditingLocation.equipamentosEncontrados.push({
        ...equipmentFound,
        isOutOfSector: isOutOfSector
    });
    
    currentEditingLocation.status = 'Verificado';

    // Re-renderiza a lista para mostrar as alterações
    renderRondaChecklist();
}


/**
 * Função chamada quando um QR Code é lido com sucesso.
 */
function onScanSuccess(decodedText, decodedResult) {
    stopScanner(); // Para o scanner assim que um código é lido
    addEquipmentToLocation(decodedText);
}

/**
 * Inicia o scanner de QR Code.
 */
export function startScanner() {
    document.getElementById('qrScannerContainer').classList.remove('hidden');
    html5QrCodeScanner = new Html5QrcodeScanner("qr-reader", { fps: 10, qrbox: { width: 250, height: 250 } }, false);
    html5QrCodeScanner.render(onScanSuccess, (error) => {});
}

/**
 * Para o scanner de QR Code.
 */
export function stopScanner() {
    if (html5QrCodeScanner) {
        html5QrCodeScanner.clear().catch(err => {});
        document.getElementById('qrScannerContainer').classList.add('hidden');
    }
}

/**
 * Gera e descarrega o relatório final da ronda.
 */
export function saveRonda() {
    if (rondaChecklist.length === 0) {
        alert("Nenhuma ronda iniciada para gerar relatório.");
        return;
    }

    const reportData = [];
    rondaChecklist.forEach(location => {
        if (location.equipamentosEncontrados.length > 0) {
            location.equipamentosEncontrados.forEach(eq => {
                reportData.push({
                    'Data da Ronda': new Date().toLocaleString('pt-BR'),
                    'Setor Auditado': location.setor,
                    'Localização Verificada': location.subLocalizacao,
                    'TAG Encontrado': eq.TAG,
                    'Equipamento Encontrado': eq.Equipamento,
                    'SN Encontrado': eq.NumeroSerie,
                    'Setor Oficial do Equipamento': eq.Setor,
                    'Status da Divergência': eq.isOutOfSector ? 'FORA DO SETOR' : 'OK',
                });
            });
        } else {
            // Adiciona uma linha para localizações onde nada foi encontrado
            reportData.push({
                'Data da Ronda': new Date().toLocaleString('pt-BR'),
                'Setor Auditado': location.setor,
                'Localização Verificada': location.subLocalizacao,
                'TAG Encontrado': 'N/A - Local Vazio',
                'Equipamento Encontrado': '', 'SN Encontrado': '', 'Setor Oficial do Equipamento': '',
                'Status da Divergência': '',
            });
        }
    });

    const ws = XLSX.utils.json_to_sheet(reportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Relatorio_Ronda_Guiada");
    XLSX.writeFile(wb, `Relatorio_Ronda_${new Date().toISOString().slice(0,10)}.xlsx`);
}