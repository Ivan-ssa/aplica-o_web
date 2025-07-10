// js/rondaManager.js

let html5QrCodeScanner;
window.rondaData = []; // Começa sempre vazio

/**
 * Adiciona um cartão de item de ronda à interface.
 * @param {Object} item - O objeto de dados do equipamento.
 */
function renderRondaCard(item) {
    const container = document.getElementById('rondaItemsContainer');
    const placeholder = container.querySelector('.ronda-placeholder');
    if (placeholder) placeholder.remove();

    const card = document.createElement('div');
    card.className = 'ronda-card';
    card.dataset.sn = item.NumeroSerie;

    card.innerHTML = `
        <div class="card-header">
            <strong>${item.Equipamento} (${item.TAG})</strong>
            <small>SN: ${item.NumeroSerie} | Setor Oficial: ${item.Setor}</small>
        </div>
        <div class="card-body">
            <div class="card-field">
                <label for="loc-${item.NumeroSerie}">Localização Encontrada:</label>
                <input type="text" id="loc-${item.NumeroSerie}" value="${item.Localizacao}" placeholder="Ex: UTI / Leito 03">
            </div>
            <div class="card-field">
                <label for="disp-${item.NumeroSerie}">Disponibilidade:</label>
                <select id="disp-${item.NumeroSerie}">
                    <option value="" ${item.Disponibilidade === '' ? 'selected' : ''}>Selecione...</option>
                    <option value="Disponível" ${item.Disponibilidade === 'Disponível' ? 'selected' : ''}>Disponível</option>
                    <option value="Em Uso" ${item.Disponibilidade === 'Em Uso' ? 'selected' : ''}>Em Uso</option>
                    <option value="Em Manutenção" ${item.Disponibilidade === 'Em Manutenção' ? 'selected' : ''}>Em Manutenção</option>
                    <option value="Desativado" ${item.Disponibilidade === 'Desativado' ? 'selected' : ''}>Desativado</option>
                    <option value="Perdido" ${item.Disponibilidade === 'Perdido' ? 'selected' : ''}>Perdido</option>
                    <option value="Outro" ${item.Disponibilidade === 'Outro' ? 'selected' : ''}>Outro</option>
                </select>
            </div>
             <div class="card-field">
                <label for="obs-${item.NumeroSerie}">Observações:</label>
                <input type="text" id="obs-${item.NumeroSerie}" value="${item.Observacoes}" placeholder="Qualquer observação relevante">
            </div>
        </div>
    `;

    card.querySelector(`#loc-${item.NumeroSerie}`).addEventListener('change', (e) => updateRondaItem(item.NumeroSerie, 'Localizacao', e.target.value));
    card.querySelector(`#disp-${item.NumeroSerie}`).addEventListener('change', (e) => updateRondaItem(item.NumeroSerie, 'Disponibilidade', e.target.value));
    card.querySelector(`#obs-${item.NumeroSerie}`).addEventListener('change', (e) => updateRondaItem(item.NumeroSerie, 'Observacoes', e.target.value));

    container.prepend(card);
    document.getElementById('rondaCount').textContent = `Equipamentos verificados: ${window.rondaData.length}`;
}

/**
 * Atualiza um item na lista de ronda (window.rondaData).
 */
function updateRondaItem(sn, property, value) {
    const item = window.rondaData.find(d => d.NumeroSerie === sn);
    if (item) {
        item[property] = value;
    }
}

/**
 * Lógica central para adicionar um equipamento à ronda, seja por QR Code ou manual.
 * @param {string} id - O ID (SN, TAG, Patrimônio) a ser procurado.
 * @returns {boolean} - Retorna true se o item foi adicionado com sucesso, false caso contrário.
 */
function addEquipmentToRonda(id) {
    const normalizedId = String(id).trim().toLowerCase();
    if (!normalizedId) return false;

    const equipmentFound = window.allEquipments.find(eq => 
        String(eq.NumeroSerie).trim().toLowerCase() === normalizedId || 
        String(eq.TAG).trim().toLowerCase() === normalizedId ||
        String(eq.Patrimonio).trim().toLowerCase() === normalizedId
    );

    if (!equipmentFound) {
        alert(`Equipamento com ID "${id}" não encontrado na base de dados.`);
        return false;
    }

    if (window.rondaData.some(item => item.NumeroSerie === equipmentFound.NumeroSerie)) {
        alert(`Equipamento "${equipmentFound.Equipamento}" já foi verificado nesta ronda.`);
        return false;
    }

    const newItem = {
        TAG: equipmentFound.TAG ?? '',
        Equipamento: equipmentFound.Equipamento ?? '',
        Setor: equipmentFound.Setor ?? '',
        NumeroSerie: equipmentFound.NumeroSerie ?? '',
        Patrimonio: equipmentFound.Patrimonio ?? '',
        Localizacao: '', Disponibilidade: '', Observacoes: ''
    };
    window.rondaData.push(newItem);
    renderRondaCard(newItem);
    return true;
}

// **NOVA FUNÇÃO para adição manual**
export function addEquipmentToRondaManually(id) {
    const success = addEquipmentToRonda(id);
    if (success) {
        document.getElementById('output').textContent = `Equipamento com ID "${id}" adicionado manualmente.`;
    }
}

/**
 * Callback para quando um QR Code é lido com sucesso.
 */
function onScanSuccess(decodedText, decodedResult) {
    html5QrCodeScanner.pause();
    const success = addEquipmentToRonda(decodedText);
    if (success) {
        document.getElementById('output').textContent = `Equipamento "${decodedText}" adicionado. Aponte para o próximo QR Code.`;
    }
    setTimeout(() => {
        if (html5QrCodeScanner.getState() === Html5QrcodeScannerState.PAUSED) {
            html5QrCodeScanner.resume();
        }
    }, 1500);
}

/**
 * Inicia o scanner de QR Code.
 */
export function startScanner() {
    document.getElementById('rondaControls').classList.add('hidden');
    document.getElementById('qrScannerContainer').classList.remove('hidden');

    html5QrCodeScanner = new Html5QrcodeScanner("qr-reader", { fps: 10, qrbox: { width: 250, height: 250 } }, false);
    html5QrCodeScanner.render(onScanSuccess, (error) => {});
}

/**
 * Para o scanner de QR Code.
 */
export function stopScanner() {
    if (html5QrCodeScanner && html5QrCodeScanner.getState() !== Html5QrcodeScannerState.NOT_STARTED) {
        html5QrCodeScanner.clear().then(_ => {
            document.getElementById('rondaControls').classList.remove('hidden');
            document.getElementById('qrScannerContainer').classList.add('hidden');
        }).catch(error => console.error("Falha ao parar o scanner.", error));
    } else {
        document.getElementById('rondaControls').classList.remove('hidden');
        document.getElementById('qrScannerContainer').classList.add('hidden');
    }
}

/**
 * Salva os dados da ronda atual para um ficheiro Excel.
 */
export function saveRonda() {
    if (window.rondaData.length === 0) {
        alert("Nenhum equipamento foi verificado nesta ronda para salvar.");
        return;
    }
    const dataToExport = window.rondaData.map(item => ({
        'TAG': item.TAG, 'Equipamento': item.Equipamento, 'Setor Oficial': item.Setor,
        'Nº de Série': item.NumeroSerie, 'Patrimônio': item.Patrimonio,   
        'Disponibilidade na Ronda': item.Disponibilidade, 'Localização na Ronda': item.Localizacao,
        'Observações da Ronda': item.Observacoes, 'Data da Ronda': new Date().toLocaleString('pt-BR')
    }));
    const ws = XLSX.utils.json_to_sheet(dataToExport);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Coleta_Ronda");
    XLSX.writeFile(wb, `Ronda_Coleta_${new Date().toISOString().slice(0,10)}.xlsx`);
    
    clearRonda();
    alert("Ronda salva com sucesso! A lista foi limpa para a próxima ronda.");
}

/**
 * Limpa a interface da ronda para um novo início.
 */
export function clearRonda() {
    stopScanner();
    window.rondaData = [];
    const container = document.getElementById('rondaItemsContainer');
    if (container) container.innerHTML = '<p class="ronda-placeholder">Escaneie um QR Code ou digite um ID para começar a ronda.</p>';
    const countSpan = document.getElementById('rondaCount');
    if (countSpan) countSpan.textContent = `Equipamentos verificados: 0`;
}