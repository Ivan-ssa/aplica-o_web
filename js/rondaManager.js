// js/rondaManager.js

// Variável para guardar a instância do scanner e os dados da ronda atual
let html5QrCodeScanner;
window.rondaData = []; // Começa sempre vazio

/**
 * Adiciona um cartão de item de ronda à interface.
 * @param {Object} item - O objeto de dados do equipamento escaneado.
 */
function renderRondaCard(item) {
    const container = document.getElementById('rondaItemsContainer');
    
    // Remove o texto inicial "Escaneie um QR Code..." se for o primeiro item
    const placeholder = container.querySelector('.ronda-placeholder');
    if (placeholder) {
        placeholder.remove();
    }

    // Cria o elemento do cartão
    const card = document.createElement('div');
    card.className = 'ronda-card';
    card.dataset.sn = item.NumeroSerie; // Identificador único para o cartão

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

    // Adiciona event listeners para atualizar os dados em tempo real
    card.querySelector(`#loc-${item.NumeroSerie}`).addEventListener('change', (e) => updateRondaItem(item.NumeroSerie, 'Localizacao', e.target.value));
    card.querySelector(`#disp-${item.NumeroSerie}`).addEventListener('change', (e) => updateRondaItem(item.NumeroSerie, 'Disponibilidade', e.target.value));
    card.querySelector(`#obs-${item.NumeroSerie}`).addEventListener('change', (e) => updateRondaItem(item.NumeroSerie, 'Observacoes', e.target.value));

    container.prepend(card); // Adiciona o novo cartão no topo da lista
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
 * Função chamada quando um QR Code é lido com sucesso.
 * @param {string} decodedText - O texto lido do QR Code (deve ser o SN ou TAG).
 * @param {Object} decodedResult - O objeto de resultado completo do scanner.
 */
function onScanSuccess(decodedText, decodedResult) {
    console.log(`Código lido: ${decodedText}`);
    html5QrCodeScanner.pause(); // Pausa o scanner para o utilizador preencher os dados

    // Normaliza o ID lido para corresponder à base de dados
    const normalizedId = String(decodedText).trim().toLowerCase();

    // Procura o equipamento na base de dados principal (allEquipments)
    const equipmentFound = window.allEquipments.find(eq => 
        String(eq.NumeroSerie).trim().toLowerCase() === normalizedId || 
        String(eq.TAG).trim().toLowerCase() === normalizedId
    );

    if (!equipmentFound) {
        alert(`Equipamento com ID "${decodedText}" não encontrado na base de dados.`);
        html5QrCodeScanner.resume(); // Retoma o scanner
        return;
    }

    // Verifica se este equipamento já foi escaneado nesta ronda
    if (window.rondaData.some(item => item.NumeroSerie === equipmentFound.NumeroSerie)) {
        alert(`Equipamento "${equipmentFound.Equipamento}" já foi verificado nesta ronda.`);
        html5QrCodeScanner.resume();
        return;
    }

    // Adiciona o novo equipamento à lista da ronda
    const newItem = {
        TAG: equipmentFound.TAG ?? '',
        Equipamento: equipmentFound.Equipamento ?? '',
        Setor: equipmentFound.Setor ?? '', // Setor oficial para referência
        NumeroSerie: equipmentFound.NumeroSerie ?? '',
        Patrimonio: equipmentFound.Patrimonio ?? '',
        Localizacao: '',     // Campo a ser preenchido pelo utilizador
        Disponibilidade: '', // Campo a ser preenchido pelo utilizador
        Observacoes: ''      // Campo a ser preenchido pelo utilizador
    };
    window.rondaData.push(newItem);

    // Renderiza o novo cartão na interface
    renderRondaCard(newItem);
    
    // Atualiza a contagem
    document.getElementById('rondaCount').textContent = `Equipamentos verificados: ${window.rondaData.length}`;

    // Mostra feedback e retoma o scanner após um pequeno atraso
    const outputDiv = document.getElementById('output');
    outputDiv.textContent = `Equipamento "${newItem.Equipamento}" adicionado. Aponte para o próximo QR Code.`;
    setTimeout(() => {
        html5QrCodeScanner.resume();
    }, 1500); // Espera 1.5 segundos antes de reativar
}

/**
 * Inicia o scanner de QR Code.
 */
export function startScanner() {
    // Esconde o botão de iniciar e mostra o container do scanner
    document.getElementById('startRondaScanButton').classList.add('hidden');
    const scannerContainer = document.getElementById('qrScannerContainer');
    scannerContainer.classList.remove('hidden');

    html5QrCodeScanner = new Html5QrcodeScanner(
        "qr-reader", // ID do elemento div no HTML
        { fps: 10, qrbox: { width: 250, height: 250 } }, // Configurações do scanner
        false // verbose = false
    );
    html5QrCodeScanner.render(onScanSuccess, (error) => {
        // Lidar com erros de scan, geralmente ignorado
    });
}

/**
 * Para o scanner de QR Code.
 */
export function stopScanner() {
    if (html5QrCodeScanner) {
        html5QrCodeScanner.clear().then(_ => {
            console.log("Scanner parado com sucesso.");
            document.getElementById('startRondaScanButton').classList.remove('hidden');
            document.getElementById('qrScannerContainer').classList.add('hidden');
        }).catch(error => {
            console.error("Falha ao parar o scanner.", error);
        });
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
        'TAG': item.TAG,
        'Equipamento': item.Equipamento,
        'Setor Oficial': item.Setor,
        'Nº de Série': item.NumeroSerie, 
        'Patrimônio': item.Patrimonio,   
        'Disponibilidade na Ronda': item.Disponibilidade,
        'Localização na Ronda': item.Localizacao,
        'Observações da Ronda': item.Observacoes,
        'Data da Ronda': new Date().toLocaleString('pt-BR')
    }));
    
    // Usa a biblioteca xlsx.js para criar o ficheiro (já que está carregada para a leitura)
    const ws = XLSX.utils.json_to_sheet(dataToExport);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Coleta_Ronda");
    XLSX.writeFile(wb, `Ronda_Coleta_${new Date().toISOString().slice(0,10)}.xlsx`);
    
    // Limpa os dados para a próxima ronda
    window.rondaData = [];
    document.getElementById('rondaItemsContainer').innerHTML = '<p class="ronda-placeholder">Escaneie um QR Code para começar a adicionar equipamentos.</p>';
    document.getElementById('rondaCount').textContent = `Equipamentos verificados: 0`;
    alert("Ronda salva com sucesso! A lista foi limpa para a próxima ronda.");
}

/**
 * Limpa a interface da ronda para um novo início.
 */
export function clearRonda() {
    stopScanner();
    window.rondaData = [];
    const container = document.getElementById('rondaItemsContainer');
    if (container) {
        container.innerHTML = '<p class="ronda-placeholder">Escaneie um QR Code para começar a adicionar equipamentos.</p>';
    }
    const countSpan = document.getElementById('rondaCount');
    if (countSpan) {
        countSpan.textContent = `Equipamentos verificados: 0`;
    }
}