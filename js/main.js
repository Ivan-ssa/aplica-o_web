// js/main.js

import { readExcelFile } from './excelReader.js';
import { renderTable, populateSectorFilter, updateEquipmentCount } from './tableRenderer.js';
import { applyFilters } from './filterLogic.js';
import { renderOsTable } from './osRenderer.js'; 
import { initRonda, loadExistingRonda, saveRonda, populateRondaSectorSelect } from './rondaManager.js'; 


// === FUNÇÃO DE NORMALIZAÇÃO DE NÚMERO DE SÉRIE / PATRIMÔNIO ===
function normalizeId(id) {
    if (id === null || id === undefined) {
        return '';
    }
    let strId = String(id).trim(); 

    if (/^\d+$/.test(strId)) { 
        return String(parseInt(strId, 10)); 
    }
    return strId.toLowerCase(); 
}
// =============================================================

let allEquipments = [];
window.consolidatedCalibratedMap = new Map(); 
window.consolidatedCalibrationsRawData = []; 
window.externalMaintenanceSNs = new Set(); 
window.osRawData = []; 
window.rondaData = []; 


// Referências aos elementos do DOM
const excelFileInput = document.getElementById('excelFileInput');
const processButton = document.getElementById('processButton');
const outputDiv = document.getElementById('output');
const equipmentTableBody = document.querySelector('#equipmentTable tbody');
const osTableBody = document.querySelector('#osTable tbody'); 

const sectorFilter = document.getElementById('sectorFilter'); 
const calibrationStatusFilter = document.getElementById('calibrationStatusFilter');
const searchInput = document.getElementById('searchInput'); 
const maintenanceFilter = document.getElementById('maintenanceFilter'); 
const exportButton = document.getElementById('exportButton');
const exportOsButton = document.getElementById('exportOsButton'); 

const headerFiltersRow = document.getElementById('headerFilters'); 

// Botões e Seções de alternância de visualização
const showEquipmentButton = document.getElementById('showEquipmentButton');
const showOsButton = document.getElementById('showOsButton');
const showRondaButton = document.getElementById('showRondaButton'); 
const equipmentSection = document.getElementById('equipmentSection');
const osSection = document.getElementById('osSection');
const rondaSection = document.getElementById('rondaSection'); 

// Elementos da seção de Ronda
const rondaSectorSelect = document.getElementById('rondaSectorSelect');
const startRondaButton = document.getElementById('startRondaButton');
const rondaFileInput = document.getElementById('rondaFileInput');
const loadRondaButton = document.getElementById('loadRondaButton');
const saveRondaButton = document.getElementById('saveRondaButton');
const rondaTableBody = document.querySelector('#rondaTable tbody');
const rondaCountSpan = document.getElementById('rondaCount');


/**
 * Alterna a visibilidade das seções entre Equipamentos, OS e Ronda.
 * @param {string} sectionToShowId - O ID da seção a ser mostrada ('equipmentSection', 'osSection', 'rondaSection').
 */
function toggleSectionVisibility(sectionToShowId) {
    if (equipmentSection) equipmentSection.classList.add('hidden');
    if (osSection) osSection.classList.add('hidden');
    if (rondaSection) rondaSection.classList.add('hidden'); 

    document.querySelectorAll('.toggle-section-button').forEach(button => {
        button.classList.remove('active');
    });

    if (sectionToShowId === 'equipmentSection' && equipmentSection) {
        equipmentSection.classList.remove('hidden');
        if (showEquipmentButton) showEquipmentButton.classList.add('active');
    } else if (sectionToShowId === 'osSection' && osSection) {
        osSection.classList.remove('hidden');
        if (showOsButton) showOsButton.classList.add('active');
    } else if (sectionToShowId === 'rondaSection' && rondaSection) { 
        rondaSection.classList.remove('hidden');
        if (showRondaButton) showRondaButton.classList.add('active');
    }
}


/**
 * Popula o dropdown de Status de Calibração (global) dinamicamente.
 * @param {Array<Object>} rawCalibrationsData - Dados brutos da consolidação para extrair fornecedores.
 * @param {HTMLElement} calibrationStatusFilterElement - O elemento <select> do filtro.
 */
function populateCalibrationStatusFilter(rawCalibrationsData) {
    const filterElement = calibrationStatusFilter; 

    filterElement.innerHTML = '<option value="">Todos os Status</option>';

    // Adiciona a nova opção "Calibrado (Consolidado)"
    const consolidatedCalibratedOption = document.createElement('option');
    consolidatedCalibratedOption.value = 'Calibrado (Consolidado)';
    consolidatedCalibratedOption.textContent = 'Calibrado (Consolidado)';
    filterElement.appendChild(consolidatedCalibratedOption);

    // As opções de divergência por fornecedor ainda podem ser úteis
    const uniqueSuppliers = new Set();
    rawCalibrationsData.forEach(item => {
        const fornecedor = String(item.FornecedorConsolidacao || item.Fornecedor || '').trim();
        if (fornecedor) {
            uniqueSuppliers.add(fornecedor); 
        }
    });

    Array.from(uniqueSuppliers).sort().forEach(fornecedor => {
        // As opções de calibração por fornecedor específico serão removidas em breve,
        // mas mantidas para o filtro de Divergência por enquanto
        const optionDivergence = document.createElement('option');
        optionDivergence.value = `Divergência (${fornecedor})`;
        optionDivergence.textContent = `Divergência (${fornecedor})`;
        filterElement.appendChild(optionDivergence);
    });

    // Opções fixas que permanecem
    const fixedOptions = [
        { value: 'Calibrado (Total)', text: 'Calibrado (Total)' }, 
        { value: 'Divergência (Todos Fornecedores)', text: 'Divergência (Todos Fornecedores)' },
        { value: 'Não Calibrado/Não Encontrado (Seu Cadastro)', text: 'Não Calibrado/Não Encontrado (Seu Cadastro)' },
        { value: 'Não Cadastrado', text: 'Não Cadastrado' }, 
        { value: 'Não Calibrado', text: 'Não Calibrado' } 
    ];

    fixedOptions.forEach(opt => {
        const option = document.createElement('option');
        option.value = opt.value;
        option.textContent = opt.text;
        filterElement.appendChild(option);
    });
}


/**
 * Lida com o processamento dos arquivos Excel selecionados pelo usuário.
 */
async function handleProcessFile() {
    outputDiv.textContent = 'Processando arquivos...';
    const files = excelFileInput.files;

    if (files.length === 0) {
        outputDiv.textContent = 'Por favor, selecione os arquivos Excel (equipamentos, consolidação de calibrações e manutenção).';
        return;
    }

    let equipmentFile = null;
    let consolidatedCalibrationsFile = null; 
    let externalMaintenanceFile = null; 
    let osCaliAbertasFile = null; 
    let rondaExistingFile = null; 

    for (const file of files) {
        const fileNameLower = file.name.toLowerCase();
        if (fileNameLower.includes('equipamentos') && !equipmentFile) {
            equipmentFile = file;
        } else if (fileNameLower.includes('empresa_cali_vba') && !consolidatedCalibrationsFile) { 
            consolidatedCalibrationsFile = file;
        } else if (fileNameLower.includes('manu_externa') && !externalMaintenanceFile) { 
            externalMaintenanceFile = file;
        } else if (fileNameLower.includes('os_cali_abertas') && !osCaliAbertasFile) { 
            osCaliAbertasFile = file;
        } else if (fileNameLower.includes('ronda_existente') && !rondaExistingFile) { 
            rondaExistingFile = file;
        }
    }

    if (!equipmentFile) {
        outputDiv.textContent = 'Arquivo de equipamentos não encontrado. Por favor, inclua um arquivo com "equipamentos" no nome.';
        return;
    }

    try {
        outputDiv.textContent += `\nLendo arquivo de equipamentos: ${equipmentFile.name}...`;
        allEquipments = await readExcelFile(equipmentFile);

        if (allEquipments.length === 0) {
            outputDiv.textContent += `\nNenhum dado encontrado no arquivo de equipamentos "${equipmentFile.name}".`;
            renderTable([], equipmentTableBody, window.consolidatedCalibratedMap, window.externalMaintenanceSNs); 
            populateSectorFilter([], sectorFilter);
            updateEquipmentCount(0);
            renderOsTable([], osTableBody, new Map(), new Map(), new Map(), new Set(), normalizeId); 
            initRonda([], rondaTableBody, rondaCountSpan, '', normalizeId); 
            return;
        }
        outputDiv.textContent += `\n${allEquipments.length} equipamentos carregados.`;

        const mainEquipmentsBySN = new Map();
        const mainEquipmentsByPatrimonio = new Map();
        allEquipments.forEach(eq => {
            const sn = normalizeId(eq.NumeroSerie); 
            const patrimonio = normalizeId(eq.Patrimonio); 
            if (sn) mainEquipmentsBySN.set(sn, eq);
            if (patrimonio) mainEquipmentsByPatrimonio.set(patrimonio, eq);
        });

        // Processa o arquivo de calibrações consolidadas
        window.consolidatedCalibratedMap.clear(); 
        window.consolidatedCalibrationsRawData = []; 

        if (consolidatedCalibrationsFile) {
            outputDiv.textContent += `\nLendo arquivo de Calibrações Consolidadas: ${consolidatedCalibrationsFile.name}...`;
            const consolidatedData = await readExcelFile(consolidatedCalibrationsFile, 'Consolidação'); 
            window.consolidatedCalibrationsRawData = consolidatedData; 

            if (consolidatedData.length > 0) {
                consolidatedData.forEach(item => {
                    const sn = normalizeId(item.NumeroSerieConsolidacao || item.NumeroSerie || item.NºdeSérie || item['Nº de Série'] || item['Número de Série']); 
                    const fornecedor = String(item.FornecedorConsolidacao || item.Fornecedor || '').trim(); 
                    const dataCalibracao = item.DataCalibracaoConsolidada || item['Data de Calibração'] || ''; 

                    if (sn && fornecedor) {
                        window.consolidatedCalibratedMap.set(sn, { fornecedor: fornecedor, dataCalibacao: dataCalibracao });
                    }
                });
                outputDiv.textContent += `\n${window.consolidatedCalibratedMap.size} SNs de calibração consolidados encontrados.`;
            } else {
                outputDiv.textContent += `\nNenhum dado encontrado no arquivo de Calibrações Consolidadas "${consolidatedCalibrationsFile.name}".`;
            }
        } else {
            outputDiv.textContent += `\nArquivo de Calibrações Consolidadas não selecionado.`;
        }

        // Processa o arquivo de Manutenção Externa
        window.externalMaintenanceSNs.clear(); 
        if (externalMaintenanceFile) {
            outputDiv.textContent += `\nLendo arquivo Manutenção Externa: ${externalMaintenanceFile.name}...`;
            const maintenanceData = await readExcelFile(externalMaintenanceFile);
            if (maintenanceData.length > 0) {
                maintenanceData.forEach(item => {
                    const sn = normalizeId(item.NumeroSerie); 
                    if (sn) {
                        window.externalMaintenanceSNs.add(sn);
                    }
                });
                outputDiv.textContent += `\n${window.externalMaintenanceSNs.size} SNs em manutenção externa encontrados.`;
            } else {
                outputDiv.textContent += `\nNenhum dado encontrado no arquivo Manutenção Externa "${externalMaintenanceFile.name}".`;
            }
        } else {
            outputDiv.textContent += `\nArquivo Manutenção Externa não selecionado.`;
        }

        // Processa o arquivo de Ordens de Serviço (OS)
        window.osRawData = []; 
        if (osCaliAbertasFile) {
            outputDiv.textContent += `\nLendo arquivo de OS Abertas: ${osCaliAbertasFile.name}...`;
            const rawOsData = await readExcelFile(osCaliAbertasFile);

            window.osRawData = rawOsData.filter(os => {
                const tipoManutencao = String(os.TipoDeManutencao || '').trim().toUpperCase(); 
                return tipoManutencao === 'CALIBRAÇÃO' || tipoManutencao === 'SEGURANÇA ELÉTRICA';
            });

            if (window.osRawData.length > 0) {
                outputDiv.textContent += `\n${window.osRawData.length} OS Abertas (filtradas por tipo) encontradas.`;
            } else {
                outputDiv.textContent += `\nNenhuma OS Aberta (filtrada por tipo) encontrada no arquivo "${osCaliAbertasFile.name}".`;
            }
        } else {
            outputDiv.textContent += `\nArquivo de OS Abertas não selecionado.`;
        }

        outputDiv.textContent += '\nProcessamento concluído. Renderizando tabelas...';
        applyAllFiltersAndRender(); 
        populateSectorFilter(allEquipments, sectorFilter); 
        populateCalibrationStatusFilter(window.consolidatedCalibrationsRawData); 
        setupHeaderFilters(allEquipments); // Chamada para configurar filtros de cabeçalho

        renderOsTable(
            window.osRawData,
            osTableBody,
            mainEquipmentsBySN, 
            mainEquipmentsByPatrimonio, 
            window.consolidatedCalibratedMap, 
            window.externalMaintenanceSNs, 
            normalizeId 
        );
        populateRondaSectorSelect(allEquipments, rondaSectorSelect);
        if (rondaExistingFile) {
            outputDiv.textContent += `\nCarregando Ronda Existente: ${rondaExistingFile.name}...`;
            const existingRondaData = await readExcelFile(rondaExistingFile);
            loadExistingRonda(existingRondaData, rondaTableBody, rondaCountSpan);
            outputDiv.textContent += `\nRonda Existente carregada.`;
        } else {
            initRonda([], rondaTableBody, rondaCountSpan, '', normalizeId); 
        }

        toggleSectionVisibility('equipmentSection'); 

    } catch (error) {
        outputDiv.textContent = `Erro ao processar os arquivos: ${error.message}`;
        console.error('Erro ao processar arquivos:', error);
    }
}

/**
 * Configura os filtros personalizados nos cabeçalhos da tabela principal.
 */
function setupHeaderFilters(equipments) {
    headerFiltersRow.innerHTML = ''; 

    const headerFilterMap = {
        'TAG': { prop: 'TAG', type: 'text' },
        'Equipamento': { prop: 'Equipamento', type: 'select_multiple' }, 
        'Modelo': { prop: 'Modelo', type: 'select_multiple' },         
        'Fabricante': { prop: 'Fabricante', type: 'select_multiple' },     
        'Setor': { prop: 'Setor', type: 'select_multiple' },             
        'Nº Série': { prop: 'NumeroSerie', type: 'text' },
        'Patrimônio': { prop: 'Patrimonio', type: 'text' },
        'Status Calibração': { prop: 'StatusCalibacao', type: 'select_multiple' }, 
        'Data Vencimento Calibração': { prop: 'DataVencimentoCalibacao', type: 'text' }, 
        'Status Manutenção': { prop: 'StatusManutencao', type: 'text' } 
    };

    const originalHeaders = document.querySelectorAll('#equipmentTable thead tr:first-child th');
    originalHeaders.forEach(th => {
        const filterCell = document.createElement('th');
        const headerText = th.textContent.trim();
        const columnInfo = headerFilterMap[headerText];

        if (columnInfo) {
            if (columnInfo.type === 'text') {
                const filterInput = document.createElement('input');
                filterInput.type = 'text';
                filterInput.placeholder = `Filtrar ${headerText}...`;
                filterInput.dataset.property = columnInfo.prop;
                filterInput.addEventListener('keyup', applyAllFiltersAndRender);
                filterInput.addEventListener('change', applyAllFiltersAndRender);
                filterCell.appendChild(filterInput);
            } else if (columnInfo.type === 'select_multiple') {
                const filterButton = document.createElement('div');
                filterButton.className = 'filter-button';
                filterButton.textContent = `Filtrar ${headerText}`; 
                filterButton.dataset.property = columnInfo.prop;

                const filterPopup = document.createElement('div');
                filterPopup.className = 'filter-popup';
                filterPopup.dataset.property = columnInfo.prop; 

                const searchPopupInput = document.createElement('input');
                searchPopupInput.type = 'text';
                searchPopupInput.placeholder = 'Buscar...';
                searchPopupInput.className = 'filter-search-input';
                filterPopup.appendChild(searchPopupInput);

                const optionsContainer = document.createElement('div'); 
                optionsContainer.className = 'filter-options-container'; 
                filterPopup.appendChild(optionsContainer);


                const uniqueValues = new Set();
                equipments.forEach(eq => {
                    let value;
                    if (columnInfo.prop === 'StatusCalibacao') {
                        const calibInfo = window.consolidatedCalibratedMap.get(normalizeId(eq.NumeroSerie));
                        if (calibInfo) {
                            value = 'Calibrado (Consolidado)'; // Valor genérico
                        } else {
                            const originalStatusLower = String(eq?.StatusCalibacao || '').toLowerCase();
                            if (originalStatusLower.includes('não calibrado') || originalStatusLower.includes('não cadastrado')) {
                                value = 'Não Calibrado/Não Encontrado (Seu Cadastro)';
                            } else if (originalStatusLower.includes('calibrado (total)')) {
                                value = 'Calibrado (Total)';
                            } else {
                                value = 'Não Calibrado/Não Encontrado (Seu Cadastro)';
                            }
                        }
                    } else if (columnInfo.prop === 'StatusManutencao') { 
                         const equipmentSN = normalizeId(eq?.NumeroSerie); 
                         if (window.externalMaintenanceSNs.has(equipmentSN)) {
                             value = 'Em Manutenção Externa';
                         } else {
                             value = String(eq?.StatusManutencao || ''); 
                         }
                    } else if (columnInfo.prop === 'NumeroSerie' || columnInfo.prop === 'Patrimonio') { 
                        value = normalizeId(eq[columnInfo.prop]); 
                    }
                    else {
                        value = eq[columnInfo.prop];
                    }

                    if (value && String(value).trim() !== '') {
                        uniqueValues.add(String(value).trim());
                    }
                });

                const populateCheckboxes = (searchTerm = '') => {
                    optionsContainer.innerHTML = ''; 
                    const filteredValues = Array.from(uniqueValues).filter(val => 
                        String(val).toLowerCase().includes(searchTerm.toLowerCase())
                    ).sort();

                    const selectAllLabel = document.createElement('label');
                    selectAllLabel.className = 'select-all-label'; 
                    const selectAllCheckbox = document.createElement('input');
                    selectAllCheckbox.type = 'checkbox';
                    selectAllCheckbox.className = 'select-all';
                    selectAllCheckbox.checked = true; 

                    selectAllLabel.appendChild(selectAllCheckbox);
                    selectAllLabel.appendChild(document.createTextNode('(Selecionar Todos)'));
                    optionsContainer.appendChild(selectAllLabel);

                    selectAllCheckbox.addEventListener('change', () => {
                        const allIndividualCheckboxes = optionsContainer.querySelectorAll('input[type="checkbox"]:not(.select-all)');
                        allIndividualCheckboxes.forEach(cb => cb.checked = selectAllCheckbox.checked);
                        applyAllFiltersAndRender();
                    });

                    filteredValues.forEach(value => {
                        const label = document.createElement('label');
                        const checkbox = document.createElement('input');
                        checkbox.type = 'checkbox';
                        checkbox.value = value; 
                        checkbox.checked = true; 
                        checkbox.addEventListener('change', () => {
                            const allIndividualCheckboxes = optionsContainer.querySelectorAll('input[type="checkbox"]:not(.select-all)');
                            selectAllCheckbox.checked = Array.from(allIndividualCheckboxes).every(cb => cb.checked);
                            applyAllFiltersAndRender();
                        });

                        label.appendChild(checkbox);
                        label.appendChild(document.createTextNode(value));
                        optionsContainer.appendChild(label);
                    });
                };

                populateCheckboxes(); 

                searchPopupInput.addEventListener('keyup', (event) => {
                    populateCheckboxes(event.target.value);
                });

                filterButton.addEventListener('click', (event) => {
                    document.querySelectorAll('.filter-popup.active').forEach(popup => {
                        if (popup !== filterPopup) {
                            popup.classList.remove('active');
                        }
                    });
                    filterPopup.classList.toggle('active'); 
                    event.stopPropagation(); 
                });

                document.addEventListener('click', (event) => {
                    if (!filterPopup.contains(event.target) && !filterButton.contains(event.target)) {
                        filterPopup.classList.remove('active');
                    }
                });

                filterCell.appendChild(filterButton);
                filterCell.appendChild(filterPopup);
            }
        } else {
            filterCell.innerHTML = '&nbsp;'; 
        }
        headerFiltersRow.appendChild(filterCell);
    });
}

/**
 * Aplica os filtros atuais (globais e de cabeçalho) e renderiza a tabela.
 */
function applyAllFiltersAndRender() {
    const filters = {
        sector: sectorFilter.value, 
        calibrationStatus: calibrationStatusFilter.value,
        search: normalizeId(searchInput.value), 
        maintenance: maintenanceFilter.value,
        headerFilters: {} 
    };

    document.querySelectorAll('#headerFilters input[type="text"]').forEach(input => {
        if (input.value.trim() !== '') {
            filters.headerFilters[input.dataset.property] = normalizeId(input.value); 
        }
    });

    document.querySelectorAll('#headerFilters .filter-popup').forEach(popup => {
        const property = popup.dataset.property;
        const selectedValues = [];
        popup.querySelectorAll('input[type="checkbox"]:checked:not(.select-all)').forEach(checkbox => {
            selectedValues.push(checkbox.value); 
        });

        const allOptionsCount = popup.querySelectorAll('input[type="checkbox"]:not(.select-all)').length;

        if (selectedValues.length > 0 && selectedValues.length < allOptionsCount) {
            filters.headerFilters[property] = selectedValues;
        } else if (selectedValues.length === 0 && allOptionsCount > 0 && !popup.querySelector('.select-all').checked) {
            filters.headerFilters[property] = []; 
        }
    });

    const filteredEquipments = applyFilters(allEquipments, filters, normalizeId); 
    renderTable(filteredEquipments, equipmentTableBody, window.consolidatedCalibratedMap, window.externalMaintenanceSNs); 
    updateEquipmentCount(filteredEquipments.length);
}

/**
 * Exporta a tabela exibida atualmente para um arquivo Excel.
 */
function exportTableToExcel() {
    const table = document.getElementById('equipmentTable');
    const headerFiltersRowElement = document.getElementById('headerFilters');
    headerFiltersRowElement.style.display = 'none';

    const ws = XLSX.utils.table_to_sheet(table);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Equipamentos Filtrados");
    XLSX.writeFile(wb, "equipamentos_filtrados.xlsx");

    headerFiltersRowElement.style.display = '';
}

// Event Listeners para os elementos de interação do usuário
processButton.addEventListener('click', handleProcessFile);
sectorFilter.addEventListener('change', applyAllFiltersAndRender); 
calibrationStatusFilter.addEventListener('change', applyAllFiltersAndRender); 
searchInput.addEventListener('keyup', applyAllFiltersAndRender);
maintenanceFilter.addEventListener('change', applyAllFiltersAndRender); 
exportButton.addEventListener('click', exportTableToExcel);

exportOsButton.addEventListener('click', () => {
    const osTable = document.getElementById('osTable');
    const rows = osTable.querySelectorAll('tr');
    const data = [];

    // --- 1. Definir os Estilos ---
    // Estilo para o cabeçalho
    const headerStyle = {
        font: { bold: true, color: { rgb: "FFFFFFFF" } }, // Fonte branca e em negrito
        fill: { fgColor: { rgb: "FF4F81BD" } }  // Fundo azul escuro
    };
    // Estilo para linhas "Calibrado"
    const calibratedStyle = {
        fill: { fgColor: { rgb: "FFB3E6B3" } } // Verde claro
    };
    // Estilo para linhas "Não Calibrado"
    const notCalibratedStyle = {
        fill: { fgColor: { rgb: "FFFFCCCC" } } // Vermelho claro
    };
    // Estilo para fonte em "Manutenção Externa" (pode ser combinado com os outros)
    const maintenanceFontStyle = {
        font: { color: { rgb: "FFDC3545" }, bold: true, italic: true } // Fonte vermelha, negrito, itálico
    };

    // --- 2. Processar Cabeçalho ---
    const headerRow = [];
    osTable.querySelectorAll('thead tr th').forEach(th => {
        headerRow.push({ v: th.textContent, s: headerStyle });
    });
    data.push(headerRow);

    // --- 3. Processar Linhas do Corpo da Tabela ---
    osTable.querySelectorAll('tbody tr').forEach(tr => {
        const rowData = [];
        // Checa se a linha tem dados ou é a mensagem de "nenhuma OS encontrada"
        if (tr.querySelector('td').colSpan > 1) {
            return; // Pula a linha de mensagem
        }

        // Determina o estilo base da linha (cor de fundo)
        let baseStyle = {};
        if (tr.classList.contains('calibrated-dhme')) {
            baseStyle = calibratedStyle;
        } else if (tr.classList.contains('not-calibrated')) {
            baseStyle = notCalibratedStyle;
        }
        
        tr.querySelectorAll('td').forEach(td => {
            let cellStyle = { ...baseStyle }; // Começa com o estilo de fundo

            // Se a linha também estiver em manutenção, mescla o estilo da fonte
            if (tr.classList.contains('in-external-maintenance')) {
                // Combina o estilo de fonte com qualquer preenchimento que já exista
                cellStyle.font = { ...maintenanceFontStyle.font };
            }
            
            rowData.push({ v: td.textContent, s: cellStyle });
        });
        data.push(rowData);
    });

    // --- 4. Criar a Planilha e o Arquivo ---
    const ws = XLSX.utils.aoa_to_sheet(data);

    // Definir largura das colunas para melhor visualização
    ws['!cols'] = [
        { wch: 10 }, // OS
        { wch: 15 }, // Patrimônio
        { wch: 20 }, // Nº de Série
        { wch: 30 }, // Equipamento
        { wch: 25 }, // Modelo
        { wch: 25 }, // Fabricante
        { wch: 20 }  // Setor
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "OS Abertas Filtradas");
    XLSX.writeFile(wb, `os_abertas_filtradas_${new Date().toISOString().slice(0, 10)}.xlsx`);
});
showEquipmentButton.addEventListener('click', () => toggleSectionVisibility('equipmentSection'));
showOsButton.addEventListener('click', () => toggleSectionVisibility('osSection'));
showRondaButton.addEventListener('click', () => toggleSectionVisibility('rondaSection')); 

// Event Listeners para a funcionalidade de Ronda (delegando para o rondaManager)
startRondaButton.addEventListener('click', () => {
    initRonda(allEquipments, rondaTableBody, rondaCountSpan, rondaSectorSelect.value, normalizeId); 
});

loadRondaButton.addEventListener('click', async () => {
    const file = rondaFileInput.files[0];
    if (file) {
        outputDiv.textContent = `\nCarregando Ronda Existente: ${file.name}...`;
        const existingRondaData = await readExcelFile(file);
        loadRonda(existingRondaData, rondaTableBody, rondaCountSpan); 
        outputDiv.textContent += `\nRonda Existente carregada.`;
    } else {
        outputDiv.textContent = `\nPor favor, selecione um arquivo de Ronda Existente.`;
    }
});

saveRondaButton.addEventListener('click', () => {
    saveRonda(rondaTableBody); 
});