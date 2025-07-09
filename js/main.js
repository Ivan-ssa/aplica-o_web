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


function populateCalibrationStatusFilter(rawCalibrationsData) {
    // Esta função usa o seu código original e está correta
    const filterElement = calibrationStatusFilter; 
    filterElement.innerHTML = '<option value="">Todos os Status</option>';
    const fixedOptions = [
        { value: 'Calibrado (Consolidado)', text: 'Calibrado (Consolidado)' },
        { value: 'Calibrado (Total)', text: 'Calibrado (Total)' }, 
        { value: 'Divergência (Todos Fornecedores)', text: 'Divergência (Todos Fornecedores)' },
        { value: 'Não Calibrado/Não Encontrado (Seu Cadastro)', text: 'Não Calibrado/Não Encontrado (Seu Cadastro)' },
    ];
    fixedOptions.forEach(opt => {
        const option = document.createElement('option');
        option.value = opt.value;
        option.textContent = opt.text;
        filterElement.appendChild(option);
    });
    const uniqueSuppliers = new Set();
    rawCalibrationsData.forEach(item => {
        const fornecedor = String(item.FornecedorConsolidacao || item.Fornecedor || '').trim();
        if (fornecedor) {
            uniqueSuppliers.add(fornecedor); 
        }
    });
    Array.from(uniqueSuppliers).sort().forEach(fornecedor => {
        const optionDivergence = document.createElement('option');
        optionDivergence.value = `Divergência (${fornecedor})`;
        optionDivergence.textContent = `Divergência (${fornecedor})`;
        filterElement.appendChild(optionDivergence);
    });
}


async function handleProcessFile() {
    outputDiv.textContent = 'Processando arquivos...';
    // Garante que o objeto XLSX da biblioteca correta está disponível
    if (typeof XLSX === 'undefined') {
        alert('ERRO CRÍTICO: A biblioteca de leitura de Excel (xlsx.js) não foi carregada. Verifique o arquivo index.html.');
        return;
    }

    const files = excelFileInput.files;

    if (files.length === 0) {
        outputDiv.textContent = 'Por favor, selecione os arquivos Excel (equipamentos, consolidação de calibrações e manutenção).';
        return;
    }

    let equipmentFile = null;
    let consolidatedCalibrationsFile = null; 
    let externalMaintenanceFile = null; 
    let osCaliAbertasFile = null; 

    for (const file of files) {
        const fileNameLower = file.name.toLowerCase();
        if (fileNameLower.includes('equipamentos')) {
            equipmentFile = file;
        } else if (fileNameLower.includes('empresa_cali_vba') || fileNameLower.includes('consolidado')) { 
            consolidatedCalibrationsFile = file;
        } else if (fileNameLower.includes('manu_externa')) { 
            externalMaintenanceFile = file;
        } else if (fileNameLower.includes('os_cali_abertas')) { 
            osCaliAbertasFile = file;
        }
    }

    if (!equipmentFile) {
        outputDiv.textContent = 'Arquivo de equipamentos não encontrado. Por favor, inclua um arquivo com "equipamentos" no nome.';
        return;
    }

    try {
        outputDiv.textContent = `Lendo arquivo de equipamentos: ${equipmentFile.name}...`;
        allEquipments = await readExcelFile(equipmentFile);
        outputDiv.textContent += `\n${allEquipments.length} equipamentos carregados.`;
        
        // Continua a processar os outros ficheiros...
        const mainEquipmentsBySN = new Map();
        const mainEquipmentsByPatrimonio = new Map();
        allEquipments.forEach(eq => {
            const sn = normalizeId(eq.NumeroSerie); 
            const patrimonio = normalizeId(eq.Patrimonio); 
            if (sn) mainEquipmentsBySN.set(sn, eq);
            if (patrimonio) mainEquipmentsByPatrimonio.set(patrimonio, eq);
        });

        window.consolidatedCalibratedMap.clear(); 
        window.consolidatedCalibrationsRawData = []; 

        if (consolidatedCalibrationsFile) {
            outputDiv.textContent += `\nLendo Calibrações Consolidadas...`;
            const consolidatedData = await readExcelFile(consolidatedCalibrationsFile, 'Consolidação'); 
            window.consolidatedCalibrationsRawData = consolidatedData; 
            consolidatedData.forEach(item => {
                const sn = normalizeId(item.NumeroSerieConsolidacao || item.NumeroSerie || item.NºdeSérie || item['Nº de Série'] || item['Número de Série']); 
                const fornecedor = String(item.FornecedorConsolidacao || item.Fornecedor || '').trim(); 
                if (sn && fornecedor) {
                    window.consolidatedCalibratedMap.set(sn, { fornecedor, dataCalibricao: item.DataCalibracaoConsolidada || item['Data de Calibração'] || '' });
                }
            });
            outputDiv.textContent += `\n${window.consolidatedCalibratedMap.size} SNs de calibração consolidados encontrados.`;
        }
        
        window.externalMaintenanceSNs.clear(); 
        if (externalMaintenanceFile) {
            outputDiv.textContent += `\nLendo Manutenção Externa...`;
            const maintenanceData = await readExcelFile(externalMaintenanceFile);
            maintenanceData.forEach(item => {
                const sn = normalizeId(item.NumeroSerie || item['Nº de Série']); 
                if (sn) window.externalMaintenanceSNs.add(sn);
            });
            outputDiv.textContent += `\n${window.externalMaintenanceSNs.size} SNs em manutenção externa encontrados.`;
        }

        window.osRawData = []; 
        if (osCaliAbertasFile) {
            outputDiv.textContent += `\nLendo OS Abertas...`;
            const rawOsData = await readExcelFile(osCaliAbertasFile);
            window.osRawData = rawOsData.filter(os => {
                const tipoManutencao = String(os.TipoDeManutencao || '').trim().toUpperCase(); 
                return tipoManutencao === 'CALIBRAÇÃO' || tipoManutencao === 'SEGURANÇA ELÉTRICA';
            });
            outputDiv.textContent += `\n${window.osRawData.length} OS Abertas encontradas.`;
        }

        outputDiv.textContent = 'Processamento concluído. Renderizando tabelas...';
        applyAllFiltersAndRender(); 
        populateSectorFilter(allEquipments, sectorFilter); 
        populateCalibrationStatusFilter(window.consolidatedCalibrationsRawData); 
        setupHeaderFilters(allEquipments);

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
        initRonda([], rondaTableBody, rondaCountSpan, '', normalizeId);
        toggleSectionVisibility('equipmentSection'); 

    } catch (error) {
        outputDiv.textContent = `Erro ao processar os arquivos: ${error.message}`;
        console.error('Erro ao processar arquivos:', error);
    }
}


function setupHeaderFilters(equipments) {
    // Esta função usa o seu código original e está correta
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
    };

    const originalHeaders = document.querySelectorAll('#equipmentTable thead tr:first-child th');
    originalHeaders.forEach(th => {
        const filterCell = document.createElement('th');
        const headerText = th.textContent.trim();
        if (!headerText) return;

        const columnInfo = headerFilterMap[headerText];
        if (columnInfo) {
            if (columnInfo.type === 'text') {
                const filterInput = document.createElement('input');
                filterInput.type = 'text';
                filterInput.placeholder = `Filtrar...`;
                filterInput.dataset.property = columnInfo.prop;
                filterInput.addEventListener('keyup', applyAllFiltersAndRender);
                filterCell.appendChild(filterInput);
            } else if (columnInfo.type === 'select_multiple') {
                const filterButton = document.createElement('div');
                filterButton.className = 'filter-button';
                filterButton.textContent = `Filtrar`; 
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
                        value = calibInfo ? 'Calibrado (Consolidado)' : (String(eq.StatusCalibacao || '').toLowerCase().includes('não calibrado') || String(eq.StatusCalibacao || '').trim() === '' ? 'Não Calibrado/Não Encontrado (Seu Cadastro)' : 'Calibrado (Total)');
                    } else { value = eq[columnInfo.prop]; }
                    if (value && String(value).trim() !== '') {
                        uniqueValues.add(String(value).trim());
                    }
                });
                const populateCheckboxes = (searchTerm = '') => {
                    optionsContainer.innerHTML = ''; 
                    const filteredValues = Array.from(uniqueValues).filter(val => String(val).toLowerCase().includes(searchTerm.toLowerCase())).sort();
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
                searchPopupInput.addEventListener('keyup', (event) => populateCheckboxes(event.target.value));
                filterButton.addEventListener('click', (event) => {
                    document.querySelectorAll('.filter-popup.active').forEach(popup => {
                        if (popup !== filterPopup) popup.classList.remove('active');
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
        }
        headerFiltersRow.appendChild(filterCell);
    });
}


function applyAllFiltersAndRender() {
    // Esta função usa o seu código original e está correta
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
            selectedValues.push(checkbox.value.toLowerCase()); 
        });
        const allOptionsCount = popup.querySelectorAll('input[type="checkbox"]:not(.select-all)').length;
        if (selectedValues.length < allOptionsCount) {
            filters.headerFilters[property] = selectedValues;
        }
    });
    const filteredEquipments = applyFilters(allEquipments, filters, normalizeId); 
    renderTable(filteredEquipments, equipmentTableBody, window.consolidatedCalibratedMap, window.externalMaintenanceSNs); 
    updateEquipmentCount(filteredEquipments.length);
}

// ===================================================================================
// === NOVA FUNÇÃO DE EXPORTAÇÃO USANDO A BIBLIOTECA ExcelJS ===
// ===================================================================================
async function exportWithExcelJS(tableId, fileName) {
    const table = document.getElementById(tableId);
    if (!table) return alert(`Tabela com ID "${tableId}" não encontrada.`);
    
    if (typeof ExcelJS === 'undefined') {
        return alert('ERRO CRÍTICO: A biblioteca de exportação (ExcelJS) não foi carregada. Verifique o arquivo index.html.');
    }
    
    outputDiv.textContent = `Gerando arquivo Excel estilizado...`;

    try {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Dados');

        const headerFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0056B3' } }; 
        const headerFont = { name: 'Calibri', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
        const calibratedFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFB3E6B3' } };
        const notCalibratedFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFCCCC' } };
        const maintenanceFont = { name: 'Calibri', size: 11, color: { argb: 'FFDC3545' }, bold: true, italic: true };
        const defaultFont = { name: 'Calibri', size: 11 };
        const defaultBorder = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };

        const headerHTMLRows = Array.from(table.querySelectorAll('thead tr'));
        const headerData = [];
        headerHTMLRows.forEach(tr => {
            if (tr.id === 'headerFilters') return; 
            const rowValues = [];
            tr.querySelectorAll('th').forEach(th => rowValues.push(th.textContent));
            headerData.push(rowValues);
        });

        worksheet.columns = headerData[0].map(headerText => ({ header: headerText, key: headerText, width: 20 }));
        
        const headerRow = worksheet.getRow(1);
        headerRow.eachCell(cell => {
            cell.fill = headerFill;
            cell.font = headerFont;
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
            cell.border = defaultBorder;
        });

        const bodyHTMLRows = Array.from(table.querySelectorAll('tbody tr'));
        bodyHTMLRows.forEach(tr => {
            if (tr.querySelector('td')?.colSpan > 1) return; 

            const cellValues = Array.from(tr.querySelectorAll('td')).map(td => td.textContent);
            const addedRow = worksheet.addRow(cellValues);

            addedRow.eachCell(cell => {
                let cellFill = null;
                if (tr.classList.contains('calibrated-dhme')) {
                    cellFill = calibratedFill;
                } else if (tr.classList.contains('not-calibrated')) {
                    cellFill = notCalibratedFill;
                }
                if (cellFill) cell.fill = cellFill;

                if (tr.classList.contains('in-external-maintenance')) {
                    cell.font = maintenanceFont;
                } else {
                    cell.font = defaultFont;
                }
                cell.border = defaultBorder;
            });
        });
        
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = `${fileName}_${new Date().toISOString().slice(0, 10)}.xlsx`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(link.href);
        outputDiv.textContent = `Arquivo gerado com sucesso!`;

    } catch (error) {
        console.error("Erro ao gerar arquivo com ExcelJS:", error);
        outputDiv.textContent = `Erro ao gerar arquivo: ${error.message}`;
    }
}


// --- EVENT LISTENERS ---
processButton.addEventListener('click', handleProcessFile);
sectorFilter.addEventListener('change', applyAllFiltersAndRender); 
calibrationStatusFilter.addEventListener('change', applyAllFiltersAndRender); 
searchInput.addEventListener('keyup', applyAllFiltersAndRender);
maintenanceFilter.addEventListener('change', applyAllFiltersAndRender); 

exportButton.addEventListener('click', () => {
    exportWithExcelJS('equipmentTable', 'equipamentos_filtrados');
});

exportOsButton.addEventListener('click', () => {
    exportWithExcelJS('osTable', 'os_abertas_filtradas');
});

showEquipmentButton.addEventListener('click', () => toggleSectionVisibility('equipmentSection'));
showOsButton.addEventListener('click', () => toggleSectionVisibility('osSection'));
showRondaButton.addEventListener('click', () => toggleSectionVisibility('rondaSection')); 

startRondaButton.addEventListener('click', () => {
    initRonda(allEquipments, rondaTableBody, rondaCountSpan, rondaSectorSelect.value, normalizeId); 
});

loadRondaButton.addEventListener('click', async () => {
    const file = rondaFileInput.files[0];
    if (file) {
        try {
            outputDiv.textContent = `\nCarregando Ronda...`;
            const existingRondaData = await readExcelFile(file);
            loadExistingRonda(existingRondaData, rondaTableBody, rondaCountSpan); 
            outputDiv.textContent = `\nRonda carregada.`;
        } catch(error) {
            outputDiv.textContent = `\nErro ao carregar ronda: ${error.message}`;
        }
    } else {
        alert(`Selecione um arquivo de Ronda.`);
    }
});

saveRondaButton.addEventListener('click', () => {
    saveRonda(rondaTableBody); 
});

document.addEventListener('DOMContentLoaded', () => {
    toggleSectionVisibility('equipmentSection');
});