// js/main.js

import { readExcelFile } from './excelReader.js';
import { renderTable, populateSectorFilter, updateEquipmentCount } from './tableRenderer.js';
import { applyFilters } from './filterLogic.js';
import { renderOsTable } from './osRenderer.js'; 
// Importa as novas funções do rondaManager
import { startScanner, stopScanner, saveRonda, clearRonda } from './rondaManager.js'; 


// === FUNÇÃO DE NORMALIZAÇÃO DE NÚMERO DE SÉRIE / PATRIMÔNIO ===
function normalizeId(id) {
    if (id === null || id === undefined) return '';
    let strId = String(id).trim(); 
    if (/^\d+$/.test(strId)) return String(parseInt(strId, 10)); 
    return strId.toLowerCase(); 
}
// =============================================================

// Variáveis globais
window.allEquipments = [];
window.consolidatedCalibratedMap = new Map(); 
window.consolidatedCalibrationsRawData = []; 
window.externalMaintenanceSNs = new Set(); 
window.osRawData = []; 

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
const showEquipmentButton = document.getElementById('showEquipmentButton');
const showOsButton = document.getElementById('showOsButton');
const showRondaButton = document.getElementById('showRondaButton'); 
const equipmentSection = document.getElementById('equipmentSection');
const osSection = document.getElementById('osSection');
const rondaSection = document.getElementById('rondaSection'); 

// *** NOVOS ELEMENTOS DA RONDA ***
const startRondaScanButton = document.getElementById('startRondaScanButton');
const stopRondaScanButton = document.getElementById('stopRondaScanButton');
const saveRondaButton = document.getElementById('saveRondaButton');


function toggleSectionVisibility(sectionToShowId) {
    if (equipmentSection) equipmentSection.classList.add('hidden');
    if (osSection) osSection.classList.add('hidden');
    if (rondaSection) rondaSection.classList.add('hidden'); 

    document.querySelectorAll('.toggle-section-button').forEach(button => {
        button.classList.remove('active');
    });

    if (sectionToShowId === 'equipmentSection' && equipmentSection) {
        equipmentSection.classList.remove('hidden');
        showEquipmentButton.classList.add('active');
    } else if (sectionToShowId === 'osSection' && osSection) {
        osSection.classList.remove('hidden');
        showOsButton.classList.add('active');
    } else if (sectionToShowId === 'rondaSection' && rondaSection) { 
        rondaSection.classList.remove('hidden');
        showRondaButton.classList.add('active');
        clearRonda(); // Limpa a ronda anterior ao entrar na secção
    }
}


async function handleProcessFile() {
    outputDiv.textContent = 'Processando arquivos...';
    if (typeof XLSX === 'undefined') {
        alert('ERRO CRÍTICO: A biblioteca de leitura de Excel (xlsx.js) não foi carregada.');
        return;
    }
    const files = excelFileInput.files;
    if (files.length === 0) {
        outputDiv.textContent = 'Por favor, selecione os arquivos Excel.';
        return;
    }

    let equipmentFile = null, consolidatedCalibrationsFile = null, externalMaintenanceFile = null, osCaliAbertasFile = null;
    for (const file of files) {
        const fileNameLower = file.name.toLowerCase();
        if (fileNameLower.includes('equipamentos')) equipmentFile = file;
        else if (fileNameLower.includes('empresa_cali_vba') || fileNameLower.includes('consolidado')) consolidatedCalibrationsFile = file;
        else if (fileNameLower.includes('manu_externa')) externalMaintenanceFile = file;
        else if (fileNameLower.includes('os_cali_abertas')) osCaliAbertasFile = file;
    }

    if (!equipmentFile) {
        outputDiv.textContent = 'Arquivo de equipamentos não encontrado.';
        return;
    }

    try {
        outputDiv.textContent = `Lendo arquivo de equipamentos: ${equipmentFile.name}...`;
        window.allEquipments = await readExcelFile(equipmentFile);
        outputDiv.textContent += `\n${window.allEquipments.length} equipamentos carregados.`;
        
        const mainEquipmentsBySN = new Map();
        const mainEquipmentsByPatrimonio = new Map();
        window.allEquipments.forEach(eq => {
            const sn = normalizeId(eq.NumeroSerie); 
            if (sn) mainEquipmentsBySN.set(sn, eq);
            const patrimonio = normalizeId(eq.Patrimonio); 
            if (patrimonio) mainEquipmentsByPatrimonio.set(patrimonio, eq);
        });

        window.consolidatedCalibratedMap.clear(); 
        if (consolidatedCalibrationsFile) {
            outputDiv.textContent += `\nLendo Calibrações Consolidadas...`;
            const consolidatedData = await readExcelFile(consolidatedCalibrationsFile, 'Consolidação'); 
            window.consolidatedCalibrationsRawData = consolidatedData; 
            consolidatedData.forEach(item => {
                const sn = normalizeId(item.NumeroSerieConsolidacao || item.NumeroSerie || item['Nº de Série']);
                if (sn) window.consolidatedCalibratedMap.set(sn, { fornecedor: item.FornecedorConsolidacao || item.Fornecedor, dataCalibricao: item.DataCalibracaoConsolidada || item['Data de Calibração'] });
            });
            outputDiv.textContent += `\n${window.consolidatedCalibratedMap.size} SNs de calibração consolidados.`;
        }
        
        window.externalMaintenanceSNs.clear(); 
        if (externalMaintenanceFile) {
            outputDiv.textContent += `\nLendo Manutenção Externa...`;
            const maintenanceData = await readExcelFile(externalMaintenanceFile);
            maintenanceData.forEach(item => {
                const sn = normalizeId(item.NumeroSerie || item['Nº de Série']);
                if (sn) window.externalMaintenanceSNs.add(sn);
            });
            outputDiv.textContent += `\n${window.externalMaintenanceSNs.size} SNs em manutenção externa.`;
        }

        window.osRawData = []; 
        if (osCaliAbertasFile) {
            outputDiv.textContent += `\nLendo OS Abertas...`;
            const rawOsData = await readExcelFile(osCaliAbertasFile);
            window.osRawData = rawOsData.filter(os => {
                const tipo = String(os.TipoDeManutencao || '').toUpperCase(); 
                return tipo === 'CALIBRAÇÃO' || tipo === 'SEGURANÇA ELÉTRICA';
            });
            outputDiv.textContent += `\n${window.osRawData.length} OS Abertas encontradas.`;
        }

        outputDiv.textContent = 'Processamento concluído. Renderizando tabelas...';
        applyAllFiltersAndRender();
        populateSectorFilter(window.allEquipments, sectorFilter);
        // ... outras funções de renderização
        renderOsTable(window.osRawData, osTableBody, mainEquipmentsBySN, mainEquipmentsByPatrimonio, window.consolidatedCalibratedMap, window.externalMaintenanceSNs, normalizeId);
        toggleSectionVisibility('equipmentSection');

    } catch (error) {
        outputDiv.textContent = `Erro: ${error.message}`;
        console.error(error);
    }
}

function applyAllFiltersAndRender() {
    // Esta função permanece igual, apenas renderiza a tabela de equipamentos
    const filters = {
        sector: sectorFilter.value,
        calibrationStatus: calibrationStatusFilter.value,
        search: normalizeId(searchInput.value),
        maintenance: maintenanceFilter.value
    };
    const filteredEquipments = applyFilters(window.allEquipments, filters, normalizeId);
    renderTable(filteredEquipments, equipmentTableBody, window.consolidatedCalibratedMap, window.externalMaintenanceSNs);
    updateEquipmentCount(filteredEquipments.length);
}


async function exportWithExcelJS(tableId, fileName) {
    // Esta função permanece igual, usando ExcelJS para exportar com estilos
    const table = document.getElementById(tableId);
    if (!table) return alert(`Tabela com ID "${tableId}" não encontrada.`);
    if (typeof ExcelJS === 'undefined') return alert('ERRO: Biblioteca ExcelJS não carregada.');
    
    outputDiv.textContent = `Gerando ${fileName}.xlsx...`;
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
        headerHTMLRows.forEach(tr => {
            if (tr.id === 'headerFilters') return; 
            const rowValues = [];
            tr.querySelectorAll('th').forEach(th => rowValues.push(th.textContent));
            const headerRow = worksheet.addRow(rowValues);
            headerRow.eachCell(cell => {
                cell.fill = headerFill;
                cell.font = headerFont;
                cell.alignment = { vertical: 'middle', horizontal: 'center' };
                cell.border = defaultBorder;
            });
        });

        const bodyHTMLRows = Array.from(table.querySelectorAll('tbody tr'));
        bodyHTMLRows.forEach(tr => {
            if (tr.querySelector('td')?.colSpan > 1) return; 
            const cellValues = Array.from(tr.querySelectorAll('td')).map(td => td.textContent);
            const addedRow = worksheet.addRow(cellValues);
            addedRow.eachCell(cell => {
                if (tr.classList.contains('calibrated-dhme')) cell.fill = calibratedFill;
                else if (tr.classList.contains('not-calibrated')) cell.fill = notCalibratedFill;
                
                if (tr.classList.contains('in-external-maintenance')) cell.font = maintenanceFont;
                else cell.font = defaultFont;
                
                cell.border = defaultBorder;
            });
        });

        worksheet.columns.forEach(column => {
            let maxLength = 0;
            column.eachCell({ includeEmpty: true }, cell => {
                maxLength = Math.max(maxLength, cell.value ? cell.value.toString().length : 0);
            });
            column.width = maxLength < 12 ? 12 : maxLength + 4;
        });
        
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = `${fileName}_${new Date().toISOString().slice(0, 10)}.xlsx`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);

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
exportButton.addEventListener('click', () => exportWithExcelJS('equipmentTable', 'equipamentos_filtrados'));
exportOsButton.addEventListener('click', () => exportWithExcelJS('osTable', 'os_abertas_filtradas'));

showEquipmentButton.addEventListener('click', () => toggleSectionVisibility('equipmentSection'));
showOsButton.addEventListener('click', () => toggleSectionVisibility('osSection'));
showRondaButton.addEventListener('click', () => toggleSectionVisibility('rondaSection')); 

// *** NOVOS EVENT LISTENERS PARA A RONDA 3.0 ***
startRondaScanButton.addEventListener('click', () => {
    if (window.allEquipments.length === 0) {
        alert("Por favor, carregue primeiro os ficheiros de dados antes de iniciar a ronda.");
        return;
    }
    startScanner();
});
stopRondaScanButton.addEventListener('click', stopScanner);
saveRondaButton.addEventListener('click', saveRonda);

document.addEventListener('DOMContentLoaded', () => {
    toggleSectionVisibility('equipmentSection');
});