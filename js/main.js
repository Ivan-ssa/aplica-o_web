// js/main.js
import { readFile, parseEquipmentSheet, parseCalibrationSheet } from './excelReader.js';
import { crossReferenceData } from './dataProcessor.js';
import { renderEquipmentTable, populateSectorFilter } from './uiRenderer.js';
import { exportTableToExcel } from './excelExporter.js';

document.addEventListener('DOMContentLoaded', () => {
    // --- 1. DECLARAÇÃO DE TODOS OS ELEMENTOS HTML E VARIÁVEIS ---
    const fileInput = document.getElementById('excelFileInput');
    const processButton = document.getElementById('processButton');
    const outputDiv = document.getElementById('output');
    const equipmentTableBody = document.querySelector('#equipmentTable tbody');
    const sectorFilter = document.getElementById('sectorFilter');
    const calibrationStatusFilter = document.getElementById('calibrationStatusFilter');
    const equipmentCountSpan = document.getElementById('equipmentCount');
    const exportButton = document.getElementById('exportButton');
    const searchInput = document.getElementById('searchInput');

    let allEquipmentData = []; // Esta é a variável local
    let originalEquipmentData = [];
    let allCalibrationData = [];
    let currentlyDisplayedData = [];

    // NOVO: Expor allEquipmentData para o escopo global para depuração
    // REMOVA ESTA LINHA DEPOIS QUE A DEPURAÇÃO TERMINAR!
    window.allEquipmentData = allEquipmentData; 
    // FIM DO NOVO

    // --- 2. DECLARAÇÃO DA FUNÇÃO applyFilters ---
    const applyFilters = () => {
        let filteredData = allEquipmentData;
        const selectedSector = sectorFilter.value;
        const selectedStatus = calibrationStatusFilter.value;
        const searchTerm = searchInput.value.trim().toLowerCase();

        if (selectedSector !== "") {
            filteredData = filteredData.filter(eq => eq.Setor && eq.Setor.trim() === selectedSector);
        }

        if (selectedStatus !== "") {
            if (selectedStatus === "Calibrado (Total)") {
                filteredData = filteredData.filter(eq => 
                    eq.calibrationStatus === "Calibrado (DHMED)" || 
                    eq.calibrationStatus === "Calibrado (Sciencetech)"
                );
            } else {
                filteredData = filteredData.filter(eq => eq.calibrationStatus === selectedStatus);
            }
        }

        if (searchTerm !== "") {
            filteredData = filteredData.filter(eq => {
                const tag = String(eq.TAG || '').toLowerCase();
                const serial = String(eq['Nº Série'] || '').replace(/^0+/, '').toLowerCase();
                const patrimonio = String(eq.Patrimônio || '').toLowerCase();
                return tag.includes(searchTerm) || serial.includes(searchTerm) || patrimonio.includes(searchTerm);
            });
        }

        currentlyDisplayedData = filteredData;
        renderEquipmentTable(filteredData, equipmentTableBody, equipmentCountSpan);
    };

    // ... (resto do código igual) ...

    processButton.addEventListener('click', async () => {
        // ... (código de processamento) ...
        
        try {
            // ... (leitura dos arquivos) ...

            // ... (população de tempEquipmentData e tempCalibrationData) ...
            
            originalEquipmentData = tempEquipmentData;

            const { equipmentData: processedEquipmentData, calibratedCount, notCalibratedCount, divergentCalibrations: newDivergentCalibrations } = crossReferenceData(originalEquipmentData, tempCalibrationData, outputDiv);
            
            allEquipmentData = processedEquipmentData.concat(newDivergentCalibrations.map(cal => ({
                TAG: cal.TAG || 'N/A',
                Equipamento: cal.EQUIPAMENTO || 'N/A',
                Modelo: cal.MODELO || 'N/A',
                Fabricante: cal.FABRICANTE || cal.MARCA || 'N/A',
                Setor: cal.SETOR || 'N/A',
                'Nº Série': cal.SN || 'N/A',
                Patrimônio: cal.PATRIM || 'N/A',
                calibrationStatus: `Não Cadastrado (${cal._source || 'Desconhecido'})`,
                calibrations: [cal],
                nextCalibrationDate: cal['DATA VAL'] || 'N/A'
            })));

            // NOVO: Atualizar a referência global para o array populado
            window.allEquipmentData = allEquipmentData; 
            // FIM DO NOVO

            applyFilters();
            populateSectorFilter(originalEquipmentData, sectorFilter);
            outputDiv.textContent += '\nProcessamento concluído. Verifique a tabela abaixo.';

            // ... (mensagens de divergência) ...

        } catch (error) {
            outputDiv.textContent = `Ocorreu um erro geral no processamento: ${error.message}`;
            console.error("Erro no processamento:", error);
        }
    });
});
