// js/main.js
import { readFile, parseEquipmentSheet, parseCalibrationSheet } from './excelReader.js';
import { crossReferenceData } from './dataProcessor.js';
import { renderEquipmentTable, populateSectorFilter } from './uiRenderer.js';

document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('excelFileInput');
    const processButton = document.getElementById('processButton');
    const outputDiv = document.getElementById('output');
    const equipmentTableBody = document.querySelector('#equipmentTable tbody');
    const sectorFilter = document.getElementById('sectorFilter');
    const calibrationStatusFilter = document.getElementById('calibrationStatusFilter');
    const equipmentCountSpan = document.getElementById('equipmentCount');

    let allEquipmentData = [];
    let allCalibrationData = [];

    const applyFilters = () => {
        let filteredData = allEquipmentData;
        const selectedSector = sectorFilter.value;
        const selectedStatus = calibrationStatusFilter.value;

        if (selectedSector !== "") {
            filteredData = filteredData.filter(eq => eq.Setor && eq.Setor.trim() === selectedSector);
        }
        if (selectedStatus !== "") {
            filteredData = filteredData.filter(eq => eq.calibrationStatus === selectedStatus);
        }
        renderEquipmentTable(filteredData, equipmentTableBody, equipmentCountSpan);
    };

    sectorFilter.addEventListener('change', applyFilters);
    calibrationStatusFilter.addEventListener('change', applyFilters);

    processButton.addEventListener('click', async () => {
        const files = fileInput.files;
        if (files.length === 0) {
            outputDiv.textContent = 'Por favor, selecione pelo menos um arquivo Excel.';
            return;
        }

        outputDiv.textContent = 'Processando arquivos...';
        allEquipmentData = [];
        allCalibrationData = [];
        equipmentTableBody.innerHTML = '';
        sectorFilter.innerHTML = '<option value="">Todos os Setores</option>';
        calibrationStatusFilter.value = "";
        equipmentCountSpan.textContent = `Total: 0 equipamentos`;

        try {
            const fileResults = await Promise.all(Array.from(files).map(readFile));

            fileResults.forEach(result => {
                const { fileName, workbook } = result;

                // Processa planilha de Equipamentos
                if (workbook.SheetNames.includes('Equipamentos')) {
                    const parsedEquipments = parseEquipmentSheet(workbook.Sheets['Equipamentos']);
                    allEquipmentData = allEquipmentData.concat(parsedEquipments);
                    outputDiv.textContent += `\n- Arquivo de Equipamentos (${fileName}) carregado. Total: ${parsedEquipments.length} registros.`;
                }

                // Processa planilhas de Calibração
                workbook.SheetNames.forEach(sheetName => {
                    const parsedCalibrations = parseCalibrationSheet(workbook.Sheets[sheetName]);
                    if (parsedCalibrations.length > 0) {
                        allCalibrationData = allCalibrationData.concat(parsedCalibrations);
                        outputDiv.textContent += `\n- Arquivo de Calibração (${fileName} - Planilha: ${sheetName}) carregado. Total: ${parsedCalibrations.length} registros.`;
                    }
                });
            });

            // Cruza os dados
            const { equipmentData, calibratedCount, notCalibratedCount } = crossReferenceData(allEquipmentData, allCalibrationData, outputDiv);
            allEquipmentData = equipmentData; // Atualiza a variável global com os dados cruzados

            // Renderiza e popula UI
            renderEquipmentTable(allEquipmentData, equipmentTableBody, equipmentCountSpan);
            populateSectorFilter(allEquipmentData, sectorFilter);
            applyFilters(); // Para garantir que a contagem e os filtros iniciais estejam corretos
            outputDiv.textContent += '\nProcessamento concluído. Verifique a tabela abaixo.';

        } catch (error) {
            outputDiv.textContent = `Ocorreu um erro geral no processamento: ${error.message}`;
            console.error("Erro no processamento:", error);
        }
    });
});