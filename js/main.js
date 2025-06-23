// js/main.js
import { readFile, parseEquipmentSheet, parseCalibrationSheet } from './excelReader.js';
import { crossReferenceData } from './dataProcessor.js';
import { renderEquipmentTable, populateSectorFilter } from './uiRenderer.js';
import { exportTableToExcel } from './excelExporter.js';

document.addEventListener('DOMContentLoaded', () => {
    // TODAS AS REFERÊNCIAS A ELEMENTOS HTML DEVEM ESTAR AQUI DENTRO
    const fileInput = document.getElementById('excelFileInput');
    const processButton = document.getElementById('processButton');
    const outputDiv = document.getElementById('output');
    const equipmentTableBody = document.querySelector('#equipmentTable tbody');
    const sectorFilter = document.getElementById('sectorFilter');
    const calibrationStatusFilter = document.getElementById('calibrationStatusFilter');
    const equipmentCountSpan = document.getElementById('equipmentCount');
    const exportButton = document.getElementById('exportButton'); // Garanta que este ID está no HTML

    // Verifica se algum elemento é null antes de adicionar listeners
    // Isso é uma verificação de segurança extra para ajudar a depurar
    if (!fileInput) console.error("Elemento com ID 'excelFileInput' não encontrado!");
    if (!processButton) console.error("Elemento com ID 'processButton' não encontrado!");
    if (!outputDiv) console.error("Elemento com ID 'output' não encontrado!");
    if (!equipmentTableBody) console.error("Elemento com seletor '#equipmentTable tbody' não encontrado!");
    if (!sectorFilter) console.error("Elemento com ID 'sectorFilter' não encontrado!");
    if (!calibrationStatusFilter) console.error("Elemento com ID 'calibrationStatusFilter' não encontrado!");
    if (!equipmentCountSpan) console.error("Elemento com ID 'equipmentCount' não encontrado!");
    if (!exportButton) console.error("Elemento com ID 'exportButton' não encontrado!");


    let allEquipmentData = [];
    let allCalibrationData = [];
    let currentlyDisplayedData = [];
    let divergentCalibrations = [];

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
        currentlyDisplayedData = filteredData;
        renderEquipmentTable(filteredData, equipmentTableBody, equipmentCountSpan);
    };

    // Listeners para os filtros (dentro de DOMContentLoaded)
    if (sectorFilter) sectorFilter.addEventListener('change', applyFilters);
    if (calibrationStatusFilter) calibrationStatusFilter.addEventListener('change', applyFilters);

    // Listener para o botão de exportar (dentro de DOMContentLoaded)
    if (exportButton) { // Verifica se o botão existe antes de adicionar o listener
        exportButton.addEventListener('click', () => {
            if (currentlyDisplayedData.length > 0) {
                exportTableToExcel(currentlyDisplayedData, 'Equipamentos_Calibracao_Filtrados');
                outputDiv.textContent = 'Exportando dados para Excel...';
            } else {
                outputDiv.textContent = 'Não há dados para exportar. Por favor, carregue e processe os arquivos primeiro.';
            }
        });
    }


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
        currentlyDisplayedData = [];
        divergentCalibrations = [];

        try {
            const fileResults = await Promise.all(Array.from(files).map(readFile));

            fileResults.forEach(result => {
                const { fileName, workbook } = result;

                if (workbook.SheetNames.includes('Equipamentos')) {
                    const parsedEquipments = parseEquipmentSheet(workbook.Sheets['Equipamentos']);
                    allEquipmentData = allEquipmentData.concat(parsedEquipments);
                    outputDiv.textContent += `\n- Arquivo de Equipamentos (${fileName}) carregado. Total: ${parsedEquipments.length} registros.`;
                }

                workbook.SheetNames.forEach(sheetName => {
                    const parsedCalibrations = parseCalibrationSheet(workbook.Sheets[sheetName]);
                    if (parsedCalibrations.length > 0) {
                        allCalibrationData = allCalibrationData.concat(parsedCalibrations);
                        outputDiv.textContent += `\n- Arquivo de Calibração (${fileName} - Planilha: ${sheetName}) carregado. Total: ${parsedCalibrations.length} registros.`;
                    }
                });
            });

            const { equipmentData, calibratedCount, notCalibratedCount, divergentCalibrations: newDivergentCalibrations } = crossReferenceData(allEquipmentData, allCalibrationData, outputDiv);
            allEquipmentData = equipmentData;
            divergentCalibrations = newDivergentCalibrations;

            applyFilters();
            populateSectorFilter(allEquipmentData, sectorFilter);
            outputDiv.textContent += '\nProcessamento concluído. Verifique a tabela abaixo.';

            if (divergentCalibrations.length > 0) {
                outputDiv.textContent += `\n\n--- Divergências Encontradas (${divergentCalibrations.length}) ---`;
                divergentCalibrations.forEach(divCal => {
                    outputDiv.textContent += `\n- SN: ${divCal.SN || 'N/A'}, Equipamento Calibração: ${divCal.EQUIPAMENTO || 'N/A'}, Data Val: ${divCal['DATA VAL'] || 'N/A'}`;
                });
            } else {
                outputDiv.textContent += `\n\nNão foram encontradas calibrações sem equipamento correspondente.`;
            }

        } catch (error) {
            outputDiv.textContent = `Ocorreu um erro geral no processamento: ${error.message}`;
            console.error("Erro no processamento:", error);
        }
    });
});
