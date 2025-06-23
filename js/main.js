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
    const searchInput = document.getElementById('searchInput'); // ESTE AQUI!

    let allEquipmentData = [];
    let originalEquipmentData = [];
    let allCalibrationData = [];
    let currentlyDisplayedData = [];

    // --- 2. DECLARAÇÃO DA FUNÇÃO applyFilters ---
    // ESTA FUNÇÃO DEVE SER DECLARADA ANTES DE SER USADA EM QUALQUER LISTENER
    const applyFilters = () => {
        let filteredData = allEquipmentData;
        const selectedSector = sectorFilter.value;
        const selectedStatus = calibrationStatusFilter.value;
        const searchTerm = searchInput.value.trim().toLowerCase(); // searchInput já declarado acima

        // Aplicar filtro por setor
        if (selectedSector !== "") {
            filteredData = filteredData.filter(eq => eq.Setor && eq.Setor.trim() === selectedSector);
        }

        // Aplicar filtro por status de calibração
        if (selectedStatus !== "") {
            filteredData = filteredData.filter(eq => eq.calibrationStatus === selectedStatus);
        }

        // Aplicar filtro de busca por termo
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

    // --- 3. EVENT LISTENERS QUE USAM applyFilters ---
    // Agora que applyFilters está definida, podemos adicionar os listeners
    if (sectorFilter) sectorFilter.addEventListener('change', applyFilters);
    if (calibrationStatusFilter) calibrationStatusFilter.addEventListener('change', applyFilters);
    // Este é o listener da linha 23 que está causando o erro
    if (searchInput) searchInput.addEventListener('input', applyFilters); // Garanta que searchInput não é null

    // ... (restante do código: exportButton.addEventListener, processButton.addEventListener) ...

    exportButton.addEventListener('click', () => {
        if (currentlyDisplayedData.length > 0) {
            exportTableToExcel(currentlyDisplayedData, 'Equipamentos_Calibracao_Filtrados');
            outputDiv.textContent = 'Exportando dados para Excel...';
        } else {
            outputDiv.textContent = 'Não há dados para exportar. Por favor, carregue e processe os arquivos primeiro.';
        }
    });

    processButton.addEventListener('click', async () => {
        const files = fileInput.files;
        if (files.length === 0) {
            outputDiv.textContent = 'Por favor, selecione pelo menos um arquivo Excel.';
            return;
        }

        outputDiv.textContent = 'Processando arquivos...';
        allEquipmentData = [];
        originalEquipmentData = [];
        allCalibrationData = [];
        equipmentTableBody.innerHTML = '';
        sectorFilter.innerHTML = '<option value="">Todos os Setores</option>';
        calibrationStatusFilter.value = "";
        equipmentCountSpan.textContent = `Total: 0 equipamentos`;
        currentlyDisplayedData = [];
        if (searchInput) searchInput.value = ''; // Limpa o campo de busca ao processar novos arquivos

        try {
            const fileResults = await Promise.all(Array.from(files).map(readFile));

            let tempEquipmentData = [];
            let tempCalibrationData = [];

            fileResults.forEach(result => {
                const { fileName, workbook } = result;

                if (workbook.SheetNames.includes('Equipamentos')) {
                    const parsedEquipments = parseEquipmentSheet(workbook.Sheets['Equipamentos']);
                    tempEquipmentData = tempEquipmentData.concat(parsedEquipments);
                    outputDiv.textContent += `\n- Arquivo de Equipamentos (${fileName}) carregado. Total: ${parsedEquipments.length} registros.`;
                }

                workbook.SheetNames.forEach(sheetName => {
                    const parsedCalibrations = parseCalibrationSheet(workbook.Sheets[sheetName]);
                    if (parsedCalibrations.length > 0) {
                        const calibrationsWithSource = parsedCalibrations.map(cal => ({
                            ...cal,
                            _source: fileName.toLowerCase().includes('sciencetech') ? 'Sciencetech' : 'DHME',
                        }));
                        tempCalibrationData = tempCalibrationData.concat(calibrationsWithSource);
                        outputDiv.textContent += `\n- Arquivo de Calibração (${fileName} - Planilha: ${sheetName}) carregado. Total: ${parsedCalibrations.length} registros.`;
                    }
                });
            });
            
            originalEquipmentData = tempEquipmentData;

            const { equipmentData: processedEquipmentData, calibratedCount, notCalibratedCount, divergentCalibrations: newDivergentCalibrations } = crossReferenceData(originalEquipmentData, tempCalibrationData, outputDiv);
            
            allEquipmentData = processedEquipmentData.concat(newDivergentCalibrations.map(cal => ({
                TAG: cal.TAG || 'N/A',
                Equipamento: cal.EQUIPAMENTO || 'N/A',
                Modelo: cal.MODELO || 'N/A',
                Fabricante: cal.MARCA || 'N/A',
                Setor: cal.SETOR || 'N/A',
                'Nº Série': cal.SN || 'N/A',
                Patrimônio: cal.PATRIM || 'N/A',
                calibrationStatus: `Não Cadastrado (${cal._source || 'Desconhecido'})`,
                calibrations: [cal],
                nextCalibrationDate: cal['DATA VAL'] || 'N/A'
            })));


            applyFilters();
            populateSectorFilter(originalEquipmentData, sectorFilter);
            outputDiv.textContent += '\nProcessamento concluído. Verifique a tabela abaixo.';

            if (newDivergentCalibrations.length > 0) {
                const dhmeDivergences = newDivergentCalibrations.filter(cal => cal._source === 'DHME').length;
                const sciencetechDivergences = newDivergentCalibrations.filter(cal => cal._source === 'Sciencetech').length;
                outputDiv.textContent += `\n\n--- Calibrações com Divergência (${newDivergentCalibrations.length}) ---`;
                outputDiv.textContent += `\n  - DHME: ${dhmeDivergences} (Status: "Não Cadastrado (DHME)")`;
                outputDiv.textContent += `\n  - Sciencetech: ${sciencetechDivergences} (Status: "Não Cadastrado (Sciencetech)")`;
                outputDiv.textContent += `\nListadas na tabela principal com o status correspondente.`;
            } else {
                outputDiv.textContent += `\n\nNão foram encontradas calibrações sem equipamento correspondente.`;
            }

        } catch (error) {
            outputDiv.textContent = `Ocorreu um erro geral no processamento: ${error.message}`;
            console.error("Erro no processamento:", error);
        }
    });
});
