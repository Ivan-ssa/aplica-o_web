// js/main.js
import { readFile, parseEquipmentSheet, parseCalibrationSheet } from './excelReader.js';
import { crossReferenceData } from './dataProcessor.js';
// uiRenderer.js vai precisar de uma pequena alteração
import { renderEquipmentTable, populateSectorFilter, renderDivergentCalibrationsTable } from './uiRenderer.js'; 
import { exportTableToExcel } from './excelExporter.js'; 

document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('excelFileInput');
    const processButton = document.getElementById('processButton');
    const outputDiv = document.getElementById('output');
    const equipmentTableBody = document.querySelector('#equipmentTable tbody');
    const sectorFilter = document.getElementById('sectorFilter');
    const calibrationStatusFilter = document.getElementById('calibrationStatusFilter');
    const equipmentCountSpan = document.getElementById('equipmentCount');
    const exportButton = document.getElementById('exportButton');
    // REMOVER a referência à tabela de divergências se você a removeu do HTML
    // const divergentCalibrationsTableBody = document.querySelector('#divergentCalibrationsTable tbody');

    let allEquipmentData = []; // Armazenará equipamentos + "equipamentos" divergentes
    let allCalibrationData = []; 
    let currentlyDisplayedData = []; 
    // A variável divergentCalibrations não será mais necessária como um array separado aqui,
    // pois os itens divergentes serão injetados em allEquipmentData com um status especial.
    // let divergentCalibrations = []; 

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

    sectorFilter.addEventListener('change', applyFilters);
    calibrationStatusFilter.addEventListener('change', applyFilters);

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
        allCalibrationData = [];
        equipmentTableBody.innerHTML = '';
        sectorFilter.innerHTML = '<option value="">Todos os Setores</option>';
        calibrationStatusFilter.value = "";
        equipmentCountSpan.textContent = `Total: 0 equipamentos`;
        currentlyDisplayedData = [];
        // divergentCalibrations = []; // Remover esta linha

        // Limpar a tabela de divergências se ela ainda estiver no HTML
        // if (divergentCalibrationsTableBody) {
        //     divergentCalibrationsTableBody.innerHTML = '<tr><td colspan="6">Nenhum dado processado.</td></tr>';
        // }


        try {
            const fileResults = await Promise.all(Array.from(files).map(readFile));

            let tempEquipmentData = []; // Para armazenar equipamentos lidos antes de adicionar divergências
            let tempCalibrationData = []; // Para armazenar calibrações lidas

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
                        tempCalibrationData = tempCalibrationData.concat(parsedCalibrations);
                        outputDiv.textContent += `\n- Arquivo de Calibração (${fileName} - Planilha: ${sheetName}) carregado. Total: ${parsedCalibrations.length} registros.`;
                    }
                });
            });

            // Re-executa o cruzamento de dados para injetar divergências
            // E agora retorna os dados de divergência para que possamos injetá-los em allEquipmentData
            const { equipmentData: processedEquipmentData, calibratedCount, notCalibratedCount, divergentCalibrations: newDivergentCalibrations } = crossReferenceData(tempEquipmentData, tempCalibrationData, outputDiv);
            
            // Aqui está a grande mudança: adicionamos os "equipamentos" divergentes à lista principal
            allEquipmentData = processedEquipmentData.concat(newDivergentCalibrations.map(cal => ({
                TAG: cal.TAG || 'N/A',
                Equipamento: cal.EQUIPAMENTO || 'N/A',
                Modelo: cal.MODELO || 'N/A',
                Fabricante: cal.MARCA || 'N/A', // Usar MARCA como Fabricante
                Setor: cal.SETOR || 'N/A',
                'Nº Série': cal.SN || 'N/A', // SN do DHME
                Patrimônio: cal.PATRIM || 'N/A',
                calibrationStatus: 'Não Cadastrado (DHME)', // NOVO STATUS PARA DIVERGÊNCIA
                calibrations: [cal], // Mantenha a calibração original aqui para referência
                nextCalibrationDate: cal['DATA VAL'] || 'N/A'
            })));


            // allEquipmentData = equipmentData; // Esta linha será removida ou ajustada

            applyFilters(); // Renderiza com a lista completa, incluindo os divergentes
            populateSectorFilter(allEquipmentData, sectorFilter);
            outputDiv.textContent += '\nProcessamento concluído. Verifique a tabela abaixo.';

            // REMOVER o código de renderização de divergências na área de output,
            // pois elas agora estarão na tabela principal com o novo status
            // if (divergentCalibrations.length > 0) {
            //     outputDiv.textContent += `\n\n--- Calibrações com Divergência (${divergentCalibrations.length}) listadas na tabela abaixo. ---`;
            // } else {
            //     outputDiv.textContent += `\n\nNão foram encontradas calibrações sem equipamento correspondente.`;
            // }

        } catch (error) {
            outputDiv.textContent = `Ocorreu um erro geral no processamento: ${error.message}`;
            console.error("Erro no processamento:", error);
        }
    });
});
