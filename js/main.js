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
    const searchInput = document.getElementById('searchInput'); // Elemento do buscador

    let allEquipmentData = []; // Contém equipamentos originais + divergentes injetados
    let originalEquipmentData = []; // Para armazenar apenas os equipamentos originais (para o filtro de setor)
    let allCalibrationData = []; // Calibrações lidas de TODAS as fontes
    let currentlyDisplayedData = []; // Dados atualmente visíveis na tabela (após filtros e busca)

    // --- 2. DECLARAÇÃO DA FUNÇÃO applyFilters ---
    // Esta função DEVE ser declarada ANTES de ser usada nos addEventListener
    const applyFilters = () => {
        let filteredData = allEquipmentData; // Começa com todos os dados (originais + divergentes)
        const selectedSector = sectorFilter.value;
        const selectedStatus = calibrationStatusFilter.value;
        const searchTerm = searchInput.value.trim().toLowerCase();

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
                const serial = String(eq['Nº Série'] || '').replace(/^0+/, '').toLowerCase(); // Normaliza SN
                const patrimonio = String(eq.Patrimônio || '').toLowerCase();

                // Busca em TAG, Nº Série (normalizado) ou Patrimônio
                return tag.includes(searchTerm) || serial.includes(searchTerm) || patrimonio.includes(searchTerm);
            });
        }

        currentlyDisplayedData = filteredData; // Atualiza os dados que estão sendo exibidos
        renderEquipmentTable(filteredData, equipmentTableBody, equipmentCountSpan);
    };

    // --- 3. EVENT LISTENERS QUE USAM applyFilters ---
    // Agora que applyFilters está definida, podemos adicionar os listeners
    if (sectorFilter) sectorFilter.addEventListener('change', applyFilters);
    if (calibrationStatusFilter) calibrationStatusFilter.addEventListener('change', applyFilters);
    
    // Listener para o buscador (verifica se o elemento existe)
    if (searchInput) {
        searchInput.addEventListener('input', applyFilters); // Aciona o filtro em cada digitação
    } else {
        console.error("Elemento com ID 'searchInput' não encontrado! Verifique o index.html."); // Ajuda a depurar
    }

    // Listener para o botão de exportar
    if (exportButton) {
        exportButton.addEventListener('click', () => {
            if (currentlyDisplayedData.length > 0) {
                exportTableToExcel(currentlyDisplayedData, 'Equipamentos_Calibacao_Filtrados');
                outputDiv.textContent = 'Exportando dados para Excel...';
            } else {
                outputDiv.textContent = 'Não há dados para exportar. Por favor, carregue e processe os arquivos primeiro.';
            }
        });
    } else {
        console.error("Elemento com ID 'exportButton' não encontrado! Verifique o index.html."); // Ajuda a depurar
    }

    // Listener para o botão de processar arquivos
    processButton.addEventListener('click', async () => {
        const files = fileInput.files;
        if (files.length === 0) {
            outputDiv.textContent = 'Por favor, selecione pelo menos um arquivo Excel.';
            return;
        }

        // Resetar variáveis e UI antes de processar novos arquivos
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

            let tempEquipmentData = []; // Temporário para equipamentos lidos
            let tempCalibrationData = []; // Temporário para calibrações lidas

            fileResults.forEach(result => {
                const { fileName, workbook } = result;

                // Processa planilha de Equipamentos
                if (workbook.SheetNames.includes('Equipamentos')) {
                    const parsedEquipments = parseEquipmentSheet(workbook.Sheets['Equipamentos']);
                    tempEquipmentData = tempEquipmentData.concat(parsedEquipments);
                    outputDiv.textContent += `\n- Arquivo de Equipamentos (${fileName}) carregado. Total: ${parsedEquipments.length} registros.`;
                }

                // Processa planilhas de Calibração (DHME, Sciencetech, etc.)
                workbook.SheetNames.forEach(sheetName => {
                    const parsedCalibrations = parseCalibrationSheet(workbook.Sheets[sheetName]);
                    if (parsedCalibrations.length > 0) {
                        // Lógica para identificar a origem da calibração
                        const calibrationsWithSource = parsedCalibrations.map(cal => {
                            let source = 'Desconhecida'; // Default

                            const lowerCaseFileName = fileName.toLowerCase();
                            const lowerCaseSheetName = sheetName.toLowerCase();

                            if (lowerCaseFileName.includes('sciencetech') || lowerCaseSheetName.includes('sciencetech')) {
                                source = 'Sciencetech';
                            } else if (lowerCaseFileName.includes('dhme') || lowerCaseSheetName.includes('dhme') || lowerCaseSheetName.includes('plan1')) {
                                source = 'DHME';
                            }
                            // Adicione mais 'else if' aqui para outras empresas/planilhas de calibração se necessário
                            // Ex: else if (lowerCaseFileName.includes('novaempresa')) { source = 'NovaEmpresa'; }

                            return {
                                ...cal,
                                _source: source,
                            };
                        });
                        tempCalibrationData = tempCalibrationData.concat(calibrationsWithSource);
                        outputDiv.textContent += `\n- Arquivo de Calibração (${fileName} - Planilha: ${sheetName}) carregado. Total: ${parsedCalibrations.length} registros.`;
                    }
                });
            });
            
            // Armazena a lista de equipamentos originais ANTES de injetar as divergências
            originalEquipmentData = tempEquipmentData;

            // Cruza os dados e identifica divergências
            const { equipmentData: processedEquipmentData, calibratedCount, notCalibratedCount, divergentCalibrations: newDivergentCalibrations } = crossReferenceData(originalEquipmentData, tempCalibrationData, outputDiv);
            
            // allEquipmentData agora contém equipamentos originais (com status atualizado) + equipamentos divergentes
            allEquipmentData = processedEquipmentData.concat(newDivergentCalibrations.map(cal => ({
                TAG: cal.TAG || 'N/A', // Pode ser que DHME/Sciencetech não tenha TAG
                Equipamento: cal.EQUIPAMENTO || 'N/A',
                Modelo: cal.MODELO || 'N/A',
                Fabricante: cal.FABRICANTE || cal.MARCA || 'N/A', // Tentar FABRICANTE ou MARCA
                Setor: cal.SETOR || 'N/A',
                'Nº Série': cal.SN || 'N/A',
                Patrimônio: cal.PATRIM || 'N/A',
                calibrationStatus: `Não Cadastrado (${cal._source || 'Desconhecido'})`, // Status com origem
                calibrations: [cal], // Mantém a calibração original para tooltip
                nextCalibrationDate: cal['DATA VAL'] || 'N/A'
            })));


            // Aplica os filtros iniciais para renderizar a tabela e popular filtros
            applyFilters();
            populateSectorFilter(originalEquipmentData, sectorFilter); // Popula filtro de setor APENAS com base nos equipamentos originais
            outputDiv.textContent += '\nProcessamento concluído. Verifique a tabela abaixo.';

            // Mensagem de resumo das divergências no output
            if (newDivergentCalibrations.length > 0) {
                const dhmeDivergences = newDivergentCalibrations.filter(cal => cal._source === 'DHME').length;
                const sciencetechDivergences = newDivergentCalibrations.filter(cal => cal._source === 'Sciencetech').length;
                const unknownDivergences = newDivergentCalibrations.filter(cal => cal._source === 'Desconhecida').length;

                outputDiv.textContent += `\n\n--- Calibrações com Divergência (${newDivergentCalibrations.length}) ---`;
                if (dhmeDivergences > 0) outputDiv.textContent += `\n  - DHME: ${dhmeDivergences}
