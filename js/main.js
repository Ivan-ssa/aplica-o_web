// js/main.js
// ... (imports e outras declarações de consts e let) ...

document.addEventListener('DOMContentLoaded', () => {
    // ... (restante do código até o processButton.addEventListener) ...

    let allEquipmentData = [];
    let originalEquipmentData = [];
    let allCalibrationData = [];
    let allMaintenanceData = []; 
    let currentlyDisplayedData = [];

    // EXPOR VARIÁVEIS GLOBAIS PARA DEPURAÇÃO
    window.allEquipmentData = allEquipmentData; 
    window.tempMaintenanceSNs_DEBUG = null; // NOVO: Para expor os SNs lidos da manutenção

    // ... (resto do código igual) ...

    processButton.addEventListener('click', async () => {
        // ... (código de processamento, incluindo reset de variáveis) ...

        let tempEquipmentData = []; 
        let tempCalibrationData = []; 
        let tempMaintenanceSNs = []; // Array que armazena os SNs da manutenção

        try {
            const fileResults = await Promise.all(Array.from(files).map(readFile));

            fileResults.forEach(result => {
                const { fileName, workbook } = result;

                if (workbook.SheetNames.includes('Equipamentos')) {
                    const parsedEquipments = parseEquipmentSheet(workbook.Sheets['Equipamentos']);
                    tempEquipmentData = tempEquipmentData.concat(parsedEquipments);
                    outputDiv.textContent += `\n- Arquivo de Equipamentos (${fileName}) carregado. Total: ${parsedEquipments.length} registros.`
                }

                workbook.SheetNames.forEach(sheetName => {
                    const lowerCaseFileName = fileName.toLowerCase();
                    const lowerCaseSheetName = sheetName.toLowerCase();
                    
                    const parsedCalibrations = parseCalibrationSheet(workbook.Sheets[sheetName]);
                    if (parsedCalibrations.length > 0) {
                        const calibrationsWithSource = parsedCalibrations.map(cal => {
                            let source = 'Desconhecida'; 
                            
                            if (lowerCaseFileName.includes('sciencetech') || lowerCaseSheetName.includes('sciencetech')) {
                                source = 'Sciencetech';
                            } else if (lowerCaseFileName.includes('dhmed') || lowerCaseFileName.includes('dhme') || 
                                       lowerCaseSheetName.includes('dhmed') || lowerCaseSheetName.includes('dhme') || 
                                       lowerCaseSheetName.includes('plan1') || lowerCaseSheetName.includes('planilha1') ||
                                       lowerCaseFileName.includes('dhm') || lowerCaseSheetName.includes('dhm')) { 
                                source = 'DHMED'; 
                            }
                            return { ...cal, _source: source };
                        });
                        tempCalibrationData = tempCalibrationData.concat(calibrationsWithSource);
                        outputDiv.textContent += `\n- Arquivo de Calibração (${fileName} - Planilha: ${sheetName}) carregado. Total: ${parsedCalibrations.length} registros.`;
                    }

                    if (lowerCaseFileName.includes('manutencao_externa') || lowerCaseSheetName.includes('manutencao_externa') || 
                        lowerCaseSheetName.includes('man_ext') || lowerCaseSheetName.includes('manut_ext') ||
                        lowerCaseFileName.includes('manu_externa') || lowerCaseSheetName.includes('manu_externa')) { 
                         // NOVO: parseMaintenanceSheet retorna APENAS os SNs
                         const parsedMaintenanceSNs = parseMaintenanceSheet(workbook.Sheets[sheetName]);
                         tempMaintenanceSNs = tempMaintenanceSNs.concat(parsedMaintenanceSNs);
                         outputDiv.textContent += `\n- Arquivo de Manutenção Externa (${fileName} - Planilha: ${sheetName}) carregado. Total: ${parsedMaintenanceSNs.length} registros.`;
                    }
                });
            });
            
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
                calibrationStatus: `Não Cadastrado (${cal._source || 'Desconhecida'})`, 
                calibrations: [cal], 
                nextCalibrationDate: cal['DATA VAL'] || 'N/A',
                maintenanceStatus: 'Não Aplicável' 
            })));

            // CRUZAMENTO PARA MANUTENÇÃO EXTERNA (AGORA COM SET DE SNs)
            if (tempMaintenanceSNs.length > 0) {
                const maintenanceSNsSet = new Set(tempMaintenanceSNs); // Converte para Set para busca rápida
                window.tempMaintenanceSNs_DEBUG = Array.from(maintenanceSNsSet); // NOVO: Expor para depuração

                allEquipmentData.forEach(eq => {
                    const equipmentId = (eq['Nº Série'] ? String(eq['Nº Série']).replace(/^0+/, '').trim() : '') || (eq.Patrimônio ? String(eq.Patrimônio).trim() : '');
                    
                    if (equipmentId && maintenanceSNsSet.has(equipmentId)) {
                        eq.maintenanceStatus = 'Em Manutenção Externa'; 
                    } else if (!eq.maintenanceStatus || eq.maintenanceStatus === 'Não Aplicável') { 
                        eq.maintenanceStatus = 'Não Aplicável';
                    }
                });
                outputDiv.textContent += `\n- Dados de Manutenção Externa cruzados com ${tempMaintenanceSNs.length} registros (SNs).`;
            } else {
                 outputDiv.textContent += `\n- Nenhum arquivo de Manutenção Externa processado.`;
            }

            window.allEquipmentData = allEquipmentData; 
            
            applyFilters();
            populateSectorFilter(originalEquipmentData, sectorFilter); 
            outputDiv.textContent += '\nProcessamento concluído. Verifique a tabela abaixo.';

            if (newDivergentCalibrations.length > 0) {
                const dhmedDivergences = newDivergentCalibrations.filter(cal => cal._source === 'DHMED').length;
                const sciencetechDivergences = newDivergentCalibrations.filter(cal => cal._source === 'Sciencetech').length;
                const unknownDivergences = newDivergentCalibrations.filter(cal => cal._source === 'Desconhecida').length;

                outputDiv.textContent += `\n\n--- Calibrações com Divergência (${newDivergentCalibrations.length}) ---`;
                if (dhmedDivergences > 0) outputDiv.textContent += `\n  - DHMED: ${dhmedDivergences} (Status: "Não Cadastrado (DHMED)")`;
                if (sciencetechDivergences > 0) outputDiv.textContent += `\n  - Sciencetech: ${sciencetechDivergences} (Status: "Não Cadastrado (Sciencetech)")`;
                if (unknownDivergences > 0) outputDiv.textContent += `\n  - Desconhecida: ${unknownDivergences} (Status: "Não Cadastrado (Desconhecido)")`;
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
