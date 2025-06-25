// js/main.js
// ... (imports e declarações de consts e let) ...

document.addEventListener('DOMContentLoaded', () => {
    // ... (restante do código até o processButton.addEventListener) ...

    processButton.addEventListener('click', async () => {
        // ... (código de processamento, incluindo reset de variáveis) ...

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
                        const calibrationsWithSource = parsedCalibrations.map(cal => {
                            let source = 'Desconhecida'; // Default

                            const lowerCaseFileName = fileName.toLowerCase();
                            const lowerCaseSheetName = sheetName.toLowerCase();

                            // LÓGICA DE IDENTIFICAÇÃO DE ORIGEM MAIS ROBUSTA E ORDENADA
                            // Prioriza Sciencetech
                            if (lowerCaseFileName.includes('sciencetech') || lowerCaseSheetName.includes('sciencetech')) {
                                source = 'Sciencetech';
                            } 
                            // Em seguida, verifica DHMED
                            else if (lowerCaseFileName.includes('dhmed') || lowerCaseFileName.includes('dhme') || 
                                     lowerCaseSheetName.includes('dhmed') || lowerCaseSheetName.includes('dhme') || 
                                     lowerCaseSheetName.includes('plan1') || lowerCaseSheetName.includes('planilha1') ||
                                     lowerCaseFileName.includes('dhm') || lowerCaseSheetName.includes('dhm')) { // Adicionado 'dhm' para ser mais genérico
                                source = 'DHMED'; 
                            }
                            // Adicione outros 'else if' aqui para novas empresas no futuro:
                            // else if (lowerCaseFileName.includes('nome_da_nova_empresa')) {
                            //    source = 'NomeDaNovaEmpresa';
                            // }

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
            
            // ... (restante do código igual, incluindo a atribuição em allEquipmentData e populateSectorFilter) ...
        } catch (error) {
            // ...
        }
    });
});
