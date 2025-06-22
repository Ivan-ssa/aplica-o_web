// js/dataProcessor.js
export const crossReferenceData = (equipmentData, calibrationData, outputDiv) => {
    let calibratedCount = 0;
    let notCalibratedCount = 0;

    if (equipmentData.length === 0 && calibrationData.length === 0) {
        outputDiv.textContent += '\nNenhum dado de equipamento ou calibração para cruzar.';
        return { equipmentData, calibratedCount, notCalibratedCount };
    }

    if (equipmentData.length > 0 && calibrationData.length > 0) {
        equipmentData.forEach(equipment => {
            // Normaliza o número de série do equipamento (remove zeros à esquerda e espaços)
            const equipmentSerial = (equipment['Nº Série'] ? String(equipment['Nº Série']).replace(/^0+/, '') : '').trim();

            const matchingCalibrations = calibrationData.filter(cal => {
                // Normaliza o SN da calibração (remove zeros à esquerda e espaços)
                const calibrationSN = (cal.SN ? String(cal.SN).replace(/^0+/, '') : '').trim();
                return equipmentSerial && calibrationSN === equipmentSerial;
            });

            if (matchingCalibrations.length > 0) {
                equipment.calibrationStatus = 'Calibrado';
                equipment.calibrations = matchingCalibrations;

                let latestDueDateObj = null; // Armazenará o objeto Date da data mais futura
                let latestDueDateFormatted = 'N/A'; // Armazenará a string formatada 'MM/YYYY'

                matchingCalibrations.forEach(cal => {
                    const currentDateValString = cal['DATA VAL']; // Já está no formato "MM/YYYY"
                    if (currentDateValString) {
                        // Tenta criar um objeto Date a partir da string "MM/YYYY"
                        const parts = currentDateValString.split('/');
                        if (parts.length === 2 && !isNaN(parseInt(parts[0])) && !isNaN(parseInt(parts[1]))) {
                            // Cria a data como o primeiro dia do mês/ano para comparação
                            // Mês no JS é 0-indexado (Janeiro = 0, Fevereiro = 1, etc.)
                            const currentParsedDate = new Date(parseInt(parts[1]), parseInt(parts[0]) - 1, 1);

                            if (!isNaN(currentParsedDate.getTime())) { // Verifica se é uma data válida
                                if (!latestDueDateObj || currentParsedDate > latestDueDateObj) {
                                    latestDueDateObj = currentParsedDate;
                                    latestDueDateFormatted = currentDateValString; // Guarda a string original formatada
                                }
                            }
                        }
                    }
                });

                equipment.nextCalibrationDate = latestDueDateFormatted; // Atribui a string formatada final
                calibratedCount++;
            } else {
                equipment.calibrationStatus = 'Não Calibrado';
                notCalibratedCount++;
                equipment.nextCalibrationDate = 'N/A';
            }
        });
        outputDiv.textContent += `\n--- Cruzamento Concluído ---\nEquipamentos Calibrados: ${calibratedCount}\nEquipamentos Não Calibrados: ${notCalibratedCount}`;
    } else if (equipmentData.length > 0) {
        equipmentData.forEach(eq => {
            eq.calibrationStatus = 'Não Calibrado';
            eq.nextCalibrationDate = 'N/A';
        });
        outputDiv.textContent += '\nNenhum arquivo de calibração encontrado. Todos os equipamentos considerados "Não Calibrados".';
    } else {
        outputDiv.textContent += '\nNenhum arquivo de "Equipamentos" com a planilha "Equipamentos" foi encontrado.';
    }
    return { equipmentData, calibratedCount, notCalibratedCount };
};
