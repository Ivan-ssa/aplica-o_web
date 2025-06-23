// js/dataProcessor.js
export const crossReferenceData = (equipmentData, calibrationData, outputDiv) => {
    let calibratedCount = 0;
    let notCalibratedCount = 0;
    let divergentCalibrations = []; // NOVO: Array para armazenar calibrações não encontradas

    // Criar um Set de todos os números de série de equipamentos para busca rápida
    const equipmentSerialsSet = new Set(
        equipmentData.map(eq => (eq['Nº Série'] ? String(eq['Nº Série']).replace(/^0+/, '') : '').trim())
    );

    if (equipmentData.length === 0 && calibrationData.length === 0) {
        outputDiv.textContent += '\nNenhum dado de equipamento ou calibração para cruzar.';
        return { equipmentData, calibratedCount, notCalibratedCount, divergentCalibrations };
    }

    // --- Processamento dos Equipamentos ---
    if (equipmentData.length > 0) { // Garante que há equipamentos para processar
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

                let latestDueDateObj = null;
                let latestDueDateFormatted = 'N/A';

                matchingCalibrations.forEach(cal => {
                    const currentDateValString = cal['DATA VAL'];
                    if (currentDateValString) {
                        const parts = currentDateValString.split('/');
                        if (parts.length === 2 && !isNaN(parseInt(parts[0])) && !isNaN(parseInt(parts[1]))) {
                            const currentParsedDate = new Date(parseInt(parts[1]), parseInt(parts[0]) - 1, 1);

                            if (!isNaN(currentParsedDate.getTime())) {
                                if (!latestDueDateObj || currentParsedDate > latestDueDateObj) {
                                    latestDueDateObj = currentParsedDate;
                                    latestDueDateFormatted = currentDateValString;
                                }
                            }
                        }
                    }
                });

                equipment.nextCalibrationDate = latestDueDateFormatted;
                calibratedCount++;
            } else {
                equipment.calibrationStatus = 'Não Calibrado';
                notCalibratedCount++;
                equipment.nextCalibrationDate = 'N/A';
            }
        });
    }

    // --- NOVO: Identificar calibrações que não encontraram um equipamento ---
    if (calibrationData.length > 0) {
        calibrationData.forEach(cal => {
            const calibrationSN = (cal.SN ? String(cal.SN).replace(/^0+/, '') : '').trim();
            // Se o SN da calibração NÃO está no Set de SNs dos equipamentos
            if (calibrationSN && !equipmentSerialsSet.has(calibrationSN)) {
                divergentCalibrations.push(cal);
            }
        });
    }

    // Atualizar mensagens de saída
    if (equipmentData.length > 0 || calibrationData.length > 0) {
        outputDiv.textContent += `\n--- Cruzamento Concluído ---`;
        outputDiv.textContent += `\nEquipamentos Calibrados: ${calibratedCount}`;
        outputDiv.textContent += `\nEquipamentos Não Calibrados: ${notCalibratedCount}`;
        outputDiv.textContent += `\nCalibrações com Divergência (SN não encontrado em Equipamentos): ${divergentCalibrations.length}`;
    } else {
        outputDiv.textContent += '\nNenhum arquivo de equipamento ou calibração foi encontrado para processar.';
    }


    // Retorna o novo array de divergências
    return { equipmentData, calibratedCount, notCalibratedCount, divergentCalibrations };
};
