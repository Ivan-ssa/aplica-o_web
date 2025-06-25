// js/dataProcessor.js
export const crossReferenceData = (equipmentData, calibrationData, outputDiv) => {
    let calibratedCount = 0;
    let notCalibratedCount = 0;
    let divergentCalibrations = []; 

    const equipmentSerialsSet = new Set(
        equipmentData.map(eq => (eq['Nº Série'] ? String(eq['Nº Série']).replace(/^0+/, '').trim() : '')) 
    );

    if (equipmentData.length === 0 && calibrationData.length === 0) {
        outputDiv.textContent += '\nNenhum dado de equipamento ou calibração para cruzar.';
        return { equipmentData, calibratedCount, notCalibratedCount, divergentCalibrations };
    }

    if (equipmentData.length > 0) { 
        equipmentData.forEach(equipment => {
            const equipmentSerial = (equipment['Nº Série'] ? String(equipment['Nº Série']).replace(/^0+/, '').trim() : '');

            const matchingCalibrations = calibrationData.filter(cal => {
                const calibrationSN = (cal.SN ? String(cal.SN).replace(/^0+/, '').trim() : ''); 
                return equipmentSerial !== '' && calibrationSN !== '' && calibrationSN === equipmentSerial; 
            });

            if (matchingCalibrations.length > 0) {
                let latestDueDateObj = null;
                let latestDueDateFormatted = 'N/A';
                // Renomeado para evitar conflito de escopo e garantir que a origem seja a final
                let finalCalibrationSource = 'Desconhecida'; 

                // Itera para encontrar a data de validade mais recente E a origem associada
                matchingCalibrations.forEach(cal => {
                    const currentDateValString = cal['DATA VAL'];
                    // Pega a _source desta calibração, se existir
                    const currentCalSource = cal._source || 'Desconhecida'; 

                    if (currentDateValString && currentDateValString !== 'N/A') {
                        const parts = currentDateValString.split('/');
                        if (parts.length === 2 && !isNaN(parseInt(parts[0])) && !isNaN(parseInt(parts[1]))) {
                            const currentParsedDate = new Date(parseInt(parts[1]), parseInt(parts[0]) - 1, 1);

                            if (!isNaN(currentParsedDate.getTime())) {
                                if (!latestDueDateObj || currentParsedDate > latestDueDateObj) {
                                    latestDueDateObj = currentParsedDate;
                                    latestDueDateFormatted = currentDateValString;
                                    finalCalibrationSource = currentCalSource; // Atualiza a origem da calibração mais recente
                                }
                            }
                        }
                    } 
                    // Se não há data de validade para a calibração atual (como Sciencetech),
                    // mas ela tem uma origem e ainda não definimos uma origem final (baseada em data),
                    // atribuímos esta como a origem final. Isso dá prioridade à origem se não tiver data.
                    else if (finalCalibrationSource === 'Desconhecida' && currentCalSource !== 'Desconhecida') {
                        finalCalibrationSource = currentCalSource;
                    }
                });

                // Fallback: se, após verificar todas as calibrações, a origem final ainda é Desconhecida,
                // mas existem calibrações correspondentes, pegamos a origem da PRIMEIRA calibração.
                // Isso cobre casos onde nenhuma calibração tem DATA VAL e a lógica acima não pegou uma source.
                if (finalCalibrationSource === 'Desconhecida' && matchingCalibrations.length > 0) {
                    finalCalibrationSource = matchingCalibrations[0]._source || 'Desconhecida';
                }

                // ATRIBUIÇÃO DO STATUS: Usa a finalCalibrationSource diretamente
                equipment.calibrationStatus = `Calibrado (${finalCalibrationSource})`;
                equipment.calibrations = matchingCalibrations; 
                equipment.nextCalibrationDate = latestDueDateFormatted;
                calibratedCount++;
            } else {
                equipment.calibrationStatus = 'Não Calibrado';
                notCalibratedCount++;
                equipment.nextCalibrationDate = 'N/A';
            }
        });
    }

    if (calibrationData.length > 0) {
        calibrationData.forEach(cal => {
            const calibrationSN = (cal.SN ? String(cal.SN).replace(/^0+/, '').trim() : '');
            if (calibrationSN && !equipmentSerialsSet.has(calibrationSN)) {
                divergentCalibrations.push(cal);
            }
        });
    }

    if (equipmentData.length > 0 || calibrationData.length > 0) {
        outputDiv.textContent += `\n--- Cruzamento Concluído ---`;
        outputDiv.textContent += `\nEquipamentos Calibrados: ${calibratedCount}`;
        outputDiv.textContent += `\nEquipamentos Não Calibrados: ${notCalibratedCount}`;
        outputDiv.textContent += `\nCalibrações com Divergência (SN não encontrado em Equipamentos): ${divergentCalibrations.length}`;
    } else {
        outputDiv.textContent += '\nNenhum arquivo de equipamento ou calibração foi encontrado para processar.';
    }

    return { equipmentData, calibratedCount, notCalibratedCount, divergentCalibrations };
};
