// js/osRenderer.js

/**
 * Renderiza a tabela de Ordens de Serviço (OS) em aberto.
 * Cruza dados da OS com os status de calibração e manutenção dos equipamentos principais.
 * @param {Array<Object>} osData - O array de objetos de dados da OS.
 * @param {HTMLElement} tableBodyElement - O elemento <tbody> da tabela de OS.
 * @param {Map<string, Object>} mainEquipmentsBySN - Mapa de SN (normalizado) para o objeto do equipamento principal.
 * @param {Map<string, Object>} mainEquipmentsByPatrimonio - Mapa de Patrimônio (normalizado) para o objeto do equipamento principal.
 * @param {Map<string, Object>} consolidatedCalibratedMap - Mapa SN -> { fornecedor, dataCalibracao }.
 * @param {Set<string>} externalMaintenanceSNs - Set de SNs em manutenção externa.
 * @param {Function} normalizeId - Função utilitária para normalizar IDs (removendo zeros à esquerda).
 */
export function renderOsTable(
    osData, 
    tableBodyElement, 
    mainEquipmentsBySN, 
    mainEquipmentsByPatrimonio,
    consolidatedCalibratedMap, 
    externalMaintenanceSNs,
    normalizeId 
) {
    tableBodyElement.innerHTML = ''; 

    if (osData.length === 0) {
        const row = tableBodyElement.insertRow();
        const cell = row.insertCell();
        cell.colSpan = 8; 
        cell.textContent = 'Nenhuma Ordem de Serviço em aberto encontrada.';
        cell.style.textAlign = 'center';
        document.getElementById('osCount').textContent = `Total: 0 OS`;
        return;
    }

    let osCount = 0;
    osData.forEach(os => {
        const osSN = normalizeId(os?.['Nº de Série'] || os?.NºdeSérie || os?.NumeroSerie); 
        const osPatrimonio = normalizeId(os?.Patrimônio || os?.Patrimonio); 

        const correspondingEquipment = mainEquipmentsBySN.get(osSN) || mainEquipmentsByPatrimonio.get(osPatrimonio);

        let osStatusCalib = 'Não Encontrado'; 
        let osStatusManutencao = 'N/A'; 

        if (correspondingEquipment) {
            const equipmentMainSN = normalizeId(correspondingEquipment.NumeroSerie); 

            // Determina o status de calibração do equipamento principal USANDO O MAPA CONSOLIDADO
            const calibInfo = consolidatedCalibratedMap.get(equipmentMainSN);
            if (calibInfo) {
                osStatusCalib = `Calibrado (${calibInfo.fornecedor})`; 
            } else {
                const originalCalibStatusLower = String(correspondingEquipment.StatusCalibacao || '').toLowerCase();
                if (originalCalibStatusLower.includes('não calibrado') || originalCalibStatusLower.includes('não cadastrado')) {
                    osStatusCalib = 'Não Calibrado/Não Encontrado (Seu Cadastro)';
                } else if (originalCalibStatusLower.includes('calibrado (total)')) {
                    osStatusCalib = 'Calibrado (Total)';
                } else {
                    osStatusCalib = 'Desconhecido'; 
                }
            }

            // Determina o status de manutenção do equipamento principal
            if (externalMaintenanceSNs.has(equipmentMainSN)) {
                osStatusManutencao = 'Em Manutenção Externa'; 
            } else {
                osStatusManutencao = 'Não em Manutenção Externa'; 
            }
        } else {
             osStatusCalib = 'Equip. Não Cadastrado';
             osStatusManutencao = 'Equip. Não Cadastrado';
        }

        const row = tableBodyElement.insertRow();
        osCount++; 

        if (osStatusCalib.includes('Calibrado (')) { 
            row.classList.add('calibrated-dhme'); 
        } else if (osStatusCalib.includes('Não Calibrado')) { 
            row.classList.add('not-calibrated');
        }

        if (osStatusManutencao.includes('Em Manutenção Externa')) {
            row.classList.add('in-external-maintenance'); 
        }


        row.insertCell().textContent = os.OS ?? '';
        row.insertCell().textContent = os.Patrimonio ?? os.Patrimônio ?? ''; 
        row.insertCell().textContent = os.NumeroSerie ?? os?.['Nº de Série'] ?? os.NºdeSérie ?? ''; 
        row.insertCell().textContent = os.Equipamento ?? '';
        row.insertCell().textContent = os.Modelo ?? '';
        row.insertCell().textContent = os.Fabricante ?? '';
        row.insertCell().textContent = correspondingEquipment?.Setor ?? 'Não Cadastrado';
        // row.insertCell().textContent = osStatusCalib; 
        // row.insertCell().textContent = osStatusManutencao; 
    });

    document.getElementById('osCount').textContent = `Total: ${osCount} OS`;
}