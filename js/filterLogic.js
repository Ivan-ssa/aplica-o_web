// js/filterLogic.js

/**
 * Aplica filtros aos dados dos equipamentos.
 * @param {Array<Object>} allEquipments - O array completo de equipamentos do cadastro principal.
 * @param {Object} filters - Objeto contendo os valores dos filtros.
 * @param {Function} normalizeId - Função utilitária para normalizar IDs (removendo zeros à esquerda).
 * @returns {Array<Object>} O array de equipamentos filtrados ou divergências.
 */
export function applyFilters(allEquipments, filters, normalizeId) { 
    const mainEquipmentSNs = new Set(allEquipments.map(eq => normalizeId(eq.NumeroSerie))); 

    // Lógica para filtrar as DIVERGÊNCIAS DOS FORNECEDORES (agora da consolidação)
    if (filters.calibrationStatus.startsWith('Divergência (')) {
        let targetFornecedor = null; // Para Divergência (Fornecedor Específico)
        if (filters.calibrationStatus.includes('(') && filters.calibrationStatus.includes(')')) {
            targetFornecedor = filters.calibrationStatus.substring(
                filters.calibrationStatus.indexOf('(') + 1, 
                filters.calibrationStatus.indexOf(')')
            ).trim().toLowerCase();
        }

        let rawDataToFilter = window.consolidatedCalibrationsRawData;

        return rawDataToFilter.filter(item => {
            const sn = normalizeId(item.NumeroSerieConsolidacao || item.NumeroSerie || item.NºdeSérie || item['Nº de Série'] || item['Número de Série']);
            const fornecedorItem = String(item.FornecedorConsolidacao || item.Fornecedor || '').trim().toLowerCase();

            const isDivergent = sn && !mainEquipmentSNs.has(sn); 

            if (filters.calibrationStatus === 'Divergência (Todos Fornecedores)') {
                return isDivergent; 
            } else if (targetFornecedor) {
                return isDivergent && fornecedorItem.includes(targetFornecedor);
            }
            return false; 
        }).map(item => ({ 
            TAG: '', 
            Equipamento: item.Equipamento || '', 
            Modelo: item.Modelo || '', 
            Fabricante: item.FornecedorConsolidacao || item.Fornecedor || '', 
            Setor: '',
            NumeroSerie: item.NumeroSerieConsolidacao || item.NumeroSerie || item.NºdeSérie || item['Nº de Série'] || item['Número de Série'], 
            Patrimonio: item.Patrimonio || item.Patrimônio || '', 
            StatusCalibacao: filters.calibrationStatus, 
            DataVencimentoCalibacao: item.DataCalibracaoConsolidada || item['Data de Calibração'] || '', 
            StatusManutencao: '' 
        }));
    }

    // Lógica para filtrar os SEUS equipamentos (cadastro principal)
    let filteredData = allEquipments.filter(equipment => {
        const sectorMatchGlobal = filters.sector === '' || (equipment.Setor && equipment.Setor === filters.sector);

        let effectiveMaintenanceStatus = equipment?.StatusManutencao || '';
        const equipmentSN = normalizeId(equipment.NumeroSerie); 

        if (window.externalMaintenanceSNs && window.externalMaintenanceSNs.has(equipmentSN)) {
            effectiveMaintenanceStatus = 'Em Manutenção Externa';
        }
        const maintenanceMatch = filters.maintenance === '' || effectiveMaintenanceStatus === filters.maintenance; 

        const searchText = filters.search;
        const searchMatch = searchText === '' ||
                            (equipment.NumeroSerie && normalizeId(String(equipment.NumeroSerie)).includes(normalizeId(searchText))) || 
                            (equipment.Patrimonio && normalizeId(String(equipment.Patrimonio)).includes(normalizeId(searchText))) || 
                            (equipment.TAG && String(equipment.TAG).toLowerCase().includes(searchText.toLowerCase())); 

        let effectiveCalibrationStatus; 
        // Obtém o status de calibração efetivo a partir do mapa consolidado
        const calibInfo = window.consolidatedCalibratedMap.get(normalizeId(equipment.NumeroSerie));
        if (calibInfo) {
            effectiveCalibrationStatus = `Calibrado (${calibInfo.fornecedor})`;
        } else {
            const originalStatusLower = String(equipment?.StatusCalibacao || '').toLowerCase();
            if (originalStatusLower.includes('não calibrado') || originalStatusLower.includes('não cadastrado')) {
                 effectiveCalibrationStatus = 'Não Calibrado/Não Encontrado (Seu Cadastro)';
            } else if (originalStatusLower.includes('calibrado (total)')) {
                effectiveCalibrationStatus = 'Calibrado (Total)';
            } else {
                // Para "Não Cadastrado" ou "Não Calibrado" (do Excel original)
                effectiveCalibrationStatus = String(equipment?.StatusCalibacao || ''); 
                if (effectiveCalibrationStatus.trim() === '') {
                     effectiveCalibrationStatus = 'Não Calibrado/Não Encontrado (Seu Cadastro)';
                }
            }
        }

        const calibrationStatusMatch = filters.calibrationStatus === '' || 
                                       effectiveCalibrationStatus === filters.calibrationStatus; 

        return sectorMatchGlobal && calibrationStatusMatch && searchMatch && maintenanceMatch;
    });

    // Aplica os filtros de cabeçalho
    if (filters.headerFilters && Object.keys(filters.headerFilters).length > 0) {
        Object.keys(filters.headerFilters).forEach(property => {
            const filterValue = filters.headerFilters[property];

            if (Array.isArray(filterValue)) {
                if (filterValue.length === 0) {
                    filteredData = []; 
                } else {
                    filteredData = filteredData.filter(equipment => {
                        let equipmentPropertyValue;
                        if (property === 'statuscalibacao') {
                            const calibInfo = window.consolidatedCalibratedMap.get(normalizeId(equipment.NumeroSerie));
                            if (calibInfo) {
                                equipmentPropertyValue = `calibrado (${calibInfo.fornecedor.toLowerCase()})`;
                            } else {
                                const originalStatusLower = String(equipment?.StatusCalibacao || '').toLowerCase();
                                if (originalStatusLower.includes('não calibrado') || originalStatusLower.includes('não cadastrado')) {
                                    equipmentPropertyValue = 'não calibrado/não encontrado (seu cadastro)';
                                } else if (originalStatusLower.includes('calibrado (total)')) {
                                    equipmentPropertyValue = 'calibrado (total)';
                                } else {
                                    equipmentPropertyValue = String(equipment?.StatusCalibacao || '').toLowerCase();
                                    if (equipmentPropertyValue.trim() === '') {
                                         equipmentPropertyValue = 'não calibrado/não encontrado (seu cadastro)';
                                    }
                                }
                            }
                        } else if (property === 'statusmanutencao') {
                            const equipmentSN = normalizeId(equipment.NumeroSerie);
                            if (window.externalMaintenanceSNs.has(equipmentSN)) {
                                equipmentPropertyValue = 'em manutenção externa';
                            } else {
                                equipmentPropertyValue = String(equipment?.StatusManutencao || '').toLowerCase();
                            }
                        } else if (property === 'numeroserie' || property === 'patrimonio') { 
                            equipmentPropertyValue = normalizeId(equipment[property]); 
                        }
                        else {
                            equipmentPropertyValue = String(equipment[property] || '').toLowerCase();
                        }

                        return filterValue.includes(equipmentPropertyValue);
                    });
                }
            } else { 
                filteredData = filteredData.filter(equipment => {
                    let equipmentPropertyValue;
                    if (property === 'numeroserie' || property === 'patrimonio') { 
                        equipmentPropertyValue = normalizeId(equipment[property]); 
                    } else {
                        equipmentPropertyValue = String(equipment[property] || '').toLowerCase();
                    }
                    return equipmentPropertyValue.includes(filterValue);
                });
            }
        });
    }

    return filteredData; 
}