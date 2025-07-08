// js/excelReader.js

/**
 * Lê um arquivo Excel e retorna os dados da primeira planilha ou de uma planilha específica
 * como um array de objetos.
 * @param {File} file - O objeto File do arquivo Excel.
 * @param {string} [sheetName=null] - O nome específico da planilha a ser lida. Se nulo, lê a primeira.
 * @returns {Promise<Array<Object>>} Uma promessa que resolve com os dados do Excel.
 */
export function readExcelFile(file, sheetName = null) { 
    return new Promise((resolve, reject) => {
        const reader = new FileReader();

        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });

                let targetSheetName;
                if (sheetName && workbook.SheetNames.includes(sheetName)) {
                    targetSheetName = sheetName; 
                } else {
                    targetSheetName = workbook.SheetNames[0]; 
                }

                const worksheet = workbook.Sheets[targetSheetName];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

                if (jsonData.length === 0) {
                    resolve([]); 
                    return;
                }

                const headers = jsonData[0]; 
                const rows = jsonData.slice(1); 

                // Mapeamento de cabeçalhos do Excel para chaves de objeto JavaScript.
                const headerMap = {
                    'TAG': 'TAG',
                    'EQUIPAMENTO': 'Equipamento',
                    'Equipamento': 'Equipamento',
                    'MODELO': 'Modelo',
                    'Modelo': 'Modelo',
                    'FABRICANTE': 'Fabricante',
                    'Fabricante': 'Fabricante',
                    'SETOR': 'Setor',
                    'Setor': 'Setor',
                    'Nº Série': 'NumeroSerie', 
                    'NÚMERO DE SÉRIE': 'NumeroSerie', 
                    'NUMERO DE SERIE': 'NumeroSerie', 
                    'Nº DE SÉRIE': 'NumeroSerie', 
                    'Nº de Série': 'NumeroSerie', 
                    'Patrimônio': 'Patrimonio', 
                    'PATRIMÔNIO': 'Patrimonio',
                    'PATRIMONIO': 'Patrimonio',
                    'SN': 'SN', 
                    'Status Calibração': 'StatusCalibacao', 
                    'STATUS CALIBRAÇÃO': 'StatusCalibacao', 
                    'StatusCalibacao': 'StatusCalibacao', 
                    'Data Vencimento Calibração': 'DataVencimentoCalibacao', 
                    'DATA VENCIMENTO CALIBRAÇÃO': 'DataVencimentoCalibacao', 
                    'Status Manutenção': 'StatusManutencao',
                    'STATUS MANUTENÇÃO': 'StatusManutencao',
                    'OS': 'OS', 
                    'Tipo de Manutenção': 'TipoDeManutencao', 
                    'TIPO DE MANUTENÇÃO': 'TipoDeManutencao', 
                    'Tipo de Manutencao': 'TipoDeManutencao', 
                    'Tipo Manutenção': 'TipoDeManutencao', 
                    // Cabeçalhos da planilha de consolidação (aba "Consolidação")
                    'Número de Série': 'NumeroSerieConsolidacao', 
                    'Fornecedor': 'FornecedorConsolidacao',     
                    'Data de Calibração': 'DataCalibracaoConsolidada', 
                };

                // Formata os dados para um array de objetos, usando as chaves padronizadas
                const formattedData = rows.map(row => {
                    let obj = {};
                    headers.forEach((header, index) => {
                        const originalHeader = String(header).trim(); 

                        let cleanKey = headerMap[originalHeader]; 

                        if (!cleanKey) {
                           const normalizedHeader = originalHeader.toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, ""); 
                           cleanKey = headerMap[normalizedHeader]; 
                        }

                        if (!cleanKey) {
                            cleanKey = originalHeader.replace(/[^a-zA-Z0-9]/g, ''); 
                        }

                        obj[cleanKey] = row[index];
                    });
                    return obj;
                });
                resolve(formattedData);
            } catch (error) {
                console.error("Erro detalhado na leitura do arquivo Excel:", error);
                reject(new Error('Erro ao ler o arquivo Excel. Certifique-se de que é um arquivo Excel válido e não está corrompido. Detalhes: ' + error.message));
            }
        };

        reader.onerror = (error) => {
            reject(new Error('Erro ao ler o arquivo: ' + error.target.error));
        };

        reader.readAsArrayBuffer(file);
    });
}