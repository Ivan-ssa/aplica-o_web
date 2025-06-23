// js/excelExporter.js

export const exportTableToExcel = (data, filename = 'export_data') => {
    if (!data || data.length === 0) {
        console.warn("Nenhum dado para exportar.");
        return;
    }

    // Define os cabeçalhos que queremos exportar e a ordem
    // Eles devem corresponder exatamente às chaves dos objetos em 'data'
    const headers = [
        "TAG",
        "Equipamento",
        "Modelo",
        "Fabricante",
        "Setor",
        "Nº Série",
        "Patrimônio",
        "Status Calibração",
        "Data Vencimento Calibração"
    ];

    // Mapeia os dados para o formato que a planilha espera, na ordem correta dos cabeçalhos
    const exportRows = data.map(item => {
        let row = {};
        headers.forEach(header => {
            row[header] = item[header] !== undefined ? item[header] : '';
        });
        return row;
    });

    // Adiciona os cabeçalhos como a primeira linha do array de dados
    const ws = XLSX.utils.json_to_sheet(exportRows, { header: headers });

    // Cria um novo workbook (pasta de trabalho)
    const wb = XLSX.utils.book_new();
    // Adiciona a planilha ao workbook com o nome "Resultados"
    XLSX.utils.book_append_sheet(wb, ws, "Resultados");

    // Gera o arquivo Excel e inicia o download
    XLSX.writeFile(wb, `${filename}.xlsx`);
};
