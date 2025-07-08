// js/divergenceRenderer.js

/**
 * Renderiza a tabela de divergências de números de série dos fornecedores.
 * @param {Array<Object>} divergences - O array de objetos de divergência.
 * @param {HTMLElement} tableBodyElement - O elemento <tbody> da tabela de divergências.
 */
export function renderSupplierDivergenceTable(divergences, tableBodyElement) {
    tableBodyElement.innerHTML = ''; // Limpa a tabela antes de preencher

    if (divergences.length === 0) {
        const row = tableBodyElement.insertRow();
        const cell = row.insertCell();
        cell.colSpan = 3; // Cobre todas as colunas da tabela de divergências
        cell.textContent = 'Nenhuma divergência de número de série encontrada nas listas dos fornecedores.';
        cell.style.textAlign = 'center';
        document.getElementById('divergenceCount').textContent = `Total: 0 divergências`;
        return;
    }

    divergences.forEach(divergence => {
        const row = tableBodyElement.insertRow();
        row.insertCell().textContent = divergence.fornecedor;
        row.insertCell().textContent = divergence.numeroSerieDivergente;
        row.insertCell().textContent = divergence.arquivoOrigem;
    });

    document.getElementById('divergenceCount').textContent = `Total: ${divergences.length} divergências`;
}