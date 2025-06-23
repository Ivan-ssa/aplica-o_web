// js/main.js
import { readFile, parseEquipmentSheet, parseCalibrationSheet } from './excelReader.js';
import { crossReferenceData } from './dataProcessor.js';
import { renderEquipmentTable, populateSectorFilter } from './uiRenderer.js';
import { exportTableToExcel } from './excelExporter.js';

document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('excelFileInput');
    const processButton = document.getElementById('processButton');
    const outputDiv = document.getElementById('output');
    const equipmentTableBody = document.querySelector('#equipmentTable tbody');
    const sectorFilter = document.getElementById('sectorFilter');
    const calibrationStatusFilter = document.getElementById('calibrationStatusFilter');
    const equipmentCountSpan = document.getElementById('equipmentCount');
    const exportButton = document.getElementById('exportButton');
    const searchInput = document.getElementById('searchInput'); // <--- GARANTIR QUE ESTA LINHA ESTÁ DESCOMENTADA E NO LUGAR CERTO

    // Resto do código...
    // ...

    // NOVO: Adiciona o event listener para o input de busca (se o elemento for encontrado)
    if (searchInput) {
        searchInput.addEventListener('input', applyFilters); // Aciona o filtro em cada digitação
    } else {
        console.error("Elemento com ID 'searchInput' não encontrado!"); // Ajudar a depurar se o ID estiver errado no HTML
    }

    // ... Resto do código da função processButton.addEventListener('click', async () => { ...
    // E aqui, dentro do try{} do processButton.addEventListener:
    // searchInput.value = ''; // <--- Descomente esta linha para limpar a busca ao processar novos arquivos
    // ...
});
