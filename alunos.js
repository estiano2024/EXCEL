
 let alunosData = [];
 let originalData = [];
 
 // Esperar carregar
 document.addEventListener('DOMContentLoaded', function() {
     // Configurar o input de arquivo
     document.getElementById('fileInput').addEventListener('change', handleFile);
     
     // Configurar os eventos dos filtros
     document.getElementById('escolaFilter').addEventListener('change', filterData);
     document.getElementById('cursoFilter').addEventListener('change', filterData);
     document.getElementById('salaFilter').addEventListener('change', filterData);
     document.getElementById('situacaoFilter').addEventListener('change', filterData);
     document.getElementById('search').addEventListener('input', filterData);
     document.getElementById('resetFilters').addEventListener('click', resetFilters);
 });
 
 // Função para lidar com o upload do arquivo
 function handleFile(e) {
     const file = e.target.files[0];
     if (!file) return;
     
     const reader = new FileReader();
     reader.onload = function(e) {
         const data = new Uint8Array(e.target.result);
         const workbook = XLSX.read(data, { type: 'array' });
         
         // Pegar a primeira planilha
         const firstSheetName = workbook.SheetNames[0];
         const worksheet = workbook.Sheets[firstSheetName];
         
         // Converter para JSON
         const jsonData = XLSX.utils.sheet_to_json(worksheet);
         
         // Processar os dados
         processData(jsonData);
     };
     reader.readAsArrayBuffer(file);
 }
 
 // Processar os dados da planilha
 function processData(data) {
     originalData = data;
     alunosData = [...originalData];
     
     // Mostrar a seção de filtros
     document.getElementById('filtersSection').classList.remove('hidden');
     
     // Preencher os dropdowns de filtro
     populateFilterOptions();
     
     // Atualizar a tabela
     updateTable();
     
     // Atualizar o resumo
     updateSummary();
 }
 
 // Popular as opções dos filtros
 function populateFilterOptions() {
     const escolas = [...new Set(originalData.map(item => item.ESCOLA))];
     const cursos = [...new Set(originalData.map(item => item.CURSO))];
     const salas = [...new Set(originalData.map(item => item.SALA))];
     const situacoes = [...new Set(originalData.map(item => item['Situação do Aluno'] || item.Situação))];
     
     fillSelect('escolaFilter', escolas);
     fillSelect('cursoFilter', cursos);
     fillSelect('salaFilter', salas);
     fillSelect('situacaoFilter', situacoes);
 }
 
 // Preencher um elemento select com opções
 function fillSelect(selectId, options) {
     const select = document.getElementById(selectId);
     
     // Manter a primeira opção (Todas)
     while (select.options.length > 1) {
         select.remove(1);
     }
     
     options.forEach(option => {
         if (option) { // Ignorar valores vazios
             const opt = document.createElement('option');
             opt.value = option;
             opt.textContent = option;
             select.appendChild(opt);
         }
     });
 }
 
 // Filtrar os dados com base nos filtros selecionados
 function filterData() {
     const escola = document.getElementById('escolaFilter').value;
     const curso = document.getElementById('cursoFilter').value;
     const sala = document.getElementById('salaFilter').value;
     const situacao = document.getElementById('situacaoFilter').value;
     const searchTerm = document.getElementById('search').value.toLowerCase();
     
     alunosData = originalData.filter(item => {
         // Aplicar filtros
         if (escola !== 'all' && item.ESCOLA !== escola) return false;
         if (curso !== 'all' && item.CURSO !== curso) return false;
         if (sala !== 'all' && item.SALA !== sala) return false;
         if (situacao !== 'all' && (item['Situação do Aluno'] || item.Situação) !== situacao) return false;
         
         // Aplicar busca geral
         if (searchTerm) {
             const searchFields = [
                 item.Aluno || '',
                 item.RA || '',
                 item['Email Microsoft'] || '',
                 item['Email Google'] || '',
                 item.SALA || '',
                 item.ESCOLA || '',
                 item.CURSO || ''
             ].join(' ').toLowerCase();
             
             if (!searchFields.includes(searchTerm)) return false;
         }
         
         return true;
     });
     
     updateTable();
     updateSummary();
 }
 
 // Resetar todos os filtros
 function resetFilters() {
     document.getElementById('escolaFilter').value = 'all';
     document.getElementById('cursoFilter').value = 'all';
     document.getElementById('salaFilter').value = 'all';
     document.getElementById('situacaoFilter').value = 'all';
     document.getElementById('search').value = '';
     
     alunosData = [...originalData];
     updateTable();
     updateSummary();
 }
 
 // Atualizar a tabela com os dados filtrados
 function updateTable() {
     const tbody = document.getElementById('tableBody');
     tbody.innerHTML = '';
     
     alunosData.forEach(item => {
         const row = document.createElement('tr');
         
         row.innerHTML = `
             <td>${item.Chamada || ''}</td>
             <td>${item.Aluno || ''}</td>
             <td>${item.RA || ''}</td>
             <td>${item['Dig. RA'] || ''}</td>
             <td>${item.Sexo || ''}</td>
             <td>${formatDate(item.Nascimento) || ''}</td>
             <td>${item['Email Microsoft'] || ''}</td>
             <td>${item['Email Google'] || ''}</td>
             <td>${item['Situação do Aluno'] || item.Situação || ''}</td>
             <td>${item.SALA || ''}</td>
             <td>${item.ESCOLA || ''}</td>
             <td>${item.ANO || ''}</td>
             <td>${item.CURSO || ''}</td>
         `;
         
         tbody.appendChild(row);
     });
     
     document.getElementById('filteredCount').textContent = `Filtrados: ${alunosData.length}`;
 }
 
 // Atualizar o resumo com estatísticas
 function updateSummary() {
     document.getElementById('totalCount').textContent = originalData.length;
     
     // Contagem por escola
     const escolas = [...new Set(originalData.map(item => item.ESCOLA))];
     let escolaSummary = '<h4>Por Escola:</h4><ul>';
     
     escolas.forEach(escola => {
         const count = originalData.filter(item => item.ESCOLA === escola).length;
         escolaSummary += `<li>${escola}: ${count} alunos</li>`;
     });
     escolaSummary += '</ul>';
     
     // Contagem por curso
     const cursos = [...new Set(originalData.map(item => item.CURSO))];
     let cursoSummary = '<h4>Por Curso:</h4><ul>';
     
     cursos.forEach(curso => {
         const count = originalData.filter(item => item.CURSO === curso).length;
         cursoSummary += `<li>${curso}: ${count} alunos</li>`;
     });
     cursoSummary += '</ul>';
     
     // Contagem por sala
     const salas = [...new Set(originalData.map(item => item.SALA))];
     let salaSummary = '<h4>Por Sala:</h4><ul>';
     
     salas.forEach(sala => {
         const count = originalData.filter(item => item.SALA === sala).length;
         salaSummary += `<li>${sala}: ${count} alunos</li>`;
     });
     salaSummary += '</ul>';
     
     document.getElementById('summaryContent').innerHTML = escolaSummary + cursoSummary + salaSummary;
 }
 
 // Formatador de data
 function formatDate(dateValue) {
     if (!dateValue) return '';
     
     // Se for um objeto de data do Excel
     if (typeof dateValue === 'object' && dateValue instanceof Date) {
         return dateValue.toLocaleDateString('pt-BR');
     }
     
     // Se for uma string de data
     if (typeof dateValue === 'string') {
         const date = new Date(dateValue);
         if (!isNaN(date.getTime())) {
             return date.toLocaleDateString('pt-BR');
         }
         return dateValue; // Retorna o valor original se não for uma data válida
     }
     
     return dateValue;
 }