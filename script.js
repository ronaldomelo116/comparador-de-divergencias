// Configura o worker do PDF.js para processar arquivos PDF sem travar a interface
if (typeof pdfjsLib !== "undefined") {
  // Define o caminho do worker compatível com a versão carregada no index.html
  pdfjsLib.GlobalWorkerOptions.workerSrc =
    "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
}

// Lista branca de extensões permitidas para validação no JavaScript
const ALLOWED_EXTENSIONS = ["xlsx", "xlsm", "csv", "pdf"];

// Referências dos elementos de upload do SAP no DOM
const sapFileInput = document.getElementById("sap-file-input");
const sapUploadCard = document.getElementById("sap-upload-card");
const sapFileNameDisplay = document.getElementById("sap-file-name");

// Referências dos elementos de upload da Revan no DOM
const revanFileInput = document.getElementById("revan-file-input");
const revanUploadCard = document.getElementById("revan-upload-card");
const revanFileNameDisplay = document.getElementById("revan-file-name");

// Referências dos elementos de status e botões principais
const statusMessage = document.getElementById("status-message");
const compareBtn = document.getElementById("compare-btn");
const resetBtn = document.getElementById("reset-btn");

// Referências da seção de resumo numérico
const summarySection = document.getElementById("summary-section");
const totalItemsCount = document.getElementById("total-items-count");
const matchItemsCount = document.getElementById("match-items-count");
const diffItemsCount = document.getElementById("diff-items-count");

// Referências da seção de resultados, filtros e tabela
const resultsSection = document.getElementById("results-section");
const filterAllBtn = document.getElementById("filter-all-btn");
const filterDiffBtn = document.getElementById("filter-diff-btn");
const exportCsvBtn = document.getElementById("export-csv-btn");
const resultsBody = document.getElementById("results-body");

// Estado central da aplicação para controlar arquivos, resultados e filtro ativo
const appState = {
  // Armazena o objeto File válido do SAP
  sapFile: null,
  // Armazena o objeto File válido da Revan
  revanFile: null,
  // Armazena o array final da comparação para filtrar e exportar
  comparisonResults: [],
  // Controla qual filtro está ativo ("all" ou "diff")
  activeFilter: "all"
};

// Exibe uma mensagem de erro ou aviso usando textContent (proteção contra XSS)
function showStatusMessage(message) {
  // Define o texto da mensagem com segurança
  statusMessage.textContent = message;
  // Remove a classe que oculta o alerta
  statusMessage.classList.remove("is-hidden");
}

// Oculta a mensagem de status e limpa seu conteúdo
function clearStatusMessage() {
  // Limpa o texto anterior
  statusMessage.textContent = "";
  // Oculta o elemento de alerta
  statusMessage.classList.add("is-hidden");
}

// Extrai a extensão do arquivo em letras minúsculas
function getFileExtension(fileName) {
  // Verifica se o nome possui ponto
  if (!fileName || !fileName.includes(".")) {
    // Retorna string vazia se não houver extensão
    return "";
  }
  // Divide pelo ponto e retorna a última parte em minúsculas
  return fileName.split(".").pop().toLowerCase().trim();
}

// Habilita o botão Comparar apenas quando os dois arquivos forem válidos
function updateCompareButtonState() {
  // Verifica se ambos os arquivos estão carregados no estado
  const hasBothFiles = Boolean(appState.sapFile && appState.revanFile);
  // Habilita ou desabilita o botão principal
  compareBtn.disabled = !hasBothFiles;
}

// Valida e registra o arquivo selecionado para SAP ou Revan
function handleFileSelection(sourceKey, inputElement, cardElement, nameElement) {
  // Limpa qualquer mensagem de erro anterior
  clearStatusMessage();

  // Pega o primeiro arquivo selecionado pelo usuário
  const selectedFile = inputElement.files[0];

  // Se o usuário cancelou a seleção, reseta apenas esse campo
  if (!selectedFile) {
    // Remove o arquivo do estado
    appState[sourceKey] = null;
    // Restaura o texto padrão do cartão
    nameElement.textContent = "Nenhum arquivo selecionado";
    // Remove o destaque visual de arquivo carregado
    cardElement.classList.remove("has-file");
    // Atualiza o botão de comparar
    updateCompareButtonState();
    // Encerra a função
    return;
  }

  // Extrai a extensão real do arquivo selecionado
  const extension = getFileExtension(selectedFile.name);

  // Verifica se a extensão está na lista branca permitida
  if (!ALLOWED_EXTENSIONS.includes(extension)) {
    // Limpa o input para descartar o arquivo inválido
    inputElement.value = "";
    // Garante que o estado fique vazio
    appState[sourceKey] = null;
    // Restaura o texto padrão
    nameElement.textContent = "Nenhum arquivo selecionado";
    // Remove o destaque visual
    cardElement.classList.remove("has-file");
    // Atualiza o botão de comparar
    updateCompareButtonState();
    // Exibe aviso amigável informando os formatos aceitos
    showStatusMessage(
  `Formato inválido (${selectedFile.name}). Envie apenas arquivos .xlsx, .csv ou .pdf.`
);
    // Interrompe o processamento
    return;
  }

  // Salva o arquivo válido no estado da aplicação
  appState[sourceKey] = selectedFile;
  // Exibe o nome do arquivo selecionado com segurança via textContent
  nameElement.textContent = selectedFile.name;
  // Aplica a borda verde indicando arquivo pronto
  cardElement.classList.add("has-file");
  // Verifica se já pode liberar o botão Comparar
  updateCompareButtonState();
}

// Normaliza textos de cabeçalho removendo acentos, espaços extras e maiúsculas
function normalizeKey(text) {
  // Converte para string, remove acentos e deixa em minúsculas
  return String(text || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .trim();
}

// Converte valores numéricos (inclusive com vírgula brasileira) em número válido
function parseQuantity(rawValue) {
  // Se já for número válido, retorna direto
  if (typeof rawValue === "number" && !Number.isNaN(rawValue)) {
    return rawValue;
  }
  // Converte para string e remove espaços
  const cleaned = String(rawValue || "").trim();
  // Retorna 0 se estiver vazio
  if (!cleaned) {
    return 0;
  }
  // Troca vírgula decimal por ponto e converte para número
  const parsed = Number(cleaned.replace(/\./g, "").replace(",", "."));
  // Se falhar, tenta conversão direta (caso o número use ponto decimal)
  const fallback = Number(cleaned.replace(",", "."));
  // Retorna o valor válido ou 0
  return Number.isNaN(fallback) ? (Number.isNaN(parsed) ? 0 : parsed) : fallback;
}

// Busca o valor de uma coluna testando múltiplos nomes possíveis de cabeçalho
function findColumnValue(row, possibleHeaders) {
  // Percorre todas as chaves (colunas) da linha atual
  const rowKeys = Object.keys(row);
  // Procura a primeira chave que corresponda a um dos nomes esperados
  for (const key of rowKeys) {
    // Normaliza o nome da coluna da planilha
    const normalized = normalizeKey(key);
    // Verifica se está na lista de cabeçalhos aceitos
    if (possibleHeaders.includes(normalized)) {
      // Retorna o valor encontrado
      return row[key];
    }
  }
  // Retorna vazio se a coluna não existir nessa linha
  return "";
}

// Lê arquivos .xlsm ou .csv usando SheetJS e retorna linhas normalizadas
async function parseSpreadsheetFile(file, sourceType) {
  // Lê o arquivo como ArrayBuffer binário
  const arrayBuffer = await file.arrayBuffer();
  // Abre a planilha com o SheetJS
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  // Pega o nome da primeira aba disponível
  const firstSheetName = workbook.SheetNames[0];
  // Acessa os dados da primeira aba
  const worksheet = workbook.Sheets[firstSheetName];
  // Converte a aba em array de objetos mantendo células vazias como ""
  const rawRows = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

  // Mapeia cada linha para o formato padrão interno
  const normalizedRows = [];

  // Percorre todas as linhas lidas da planilha
  for (const row of rawRows) {
    // Define os nomes possíveis de colunas conforme a origem (SAP ou Revan)
    const codeValue =
      sourceType === "sap"
        ? findColumnValue(row, ["material", "codigo", "cod"])
        : findColumnValue(row, ["codigo", "material", "cod"]);

    // Extrai a quantidade da coluna correspondente
    const qtyValue =
      sourceType === "sap"
        ? findColumnValue(row, ["quantidade", "qtd", "quant"])
        : findColumnValue(row, ["quant", "quantidade", "qtd"]);

    // Extrai a descrição do item (presente principalmente na base da Revan)
    const descriptionValue = findColumnValue(row, [
      "descricao do item",
      "descricao",
      "texto breve material",
      "item"
    ]);

    // Extrai o centro (ex: PI01)
    const centerValue = findColumnValue(row, ["centro", "plant"]);

    // Limpa o código do material removendo espaços
    const cleanCode = String(codeValue || "").trim();

    // Ignora linhas vazias ou de rodapé sem código de material
    if (!cleanCode) {
      continue;
    }

    // Adiciona o registro padronizado ao array
    normalizedRows.push({
      code: cleanCode,
      description: String(descriptionValue || "").trim(),
      center: String(centerValue || "").trim(),
      qty: parseQuantity(qtyValue)
    });
  }

  // Retorna todos os registros normalizados
  return normalizedRows;
}

// Lê arquivos .pdf usando PDF.js e extrai código, descrição e quantidade por linha
async function parsePdfFile(file) {
  // Verifica se a biblioteca PDF.js foi carregada corretamente
  if (typeof pdfjsLib === "undefined") {
    throw new Error("Biblioteca PDF.js não carregada. Verifique sua conexão.");
  }

  // Lê o PDF como ArrayBuffer
  const arrayBuffer = await file.arrayBuffer();
  // Carrega o documento PDF na memória
  const pdfDocument = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;

  // Array para armazenar os registros extraídos de todas as páginas
  const extractedRows = [];

  // Percorre todas as páginas do PDF
  for (let pageNum = 1; pageNum <= pdfDocument.numPages; pageNum++) {
    // Carrega a página atual
    const page = await pdfDocument.getPage(pageNum);
    // Extrai os blocos de texto com suas posições (x, y)
    const textContent = await page.getTextContent();

    // Agrupa blocos de texto que estão na mesma altura (mesma linha visual)
    const linesMap = new Map();

    // Percorre cada fragmento de texto da página
    for (const item of textContent.items) {
      // Arredonda a coordenada Y para agrupar textos alinhados horizontalmente
      const yCoord = Math.round(item.transform[5] / 2) * 2;
      // Cria a lista da linha caso ainda não exista
      if (!linesMap.has(yCoord)) {
        linesMap.set(yCoord, []);
      }
      // Adiciona o texto e sua posição X para ordenar da esquerda para a direita
      linesMap.get(yCoord).push({
        x: item.transform[4],
        text: item.str.trim()
      });
    }

    // Ordena as linhas de cima para baixo (Y maior vem primeiro no PDF)
    const sortedYKeys = Array.from(linesMap.keys()).sort((a, b) => b - a);

    // Reconstrói cada linha de texto e tenta extrair código e quantidade
    for (const yKey of sortedYKeys) {
      // Ordena os fragmentos da linha da esquerda para a direita
      const lineItems = linesMap.get(yKey).sort((a, b) => a.x - b.x);
      // Junta os textos da linha separados por espaço
      const fullLine = lineItems
        .map((part) => part.text)
        .filter(Boolean)
        .join(" ");

      // Procura um código de material numérico (8 a 12 dígitos, padrão SAP/Revan)
      const codeMatch = fullLine.match(/\b(\d{8,12})\b/);

      // Se a linha não tiver código de material, ignora (cabeçalho/rodapé)
      if (!codeMatch) {
        continue;
      }

      // Captura o código encontrado
      const materialCode = codeMatch[1];

      // Procura o centro no padrão de 2 letras + 2 números (ex: PI01)
      const centerMatch = fullLine.match(/\b([A-Z]{2}\d{2})\b/);
      const centerCode = centerMatch ? centerMatch[1] : "";

      // Remove o código e o centro para isolar os números de quantidade na linha
      const lineWithoutCode = fullLine.replace(materialCode, " ");
      // Captura todos os números isolados restantes na linha
      const numberMatches = lineWithoutCode.match(/\b\d+(?:[.,]\d+)?\b/g);

      // Se não encontrar nenhum número de quantidade, pula a linha
      if (!numberMatches || numberMatches.length === 0) {
        continue;
      }

      // Pega o último número curto (evitando números de lote longos)
      const qtyCandidate =
        numberMatches.find((num) => num.length <= 6) || numberMatches[0];

      // Adiciona o registro extraído do PDF
      extractedRows.push({
        code: materialCode,
        description: fullLine,
        center: centerCode,
        qty: parseQuantity(qtyCandidate)
      });
    }
  }

  // Retorna os registros encontrados no PDF
  return extractedRows;
}

// Escolhe automaticamente o leitor correto conforme a extensão do arquivo
async function parseUploadedFile(file, sourceType) {
  // Identifica a extensão do arquivo
  const extension = getFileExtension(file.name);

  // Encaminha para o leitor de PDF
  if (extension === "pdf") {
    return parsePdfFile(file);
  }

  // Encaminha para o leitor de planilhas (.xlsm e .csv)
  return parseSpreadsheetFile(file, sourceType);
}

// Agrupa códigos repetidos em um Map e soma suas quantidades
function aggregateByMaterialCode(rows) {
  // Cria um Map indexado pelo código do material
  const aggregatedMap = new Map();

  // Percorre cada linha normalizada
  for (const row of rows) {
    // Se o código já existe no Map, apenas soma a quantidade
    if (aggregatedMap.has(row.code)) {
      // Recupera o registro acumulado
      const existing = aggregatedMap.get(row.code);
      // Soma a nova quantidade ao total já acumulado
      existing.qty += row.qty;
      // Preenche descrição se ainda estiver vazia
      if (!existing.description && row.description) {
        existing.description = row.description;
      }
      // Preenche centro se ainda estiver vazio
      if (!existing.center && row.center) {
        existing.center = row.center;
      }
    } else {
      // Cria uma nova entrada no Map para esse código
      aggregatedMap.set(row.code, {
        code: row.code,
        description: row.description,
        center: row.center,
        qty: row.qty
      });
    }
  }

  // Retorna o Map consolidado sem duplicidades
  return aggregatedMap;
}

// Cruza os dois Maps (SAP e Revan) e calcula as divergências
function compareDatasets(sapMap, revanMap) {
  // Junta todos os códigos únicos das duas bases usando um Set
  const allCodes = new Set([...sapMap.keys(), ...revanMap.keys()]);
  // Array que vai guardar o resultado final da comparação
  const results = [];

  // Percorre cada código único encontrado
  for (const code of allCodes) {
    // Busca o item no SAP (pode ser undefined se só existir na Revan)
    const sapItem = sapMap.get(code);
    // Busca o item na Revan (pode ser undefined se só existir no SAP)
    const revanItem = revanMap.get(code);

    // Define as quantidades (0 quando ausente em uma das bases)
    const sapQty = sapItem ? sapItem.qty : 0;
    const revanQty = revanItem ? revanItem.qty : 0;
    // Calcula a diferença (Revan - SAP)
    const difference = revanQty - sapQty;

    // Define o rótulo de situação e se há divergência
    let statusLabel = "Igual";
    let isDivergent = false;

    // Classifica o tipo de divergência
    if (!sapItem && revanItem) {
      statusLabel = "Apenas na Revan";
      isDivergent = true;
    } else if (sapItem && !revanItem) {
      statusLabel = "Apenas no SAP";
      isDivergent = true;
    } else if (difference !== 0) {
      statusLabel = "Divergente";
      isDivergent = true;
    }

    // Prioriza a descrição da Revan (ou usa "-" se não houver)
    const description =
      (revanItem && revanItem.description) ||
      (sapItem && sapItem.description) ||
      "—";

    // Prioriza o centro encontrado em qualquer uma das bases
    const center =
      (sapItem && sapItem.center) || (revanItem && revanItem.center) || "—";

    // Adiciona o resultado consolidado
    results.push({
      code,
      description,
      center,
      sapQty,
      revanQty,
      difference,
      statusLabel,
      isDivergent
    });
  }

  // Ordena deixando os itens divergentes primeiro e depois por código
  results.sort((a, b) => {
    // Divergentes sobem para o topo da tabela
    if (a.isDivergent !== b.isDivergent) {
      return a.isDivergent ? -1 : 1;
    }
    // Em seguida ordena pelo código do material
    return a.code.localeCompare(b.code);
  });

  // Retorna a lista completa comparada
  return results;
}

// Atualiza os três cards de resumo numérico
function renderSummary(results) {
  // Conta o total de códigos analisados
  const total = results.length;
  // Conta quantos itens têm divergência
  const divergent = results.filter((item) => item.isDivergent).length;
  // Calcula quantos bateram exatamente
  const matched = total - divergent;

  // Atualiza os valores no DOM via textContent
  totalItemsCount.textContent = String(total);
  matchItemsCount.textContent = String(matched);
  diffItemsCount.textContent = String(divergent);

  // Exibe a seção de resumo
  summarySection.classList.remove("is-hidden");
}

// Cria uma célula <td> com texto seguro (sem innerHTML)
function createTextCell(text) {
  // Cria o elemento de célula
  const td = document.createElement("td");
  // Insere o conteúdo como texto puro
  td.textContent = String(text);
  // Retorna a célula pronta
  return td;
}

// Renderiza as linhas da tabela de resultados respeitando o filtro ativo
function renderResultsTable() {
  // Limpa as linhas anteriores da tabela
  resultsBody.replaceChildren();

  // Filtra os itens conforme o botão selecionado ("all" ou "diff")
  const filteredItems =
    appState.activeFilter === "diff"
      ? appState.comparisonResults.filter((item) => item.isDivergent)
      : appState.comparisonResults;

  // Caso não haja itens para exibir no filtro atual
  if (filteredItems.length === 0) {
    // Cria uma linha de aviso amigável
    const emptyRow = document.createElement("tr");
    const emptyCell = document.createElement("td");
    // Faz a célula ocupar todas as 7 colunas da tabela
    emptyCell.colSpan = 7;
    // Define a mensagem de acordo com o filtro
    emptyCell.textContent =
      appState.activeFilter === "diff"
        ? "Nenhuma divergência encontrada entre SAP e Revan."
        : "Nenhum item encontrado nos arquivos enviados.";
    // Anexa a célula na linha e a linha na tabela
    emptyRow.appendChild(emptyCell);
    resultsBody.appendChild(emptyRow);
    // Exibe a seção de resultados
    resultsSection.classList.remove("is-hidden");
    return;
  }

  // Usa DocumentFragment para inserir todas as linhas de uma vez (melhor performance)
  const fragment = document.createDocumentFragment();

  // Percorre cada item filtrado e monta a linha <tr>
  for (const item of filteredItems) {
    // Cria a linha da tabela
    const tr = document.createElement("tr");

    // Formata o sinal positivo na diferença (ex: +9)
    const formattedDiff =
      item.difference > 0 ? `+${item.difference}` : String(item.difference);

    // Adiciona as 6 primeiras colunas de dados
    tr.appendChild(createTextCell(item.code));
    tr.appendChild(createTextCell(item.description));
    tr.appendChild(createTextCell(item.center));
    tr.appendChild(createTextCell(item.sapQty));
    tr.appendChild(createTextCell(item.revanQty));
    tr.appendChild(createTextCell(formattedDiff));

    // Cria a célula da coluna Situação com badge colorido
    const statusCell = document.createElement("td");
    const badge = document.createElement("span");
    // Aplica a classe base e a classe de cor conforme o resultado
    badge.className = item.isDivergent
      ? "status-badge status-badge-diff"
      : "status-badge status-badge-match";
    // Define o texto do badge com segurança
    badge.textContent = item.statusLabel;

    // Anexa o badge na célula e a célula na linha
    statusCell.appendChild(badge);
    tr.appendChild(statusCell);

    // Adiciona a linha ao fragmento
    fragment.appendChild(tr);
  }

  // Insere todas as linhas no corpo da tabela de uma única vez
  resultsBody.appendChild(fragment);
  // Exibe a seção de resultados
  resultsSection.classList.remove("is-hidden");
}

// Alterna o filtro ativo entre "all" e "diff"
function setActiveFilter(filterMode) {
  // Atualiza o filtro no estado
  appState.activeFilter = filterMode;
  // Atualiza o destaque visual nos botões de filtro
  filterAllBtn.classList.toggle("is-active", filterMode === "all");
  filterDiffBtn.classList.toggle("is-active", filterMode === "diff");
  // Re-renderiza a tabela com o novo filtro
  renderResultsTable();
}

// Escapa valores de texto para exportação segura em CSV (evita quebra de colunas)
function escapeCsvField(value) {
  // Converte o valor para string
  const text = String(value ?? "");
  // Protege contra injeção de fórmula no Excel (=, +, -, @ no início da célula)
  const safeText = /^[=+\-@]/.test(text) ? `'${text}` : text;
  // Envolve em aspas duplas duplicando eventuais aspas internas
  return `"${safeText.replace(/"/g, '""')}"`;
}

// Exporta os itens divergentes para um arquivo .csv compatível com Excel em português
function exportDivergencesToCsv() {
  // Seleciona apenas os itens divergentes (ou todos se não houver divergências)
  const divergentItems = appState.comparisonResults.filter(
    (item) => item.isDivergent
  );

  // Se não houver divergências, avisa o usuário
  if (divergentItems.length === 0) {
    showStatusMessage("Não há divergências para exportar.");
    return;
  }

  // Cabeçalho do CSV usando ponto e vírgula (padrão Excel Brasil)
  const headers = [
    "Material",
    "Descrição",
    "Centro",
    "Qtd. SAP",
    "Qtd. Revan",
    "Diferença",
    "Situação"
  ];

  // Monta as linhas do CSV
  const csvLines = [headers.join(";")];

  // Converte cada item divergente em uma linha formatada
  for (const item of divergentItems) {
    const row = [
      escapeCsvField(item.code),
      escapeCsvField(item.description),
      escapeCsvField(item.center),
      item.sapQty,
      item.revanQty,
      item.difference,
      escapeCsvField(item.statusLabel)
    ];
    csvLines.push(row.join(";"));
  }

  // Adiciona BOM UTF-8 (\uFEFF) para o Excel reconhecer acentos corretamente
  const csvContent = "\uFEFF" + csvLines.join("\r\n");
  // Cria um Blob com o conteúdo do CSV
  const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
  // Gera uma URL temporária para download
  const downloadUrl = URL.createObjectURL(blob);

  // Cria um link temporário para disparar o download
  const tempLink = document.createElement("a");
  tempLink.href = downloadUrl;
  tempLink.download = "divergencias_sap_revan.csv";
  // Dispara o clique programaticamente
  tempLink.click();

  // Libera a memória da URL temporária
  URL.revokeObjectURL(downloadUrl);
}

// Executa todo o fluxo de leitura, soma e comparação ao clicar no botão principal
async function handleCompareClick() {
  // Limpa alertas anteriores
  clearStatusMessage();

  // Garante que os dois arquivos existem antes de processar
  if (!appState.sapFile || !appState.revanFile) {
    showStatusMessage("Selecione os dois arquivos (SAP e Revan) antes de comparar.");
    return;
  }

  // Desabilita o botão e indica processamento enquanto lê os arquivos
  compareBtn.disabled = true;
  compareBtn.textContent = "Comparando...";

  try {
    // Lê os dois arquivos em paralelo para ganhar velocidade
    const [sapRows, revanRows] = await Promise.all([
      parseUploadedFile(appState.sapFile, "sap"),
      parseUploadedFile(appState.revanFile, "revan")
    ]);

    // Verifica se algum dos arquivos não retornou dados válidos
    if (sapRows.length === 0 || revanRows.length === 0) {
      showStatusMessage(
        "Não foi possível encontrar colunas válidas (Material/Quantidade ou Código/Quant) em um dos arquivos."
      );
      return;
    }

    // Consolida códigos repetidos somando as quantidades em cada base
    const sapMap = aggregateByMaterialCode(sapRows);
    const revanMap = aggregateByMaterialCode(revanRows);

    // Realiza o cruzamento e salva o resultado no estado
    appState.comparisonResults = compareDatasets(sapMap, revanMap);

    // Atualiza o resumo numérico e exibe a tabela completa
    renderSummary(appState.comparisonResults);
    setActiveFilter("all");
  } catch (error) {
    // Exibe mensagem de erro amigável caso o arquivo esteja corrompido
    showStatusMessage(`Erro ao processar os arquivos: ${error.message}`);
  } finally {
    // Restaura o texto e o estado do botão principal
    compareBtn.textContent = "Comparar Arquivos";
    updateCompareButtonState();
  }
}

// Reseta toda a aplicação para o estado inicial
function handleResetClick() {
  // Limpa os inputs de arquivo
  sapFileInput.value = "";
  revanFileInput.value = "";

  // Zera o estado central
  appState.sapFile = null;
  appState.revanFile = null;
  appState.comparisonResults = [];
  appState.activeFilter = "all";

  // Restaura os textos e estilos dos cards de upload
  sapFileNameDisplay.textContent = "Nenhum arquivo selecionado";
  revanFileNameDisplay.textContent = "Nenhum arquivo selecionado";
  sapUploadCard.classList.remove("has-file");
  revanUploadCard.classList.remove("has-file");

  // Oculta alertas, resumo e tabela
  clearStatusMessage();
  summarySection.classList.add("is-hidden");
  resultsSection.classList.add("is-hidden");
  resultsBody.replaceChildren();

  // Desabilita novamente o botão de comparar
  updateCompareButtonState();
}

// Liga o evento de seleção no input do SAP
sapFileInput.addEventListener("change", () => {
  handleFileSelection("sapFile", sapFileInput, sapUploadCard, sapFileNameDisplay);
});

// Liga o evento de seleção no input da Revan
revanFileInput.addEventListener("change", () => {
  handleFileSelection(
    "revanFile",
    revanFileInput,
    revanUploadCard,
    revanFileNameDisplay
  );
});

// Liga os eventos dos botões principais
compareBtn.addEventListener("click", handleCompareClick);
resetBtn.addEventListener("click", handleResetClick);

// Liga os eventos dos botões de filtro
filterAllBtn.addEventListener("click", () => setActiveFilter("all"));
filterDiffBtn.addEventListener("click", () => setActiveFilter("diff"));

// Liga o evento do botão de exportação para CSV
exportCsvBtn.addEventListener("click", exportDivergencesToCsv);
