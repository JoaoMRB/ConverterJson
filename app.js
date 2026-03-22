const MAX_FILE_SIZE = 10 * 1024 * 1024;
const VALID_EXTENSIONS = [".xlsx", ".xls", ".csv"];
const THEME_KEY = "json-converter-theme";

const state = {
    workbooks: [],
    selectedFileIndex: 0,
    selectedSheetName: "",
    lastOutput: "",
    lastDownloadName: "conversao",
    previewRows: [],
    previewColumns: []
};

const elements = {
    inputFile: document.getElementById("inputFile"),
    dropZone: document.getElementById("dropZone"),
    fileSummary: document.getElementById("fileSummary"),
    fileSelector: document.getElementById("fileSelector"),
    statusBadge: document.getElementById("statusBadge"),
    statFiles: document.getElementById("statFiles"),
    statRows: document.getElementById("statRows"),
    statCols: document.getElementById("statCols"),
    statSheet: document.getElementById("statSheet"),
    progressPanel: document.getElementById("progressPanel"),
    progressBar: document.getElementById("progressBar"),
    progressLabel: document.getElementById("progressLabel"),
    convertButton: document.getElementById("convertButton"),
    jsonToExcelButton: document.getElementById("jsonToExcelButton"),
    resetButton: document.getElementById("resetButton"),
    sheetSelector: document.getElementById("sheetSelector"),
    structureMode: document.getElementById("structureMode"),
    containerMode: document.getElementById("containerMode"),
    keyField: document.getElementById("keyField"),
    allSheetsMode: document.getElementById("allSheetsMode"),
    prettyPrint: document.getElementById("prettyPrint"),
    minifyOutput: document.getElementById("minifyOutput"),
    apiEnvelope: document.getElementById("apiEnvelope"),
    previewTable: document.getElementById("previewTable"),
    output: document.getElementById("output"),
    outputWrapper: document.getElementById("outputWrapper"),
    copyButton: document.getElementById("copyButton"),
    downloadButton: document.getElementById("downloadButton"),
    expandButton: document.getElementById("expandButton"),
    outputEditor: document.getElementById("outputEditor"),
    downloadExcelButton: document.getElementById("downloadExcelButton"),
    loadOutputToEditor: document.getElementById("loadOutputToEditor"),
    themeToggle: document.getElementById("themeToggle"),
    themeIcon: document.getElementById("themeIcon"),
    appToast: document.getElementById("appToast"),
    toastBody: document.getElementById("toastBody")
};

const appToast = new bootstrap.Toast(elements.appToast, { delay: 2600 });

function escapeHtml(value) {
    return String(value)
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
}

function showToast(message, type = "success", icon = "bi-check-circle-fill") {
    elements.appToast.className = "toast align-items-center border-0";
    elements.appToast.classList.add(
        type === "danger" ? "text-bg-danger" : type === "warning" ? "text-bg-warning" : "text-bg-success"
    );
    elements.toastBody.innerHTML = `<i class="bi ${icon}"></i><span>${escapeHtml(message)}</span>`;
    appToast.show();
}

function setTheme(theme) {
    document.documentElement.setAttribute("data-bs-theme", theme);
    localStorage.setItem(THEME_KEY, theme);
    elements.themeIcon.className = theme === "dark" ? "bi bi-sun-fill" : "bi bi-moon-stars-fill";
}

function initTheme() {
    setTheme(localStorage.getItem(THEME_KEY) || "dark");
}

function updateDropZoneState(mode) {
    elements.dropZone.classList.remove("is-valid", "is-invalid");

    if (mode === "valid") {
        elements.dropZone.classList.add("is-valid");
    }

    if (mode === "invalid") {
        elements.dropZone.classList.add("is-invalid");
    }
}

function updateProgress(percent, label) {
    const safePercent = Math.max(0, Math.min(100, Math.round(percent)));
    elements.progressPanel.classList.remove("d-none");
    elements.progressBar.style.width = `${safePercent}%`;
    elements.progressBar.textContent = `${safePercent}%`;
    elements.progressBar.setAttribute("aria-valuenow", String(safePercent));
    elements.progressLabel.textContent = label;
}

function hideProgress() {
    elements.progressPanel.classList.add("d-none");
    elements.progressBar.style.width = "0%";
    elements.progressBar.textContent = "0%";
    elements.progressBar.setAttribute("aria-valuenow", "0");
    elements.progressLabel.textContent = "A ler ficheiro...";
}

function setStatus(label, tone = "secondary") {
    elements.statusBadge.className = `badge text-bg-${tone}`;
    elements.statusBadge.textContent = label;
}

function getExtension(fileName) {
    return `.${fileName.split(".").pop().toLowerCase()}`;
}

function validateFile(file) {
    const extension = getExtension(file.name);

    if (!VALID_EXTENSIONS.includes(extension)) {
        return { valid: false, message: "Ficheiro nao suportado." };
    }

    if (file.size > MAX_FILE_SIZE) {
        return { valid: false, message: "Ficheiro demasiado grande. Maximo 10MB." };
    }

    return { valid: true };
}

function readFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();

        reader.onerror = () => reject(new Error("Falha ao ler o ficheiro."));
        reader.onprogress = (event) => {
            if (event.lengthComputable) {
                updateProgress((event.loaded / event.total) * 100, `A ler ${file.name}...`);
            }
        };
        reader.onload = () => resolve(reader.result);

        if (getExtension(file.name) === ".csv") {
            reader.readAsText(file, "utf-8");
            return;
        }

        reader.readAsArrayBuffer(file);
    });
}

function readWorkbookFromFile(file, rawContent) {
    try {
        if (getExtension(file.name) === ".csv") {
            return XLSX.read(rawContent, { type: "string" });
        }

        return XLSX.read(rawContent, { type: "array" });
    } catch (error) {
        throw new Error(`Erro a ler ficheiro: ${error.message}`);
    }
}

function getRowsForSheet(workbookEntry, sheetName) {
    if (!workbookEntry || !sheetName) {
        return [];
    }

    const sheet = workbookEntry.workbook.Sheets[sheetName];
    return XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
}

function populateFileSelector() {
    elements.fileSelector.innerHTML = "";

    if (!state.workbooks.length) {
        elements.fileSelector.innerHTML = '<option value="0">Escolhe um ficheiro</option>';
        elements.fileSelector.disabled = true;
        return;
    }

    state.workbooks.forEach((entry, index) => {
        const option = document.createElement("option");
        option.value = String(index);
        option.textContent = entry.fileName;
        option.selected = index === state.selectedFileIndex;
        elements.fileSelector.appendChild(option);
    });

    elements.fileSelector.disabled = false;
}

function populateKeyFieldOptions(rows) {
    const columns = rows.length ? Object.keys(rows[0]) : [];
    elements.keyField.innerHTML = '<option value="">Auto (index)</option>';

    columns.forEach((column) => {
        const option = document.createElement("option");
        option.value = column;
        option.textContent = column;
        elements.keyField.appendChild(option);
    });

    elements.keyField.disabled = !columns.length;
}

function populateSheetSelector() {
    const workbookEntry = state.workbooks[state.selectedFileIndex];
    const sheetNames = workbookEntry ? workbookEntry.workbook.SheetNames : [];

    elements.sheetSelector.innerHTML = "";
    elements.keyField.innerHTML = '<option value="">Auto (index)</option>';

    if (!sheetNames.length) {
        elements.sheetSelector.innerHTML = '<option value="">Escolhe uma sheet</option>';
        elements.sheetSelector.disabled = true;
        elements.keyField.disabled = true;
        return;
    }

    sheetNames.forEach((sheetName, index) => {
        const option = document.createElement("option");
        option.value = sheetName;
        option.textContent = sheetName;
        option.selected = index === 0;
        elements.sheetSelector.appendChild(option);
    });

    state.selectedSheetName = sheetNames[0];
    elements.sheetSelector.disabled = false;
    populateKeyFieldOptions(getRowsForSheet(workbookEntry, state.selectedSheetName));
}

function updateStats(rows, sheetName) {
    const columns = rows.length ? Object.keys(rows[0]).length : 0;
    elements.statRows.textContent = String(rows.length);
    elements.statCols.textContent = String(columns);
    elements.statSheet.textContent = sheetName || "-";
    elements.fileSummary.textContent = state.workbooks.map((entry) => entry.fileName).join(", ");
}

function renderEmptyPreview() {
    elements.previewTable.innerHTML = [
        '<thead class="table-dark"><tr><th>Sem dados</th></tr></thead>',
        '<tbody><tr><td class="text-soft">Carrega um ficheiro para ver o preview.</td></tr></tbody>'
    ].join("");
    elements.statRows.textContent = "0";
    elements.statCols.textContent = "0";
    elements.statSheet.textContent = "-";
}

function renderPreviewTable() {
    if (!state.previewColumns.length) {
        renderEmptyPreview();
        return;
    }

    const thead = `<thead class="table-dark"><tr>${state.previewColumns
        .map((column) => `<th>${escapeHtml(column)}</th>`)
        .join("")}</tr></thead>`;

    const tbody = state.previewRows
        .map((row, rowIndex) => {
            const cells = state.previewColumns
                .map((column) => {
                    const value = row[column] == null ? "" : String(row[column]);
                    return `<td contenteditable="true" data-row="${rowIndex}" data-column="${escapeHtml(column)}">${escapeHtml(value)}</td>`;
                })
                .join("");
            return `<tr>${cells}</tr>`;
        })
        .join("");

    elements.previewTable.innerHTML = `${thead}<tbody>${tbody}</tbody>`;
}

function refreshSheetPreview() {
    const workbookEntry = state.workbooks[state.selectedFileIndex];

    if (!workbookEntry) {
        renderEmptyPreview();
        return;
    }

    const sheetName = elements.sheetSelector.value || workbookEntry.workbook.SheetNames[0];
    const rows = getRowsForSheet(workbookEntry, sheetName);

    state.selectedSheetName = sheetName;
    state.previewRows = rows.slice(0, 10).map((row) => ({ ...row }));
    state.previewColumns = rows.length ? Object.keys(rows[0]) : [];

    populateKeyFieldOptions(rows);
    renderPreviewTable();
    updateStats(rows, sheetName);
}

function setByPath(target, path, value) {
    const keys = path.split(".");
    let current = target;

    keys.forEach((key, index) => {
        if (index === keys.length - 1) {
            current[key] = value;
            return;
        }

        if (typeof current[key] !== "object" || current[key] === null || Array.isArray(current[key])) {
            current[key] = {};
        }

        current = current[key];
    });
}

function applyStructure(row, mode) {
    if (mode !== "nested") {
        return { ...row };
    }

    const nested = {};

    Object.entries(row).forEach(([key, value]) => {
        if (key.includes(".")) {
            setByPath(nested, key, value);
            return;
        }

        nested[key] = value;
    });

    return nested;
}

function applyContainer(rows) {
    if (elements.containerMode.value === "array") {
        return rows;
    }

    const keyField = elements.keyField.value;

    return rows.reduce((accumulator, row, index) => {
        const rawKey = keyField ? row[keyField] : index;
        const key = rawKey === "" || rawKey == null ? `item_${index}` : String(rawKey);
        accumulator[key] = row;
        return accumulator;
    }, {});
}

function decorateOutput(payload, meta) {
    if (!elements.apiEnvelope.checked) {
        return payload;
    }

    return { meta, data: payload };
}

function buildSheetPayload(workbookEntry, sheetName, preferPreview = false) {
    const rows = getRowsForSheet(workbookEntry, sheetName);
    const mergedRows = preferPreview && sheetName === state.selectedSheetName
        ? rows.map((row, index) => (index < state.previewRows.length ? { ...row, ...state.previewRows[index] } : row))
        : rows;

    return mergedRows.map((row) => applyStructure(row, elements.structureMode.value));
}

function stringifyOutput(data) {
    const spacing = elements.minifyOutput.checked ? 0 : elements.prettyPrint.checked ? 2 : 0;
    return JSON.stringify(data, null, spacing);
}

function highlightOutput() {
    if (typeof hljs.highlightElement === "function") {
        hljs.highlightElement(elements.output);
    }
}

function updateOutput(jsonString, filename) {
    state.lastOutput = jsonString;
    state.lastDownloadName = filename;
    elements.output.textContent = jsonString || '{\n  "status": "Sem resultado"\n}';
    elements.output.removeAttribute("data-highlighted");
    highlightOutput();
    elements.copyButton.disabled = !jsonString;
    elements.downloadButton.disabled = !jsonString;
}

function convertToJson() {
    if (!state.workbooks.length) {
        showToast("Carrega primeiro um ficheiro.", "warning", "bi-exclamation-triangle-fill");
        return;
    }

    const selectedWorkbook = state.workbooks[state.selectedFileIndex];
    const targetSheets = elements.allSheetsMode.checked
        ? selectedWorkbook.workbook.SheetNames
        : [elements.sheetSelector.value];

    const result = elements.allSheetsMode.checked
        ? targetSheets.reduce((accumulator, sheetName) => {
            const mappedRows = buildSheetPayload(selectedWorkbook, sheetName, true);
            accumulator[sheetName] = decorateOutput(applyContainer(mappedRows), {
                rows: mappedRows.length,
                columns: mappedRows[0] ? Object.keys(mappedRows[0]).length : 0,
                sheet: sheetName,
                file: selectedWorkbook.fileName
            });
            return accumulator;
        }, {})
        : (() => {
            const mappedRows = buildSheetPayload(selectedWorkbook, targetSheets[0], true);
            return decorateOutput(applyContainer(mappedRows), {
                rows: mappedRows.length,
                columns: mappedRows[0] ? Object.keys(mappedRows[0]).length : 0,
                sheet: targetSheets[0],
                file: selectedWorkbook.fileName
            });
        })();

    updateOutput(stringifyOutput(result), selectedWorkbook.baseName);
    showToast("JSON gerado com sucesso.");
}

function downloadTextFile(content, filename, mimeType) {
    const blob = new Blob([content], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");

    link.href = url;
    link.download = filename;
    link.click();

    URL.revokeObjectURL(url);
}

function downloadJson() {
    if (!state.lastOutput) {
        showToast("Ainda nao existe JSON para baixar.", "warning", "bi-exclamation-triangle-fill");
        return;
    }

    downloadTextFile(state.lastOutput, `${state.lastDownloadName}.json`, "application/json");
    showToast("JSON descarregado.");
}

async function copyOutput() {
    if (!state.lastOutput) {
        return;
    }

    try {
        await navigator.clipboard.writeText(state.lastOutput);
        showToast("JSON copiado para a area de transferencia.");
    } catch (error) {
        showToast("Nao foi possivel copiar automaticamente.", "danger", "bi-x-octagon-fill");
    }
}

function toggleOutputExpand() {
    elements.outputWrapper.classList.toggle("is-maximized");
    const maximized = elements.outputWrapper.classList.contains("is-maximized");

    elements.expandButton.innerHTML = maximized
        ? '<i class="bi bi-fullscreen-exit"></i> Fechar'
        : '<i class="bi bi-arrows-fullscreen"></i> Expandir';
}

function resetState(showFeedback = true) {
    state.workbooks = [];
    state.selectedFileIndex = 0;
    state.selectedSheetName = "";
    state.lastOutput = "";
    state.lastDownloadName = "conversao";
    state.previewRows = [];
    state.previewColumns = [];

    elements.fileSummary.textContent = "Nenhum ficheiro selecionado.";
    elements.fileSelector.innerHTML = '<option value="0">Escolhe um ficheiro</option>';
    elements.fileSelector.disabled = true;
    elements.sheetSelector.innerHTML = '<option value="">Escolhe uma sheet</option>';
    elements.sheetSelector.disabled = true;
    elements.keyField.innerHTML = '<option value="">Auto (index)</option>';
    elements.keyField.disabled = true;
    elements.convertButton.disabled = true;
    elements.copyButton.disabled = true;
    elements.downloadButton.disabled = true;
    elements.inputFile.value = "";
    elements.outputEditor.value = "";
    renderEmptyPreview();
    updateOutput("", "conversao");
    updateDropZoneState();
    setStatus("Aguardando", "secondary");
    elements.statFiles.textContent = "0";

    if (showFeedback) {
        showToast("Interface limpa e pronta para novo ficheiro.");
    }
}

function loadOutputIntoEditor() {
    elements.outputEditor.value = state.lastOutput;
    showToast("Output enviado para o editor.");
}

function normalizeJsonInput(input) {
    const parsed = JSON.parse(input);

    if (Array.isArray(parsed)) {
        return parsed;
    }

    if (parsed && Array.isArray(parsed.data)) {
        return parsed.data;
    }

    if (parsed && typeof parsed === "object") {
        return Object.values(parsed);
    }

    throw new Error("O JSON precisa de ser um array ou objeto.");
}

function jsonToExcel() {
    const raw = elements.outputEditor.value.trim();

    if (!raw) {
        showToast("Cola JSON valido antes de gerar Excel.", "warning", "bi-exclamation-triangle-fill");
        return;
    }

    try {
        const normalized = normalizeJsonInput(raw);
        const sheet = XLSX.utils.json_to_sheet(normalized);
        const workbook = XLSX.utils.book_new();

        XLSX.utils.book_append_sheet(workbook, sheet, "Dados");
        XLSX.writeFile(workbook, "json-convertido.xlsx");

        showToast("Excel gerado a partir do JSON.");
    } catch (error) {
        showToast(`Erro ao converter JSON para Excel: ${error.message}`, "danger", "bi-x-octagon-fill");
    }
}

function syncPrettyMinify(source) {
    if (source === "pretty" && elements.prettyPrint.checked) {
        elements.minifyOutput.checked = false;
    }

    if (source === "minify" && elements.minifyOutput.checked) {
        elements.prettyPrint.checked = false;
    }
}

async function handleFiles(fileList) {
    const files = Array.from(fileList || []);

    if (!files.length) {
        return;
    }

    resetState(false);
    elements.statFiles.textContent = String(files.length);
    setStatus("A processar", "info");

    for (const file of files) {
        const validation = validateFile(file);
        if (!validation.valid) {
            updateDropZoneState("invalid");
            hideProgress();
            setStatus("Erro", "danger");
            showToast(validation.message, "danger", "bi-x-octagon-fill");
            return;
        }
    }

    updateDropZoneState("valid");

    try {
        const loaded = [];

        for (let index = 0; index < files.length; index += 1) {
            const file = files[index];
            updateProgress((index / files.length) * 100, `A preparar ${file.name}...`);
            const content = await readFile(file);
            updateProgress(((index + 0.6) / files.length) * 100, `A interpretar ${file.name}...`);
            const workbook = readWorkbookFromFile(file, content);

            loaded.push({
                fileName: file.name,
                baseName: file.name.replace(/\.[^.]+$/, ""),
                workbook
            });

            updateProgress(((index + 1) / files.length) * 100, `${file.name} carregado.`);
        }

        state.workbooks = loaded;
        state.selectedFileIndex = 0;
        populateFileSelector();
        populateSheetSelector();
        refreshSheetPreview();
        setStatus("Pronto", "success");
        elements.convertButton.disabled = false;
        showToast(files.length > 1 ? "Ficheiros carregados com sucesso." : "Ficheiro carregado com sucesso.");
    } catch (error) {
        updateDropZoneState("invalid");
        setStatus("Erro", "danger");
        showToast(error.message, "danger", "bi-x-octagon-fill");
    } finally {
        hideProgress();
    }
}

elements.dropZone.addEventListener("dragover", (event) => {
    event.preventDefault();
    elements.dropZone.classList.add("is-dragover");
});

elements.dropZone.addEventListener("dragleave", () => {
    elements.dropZone.classList.remove("is-dragover");
});

elements.dropZone.addEventListener("drop", (event) => {
    event.preventDefault();
    elements.dropZone.classList.remove("is-dragover");
    handleFiles(event.dataTransfer.files);
});

elements.inputFile.addEventListener("change", (event) => {
    handleFiles(event.target.files);
});

elements.fileSelector.addEventListener("change", (event) => {
    state.selectedFileIndex = Number(event.target.value) || 0;
    populateSheetSelector();
    refreshSheetPreview();
});

elements.sheetSelector.addEventListener("change", refreshSheetPreview);
elements.convertButton.addEventListener("click", convertToJson);
elements.downloadButton.addEventListener("click", downloadJson);
elements.copyButton.addEventListener("click", copyOutput);
elements.expandButton.addEventListener("click", toggleOutputExpand);
elements.resetButton.addEventListener("click", () => resetState(true));
elements.jsonToExcelButton.addEventListener("click", loadOutputIntoEditor);
elements.downloadExcelButton.addEventListener("click", jsonToExcel);
elements.loadOutputToEditor.addEventListener("click", loadOutputIntoEditor);

elements.themeToggle.addEventListener("click", () => {
    const currentTheme = document.documentElement.getAttribute("data-bs-theme");
    setTheme(currentTheme === "dark" ? "light" : "dark");
});

elements.previewTable.addEventListener("input", (event) => {
    const target = event.target;

    if (!(target instanceof HTMLElement) || target.tagName !== "TD") {
        return;
    }

    const rowIndex = Number(target.dataset.row);
    const column = target.dataset.column;

    if (Number.isNaN(rowIndex) || !column || !state.previewRows[rowIndex]) {
        return;
    }

    state.previewRows[rowIndex][column] = target.textContent || "";
});

elements.prettyPrint.addEventListener("change", () => syncPrettyMinify("pretty"));
elements.minifyOutput.addEventListener("change", () => syncPrettyMinify("minify"));

elements.allSheetsMode.addEventListener("change", () => {
    elements.sheetSelector.disabled = elements.allSheetsMode.checked || !state.workbooks.length;
});

window.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && elements.outputWrapper.classList.contains("is-maximized")) {
        toggleOutputExpand();
    }
});

document.addEventListener("DOMContentLoaded", () => {
    initTheme();
    updateOutput("", "conversao");
    hideProgress();
    setStatus("Aguardando", "secondary");
    renderEmptyPreview();

    document.querySelectorAll('[data-bs-toggle="tooltip"]').forEach((element) => {
        new bootstrap.Tooltip(element);
    });
});
