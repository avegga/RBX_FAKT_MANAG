const emptyDataset = () => ({
    file_name: "",
    columns: [],
    rows: [],
    warnings: [],
    errors: [],
    partial: false,
    missing_columns: [],
    row_count: 0,
    template: null,
});

const defaultTableSettings = () => ({
    hideMoneyCents: false,
    numbersAsIntegers: false,
    rowLimit: 100,
});

const emptyDerivedState = () => ({
    rows: [],
    filteredRows: [],
    displayRows: [],
    visibleColumns: [],
    orderedColumns: [],
    errors: [],
    errorLookup: new Set(),
});

const state = {
    settings: {},
    templates: [],
    ui: {},
    activeTemplateId: null,
    activeRightTab: "general",
    dataset: emptyDataset(),
    table: {
        settings: defaultTableSettings(),
        columns: [],
        filters: {},
        sort: { columnName: null, direction: null },
        derived: emptyDerivedState(),
    },
};

const columnTypes = ["text", "number", "money", "date", "datetime", "boolean"];
const rowHeight = 44;
const rowOverscan = 8;
let stateSaveTimer = null;

const elements = {
    body: document.body,
    detailsPanel: document.getElementById("details-panel"),
    tabButtons: Array.from(document.querySelectorAll(".tab-button")),
    tabPanels: Array.from(document.querySelectorAll(".tab-panel")),
    rightTabButtons: Array.from(document.querySelectorAll(".right-tab-button")),
    rightTabPanels: Array.from(document.querySelectorAll(".right-tab-panel")),
    activeTabTitle: document.getElementById("active-tab-title"),
    saveStatus: document.getElementById("save-status"),
    serviceMessage: document.getElementById("service-message"),
    settingsForm: document.getElementById("settings-form"),
    settingsErrors: document.getElementById("settings-errors"),
    leftPanelWidth: document.getElementById("left-panel-width"),
    rightPanelVisible: document.getElementById("right-panel-visible"),
    errorsVisible: document.getElementById("errors-visible"),
    errorsHeight: document.getElementById("errors-height"),
    lastLoadedFile: document.getElementById("last-loaded-file"),
    templateList: document.getElementById("template-list"),
    templateForm: document.getElementById("template-form"),
    templateId: document.getElementById("template-id"),
    templateName: document.getElementById("template-name"),
    columnsContainer: document.getElementById("columns-container"),
    templateErrors: document.getElementById("template-errors"),
    newTemplateButton: document.getElementById("new-template-button"),
    addColumnButton: document.getElementById("add-column-button"),
    deleteTemplateButton: document.getElementById("delete-template-button"),
    journalList: document.getElementById("journal-list"),
    lastTemplateLabel: document.getElementById("last-template-label"),
    lastFileLabel: document.getElementById("last-file-label"),
    rightPanelLabel: document.getElementById("right-panel-label"),
    errorsPanelLabel: document.getElementById("errors-panel-label"),
    templateCount: document.getElementById("template-count"),
    columnCount: document.getElementById("column-count"),
    errorsPanelMessage: document.getElementById("errors-panel-message"),
    issuesSummary: document.getElementById("issues-summary"),
    issuesList: document.getElementById("issues-list"),
    uploadTemplateSelect: document.getElementById("upload-template-select"),
    currentFileName: document.getElementById("current-file-name"),
    uploadButton: document.getElementById("upload-button"),
    exportButton: document.getElementById("export-button"),
    toggleRightPanelButton: document.getElementById("toggle-right-panel-button"),
    displayAllToggle: document.getElementById("display-all-toggle"),
    csvFileInput: document.getElementById("csv-file-input"),
    datasetStats: document.getElementById("dataset-stats"),
    dataEmptyState: document.getElementById("data-empty-state"),
    dataTableShell: document.getElementById("data-table-shell"),
    dataTableHeader: document.getElementById("data-table-header"),
    dataTableViewport: document.getElementById("data-table-viewport"),
    dataTableSpacerTop: document.getElementById("data-table-spacer-top"),
    dataTableRows: document.getElementById("data-table-rows"),
    dataTableSpacerBottom: document.getElementById("data-table-spacer-bottom"),
    generalPanelContent: document.getElementById("general-panel-content"),
    columnsPanelContent: document.getElementById("columns-panel-content"),
    filtersPanelContent: document.getElementById("filters-panel-content"),
    typesPanelContent: document.getElementById("types-panel-content"),
    processingPanelContent: document.getElementById("processing-panel-content"),
};

async function requestJson(url, options = {}) {
    const isFormData = options.body instanceof FormData;
    const headers = isFormData ? {} : { "Content-Type": "application/json" };
    const response = await fetch(url, {
        headers: { ...headers, ...(options.headers || {}) },
        ...options,
    });

    const data = await response.json().catch(() => ({}));
    if (!response.ok) {
        throw data;
    }

    return data;
}

function escapeHtml(value) {
    return String(value ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;");
}

function setStatus(message, level = "info") {
    elements.saveStatus.textContent = message;
    elements.saveStatus.dataset.level = level;
    if (elements.serviceMessage) {
        elements.serviceMessage.textContent = message;
        elements.serviceMessage.className = `service-message ${level}`;
    }
}

function switchTab(tabName) {
    elements.tabButtons.forEach((button) => {
        button.classList.toggle("active", button.dataset.tab === tabName);
    });

    elements.tabPanels.forEach((panel) => {
        panel.classList.toggle("active", panel.dataset.panel === tabName);
    });

    const activeButton = elements.tabButtons.find((button) => button.dataset.tab === tabName);
    elements.activeTabTitle.textContent = activeButton ? activeButton.textContent : "Обзор";
}

function switchRightTab(tabName) {
    state.activeRightTab = tabName;
    elements.rightTabButtons.forEach((button) => {
        button.classList.toggle("active", button.dataset.rightTab === tabName);
    });
    elements.rightTabPanels.forEach((panel) => {
        panel.classList.toggle("active", panel.dataset.rightPanel === tabName);
    });
}

function getActiveTemplate() {
    return state.templates.find((item) => item.id === state.activeTemplateId) || null;
}

function createFilterState(type) {
    if (type === "text") {
        return { value: "" };
    }
    if (type === "number" || type === "money") {
        return { operator: "contains", value: "", from: "", to: "" };
    }
    if (type === "date" || type === "datetime") {
        return { from: "", to: "" };
    }
    if (type === "boolean") {
        return { value: "any" };
    }
    return { value: "" };
}

function resetTableModel(preserveSettings = true) {
    const settings = preserveSettings ? state.table.settings : defaultTableSettings();
    state.table = {
        settings,
        columns: [],
        filters: {},
        sort: { columnName: null, direction: null },
        derived: emptyDerivedState(),
    };
}

function initializeTableModel() {
    if (!state.dataset.columns.length) {
        resetTableModel(true);
        return;
    }

    state.table.columns = state.dataset.columns.map((column, index) => ({
        name: column.name,
        type: column.type,
        visible: true,
        position: index,
        fromTemplate: column.from_template,
    }));
    state.table.filters = Object.fromEntries(
        state.table.columns.map((column) => [column.name, createFilterState(column.type)])
    );
    state.table.sort = { columnName: null, direction: null };
    recomputeTableState();
}

function syncSettingsForm() {
    elements.settingsForm.source_folder.value = state.settings.source_folder || "";
    elements.settingsForm.output_folder.value = state.settings.output_folder || "";
    elements.settingsForm.instructions_file.value = state.settings.instructions_file || "";
}

function applyUiState() {
    const ui = state.ui;
    document.documentElement.style.setProperty("--sidebar-width", `${ui.left_panel_width || 280}px`);
    document.documentElement.style.setProperty("--errors-height", `${ui.errors_height || 180}px`);
    elements.body.classList.toggle("hide-right-panel", !ui.right_panel_visible);
    elements.body.classList.toggle("hide-errors-panel", !ui.errors_visible);

    elements.leftPanelWidth.value = ui.left_panel_width || 280;
    elements.rightPanelVisible.checked = Boolean(ui.right_panel_visible);
    elements.errorsVisible.checked = Boolean(ui.errors_visible);
    elements.errorsHeight.value = ui.errors_height || 180;
    elements.lastLoadedFile.value = ui.last_loaded_file || "";
    elements.toggleRightPanelButton.textContent = ui.right_panel_visible ? "Скрыть панель" : "Показать панель";

    syncStateLabels();
}

function syncStateLabels() {
    const activeTemplate = getActiveTemplate();
    elements.lastTemplateLabel.textContent = activeTemplate ? activeTemplate.name : "Не выбран";
    elements.lastFileLabel.textContent = state.ui.last_loaded_file || "Не задан";
    elements.rightPanelLabel.textContent = state.ui.right_panel_visible ? "Включена" : "Скрыта";
    elements.errorsPanelLabel.textContent = state.ui.errors_visible ? "Включена" : "Скрыта";
}

function renderTemplateList() {
    elements.templateList.innerHTML = "";
    elements.templateCount.textContent = String(state.templates.length);

    if (!state.templates.length) {
        const item = document.createElement("li");
        item.textContent = "Шаблоны пока не созданы.";
        elements.templateList.appendChild(item);
        return;
    }

    state.templates.forEach((template) => {
        const item = document.createElement("li");
        const button = document.createElement("button");
        button.type = "button";
        button.textContent = template.name;
        button.classList.toggle("active", template.id === state.activeTemplateId);
        button.addEventListener("click", () => setActiveTemplate(template.id));
        item.appendChild(button);
        elements.templateList.appendChild(item);
    });
}

function createColumnRow(column = { name: "", type: "text" }) {
    const row = document.createElement("div");
    row.className = "column-row";

    const nameInput = document.createElement("input");
    nameInput.type = "text";
    nameInput.placeholder = "Название столбца";
    nameInput.value = column.name || "";

    const typeSelect = document.createElement("select");
    columnTypes.forEach((type) => {
        const option = document.createElement("option");
        option.value = type;
        option.textContent = type;
        option.selected = type === (column.type || "text");
        typeSelect.appendChild(option);
    });

    const removeButton = document.createElement("button");
    removeButton.type = "button";
    removeButton.className = "icon-button";
    removeButton.textContent = "×";
    removeButton.addEventListener("click", () => {
        row.remove();
        updateTemplateColumnCount();
    });

    row.append(nameInput, typeSelect, removeButton);
    return row;
}

function getTemplateFormPayload() {
    const columns = Array.from(elements.columnsContainer.querySelectorAll(".column-row")).map((row) => {
        const [nameInput, typeSelect] = row.querySelectorAll("input, select");
        return {
            name: nameInput.value.trim(),
            type: typeSelect.value,
        };
    });

    return {
        name: elements.templateName.value.trim(),
        columns,
    };
}

function fillTemplateForm(template) {
    elements.templateId.value = template?.id || "";
    elements.templateName.value = template?.name || "";
    elements.columnsContainer.innerHTML = "";

    const columns = template?.columns?.length ? template.columns : [{ name: "", type: "text" }];
    columns.forEach((column) => {
        elements.columnsContainer.appendChild(createColumnRow(column));
    });

    elements.deleteTemplateButton.disabled = !template;
    updateTemplateColumnCount();
}

function updateTemplateColumnCount() {
    const activeDatasetColumns = state.table.derived.visibleColumns.length || state.table.columns.length;
    elements.columnCount.textContent = String(activeDatasetColumns || elements.columnsContainer.querySelectorAll(".column-row").length);
}

function setActiveTemplate(templateId, options = {}) {
    const template = state.templates.find((item) => item.id === templateId) || null;
    const persist = options.persist !== false;
    state.activeTemplateId = template ? template.id : null;
    fillTemplateForm(template);
    renderTemplateList();
    renderUploadTemplateOptions();
    renderRightPanels();
    syncStateLabels();
    if (persist) {
        scheduleStateSave({ last_template_id: state.activeTemplateId }, { notify: false });
    }
}

function renderUploadTemplateOptions() {
    elements.uploadTemplateSelect.innerHTML = "";

    const emptyOption = document.createElement("option");
    emptyOption.value = "";
    emptyOption.textContent = "Без шаблона";
    elements.uploadTemplateSelect.appendChild(emptyOption);

    state.templates.forEach((template) => {
        const option = document.createElement("option");
        option.value = String(template.id);
        option.textContent = template.name;
        elements.uploadTemplateSelect.appendChild(option);
    });

    elements.uploadTemplateSelect.value = state.activeTemplateId ? String(state.activeTemplateId) : "";
    elements.currentFileName.textContent = state.dataset.file_name || state.ui.last_loaded_file || "Файл не выбран";
}

async function loadJournal() {
    const entries = await requestJson("/api/logs");
    elements.journalList.innerHTML = "";

    if (!entries.length) {
        elements.journalList.textContent = "Событий пока нет.";
        return;
    }

    entries.forEach((entry) => {
        const item = document.createElement("div");
        item.className = "journal-item";
        const title = document.createElement("strong");
        title.textContent = entry.event_type;
        const text = document.createElement("span");
        text.textContent = entry.event_description;
        item.append(title, text);
        elements.journalList.appendChild(item);
    });
}

function getOrderedColumns() {
    return [...state.table.columns].sort((left, right) => left.position - right.position);
}

function getVisibleColumns() {
    return getOrderedColumns().filter((column) => column.visible);
}

function getColumnState(columnName) {
    return state.table.columns.find((column) => column.name === columnName) || null;
}

function parseStrictNumber(rawValue) {
    const normalized = String(rawValue ?? "").trim();
    if (!normalized) {
        return null;
    }
    if (!/^[+-]?(?:\d+|\d{1,3}(?: \d{3})+)(?:\.\d+)?$/.test(normalized)) {
        throw new Error("invalid-number");
    }
    return Number(normalized.replaceAll(" ", ""));
}

function createValidDate(year, month, day, hours = 0, minutes = 0, seconds = 0) {
    const date = new Date(year, month - 1, day, hours, minutes, seconds, 0);
    if (
        date.getFullYear() !== year ||
        date.getMonth() !== month - 1 ||
        date.getDate() !== day ||
        date.getHours() !== hours ||
        date.getMinutes() !== minutes ||
        date.getSeconds() !== seconds
    ) {
        throw new Error("invalid-date");
    }
    return date;
}

function parseDateStrict(rawValue) {
    const normalized = String(rawValue ?? "").trim();
    if (!normalized) {
        return null;
    }
    let match = normalized.match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
    if (match) {
        return createValidDate(Number(match[3]), Number(match[2]), Number(match[1]));
    }
    match = normalized.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (match) {
        return createValidDate(Number(match[1]), Number(match[2]), Number(match[3]));
    }
    throw new Error("invalid-date");
}

function parseDateTimeStrict(rawValue) {
    const normalized = String(rawValue ?? "").trim();
    if (!normalized) {
        return null;
    }
    let match = normalized.match(/^(\d{2})\.(\d{2})\.(\d{4}) (\d{2}):(\d{2})$/);
    if (match) {
        return createValidDate(
            Number(match[3]),
            Number(match[2]),
            Number(match[1]),
            Number(match[4]),
            Number(match[5]),
            0
        );
    }
    match = normalized.match(/^(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2})$/);
    if (match) {
        return createValidDate(
            Number(match[1]),
            Number(match[2]),
            Number(match[3]),
            Number(match[4]),
            Number(match[5]),
            0
        );
    }
    match = normalized.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})$/);
    if (match) {
        return createValidDate(
            Number(match[1]),
            Number(match[2]),
            Number(match[3]),
            Number(match[4]),
            Number(match[5]),
            Number(match[6])
        );
    }
    throw new Error("invalid-datetime");
}

function parseBooleanStrict(rawValue) {
    const normalized = String(rawValue ?? "").trim().toLowerCase();
    if (!normalized) {
        return null;
    }
    if (["true", "да", "1", "истина"].includes(normalized)) {
        return true;
    }
    if (["false", "нет", "0", "ложь"].includes(normalized)) {
        return false;
    }
    throw new Error("invalid-boolean");
}

function formatDate(date) {
    return `${String(date.getDate()).padStart(2, "0")}.${String(date.getMonth() + 1).padStart(2, "0")}.${date.getFullYear()}`;
}

function formatDateTime(date) {
    return `${formatDate(date)} ${String(date.getHours()).padStart(2, "0")}:${String(date.getMinutes()).padStart(2, "0")}`;
}

function formatNumberValue(value, type) {
    let normalized = value;
    if (type === "number" && state.table.settings.numbersAsIntegers) {
        normalized = Math.trunc(normalized);
    }
    if (type === "money" && state.table.settings.hideMoneyCents) {
        normalized = Math.trunc(normalized);
    }
    if (Number.isInteger(normalized)) {
        return String(normalized);
    }
    return String(normalized);
}

function evaluateCell(rawValue, column) {
    const raw = String(rawValue ?? "");
    const trimmed = raw.trim();
    const base = {
        raw,
        display: raw,
        comparable: raw.toLowerCase(),
        empty: trimmed === "",
        normalized: raw,
        error: null,
        type: column.type,
    };

    if (column.type === "text") {
        return base;
    }

    if (column.type === "number" || column.type === "money") {
        if (!trimmed) {
            return {
                ...base,
                normalized: 0,
                comparable: 0,
                display: formatNumberValue(0, column.type),
                error: "пустое значение трактуется как 0",
            };
        }
        try {
            const parsed = parseStrictNumber(trimmed);
            return {
                ...base,
                normalized: parsed,
                comparable: parsed,
                display: formatNumberValue(parsed, column.type),
            };
        } catch {
            return { ...base, comparable: null, error: "ожидалось числовое значение" };
        }
    }

    if (column.type === "date") {
        if (!trimmed) {
            const fallback = createValidDate(1900, 1, 1);
            return {
                ...base,
                normalized: fallback,
                comparable: fallback.getTime(),
                display: formatDate(fallback),
                error: "пустое значение трактуется как 01.01.1900",
            };
        }
        try {
            const parsed = parseDateStrict(trimmed);
            return {
                ...base,
                normalized: parsed,
                comparable: parsed.getTime(),
                display: formatDate(parsed),
            };
        } catch {
            return { ...base, comparable: null, error: "ожидалась дата" };
        }
    }

    if (column.type === "datetime") {
        if (!trimmed) {
            return { ...base, comparable: null, error: "ожидались дата и время" };
        }
        try {
            const parsed = parseDateTimeStrict(trimmed);
            return {
                ...base,
                normalized: parsed,
                comparable: parsed.getTime(),
                dayComparable: createValidDate(parsed.getFullYear(), parsed.getMonth() + 1, parsed.getDate()).getTime(),
                display: formatDateTime(parsed),
            };
        } catch {
            return { ...base, comparable: null, error: "ожидались дата и время" };
        }
    }

    if (column.type === "boolean") {
        if (!trimmed) {
            return {
                ...base,
                normalized: false,
                comparable: 0,
                display: "Нет",
                error: "пустое значение трактуется как Нет",
            };
        }
        try {
            const parsed = parseBooleanStrict(trimmed);
            return {
                ...base,
                normalized: parsed,
                comparable: parsed ? 1 : 0,
                display: parsed ? "Да" : "Нет",
            };
        } catch {
            return { ...base, comparable: null, error: "ожидалось логическое значение" };
        }
    }

    return base;
}

function getDateOnlyComparable(cell) {
    if (cell.type === "date" && cell.normalized instanceof Date) {
        return cell.normalized.getTime();
    }
    if (cell.type === "datetime" && cell.normalized instanceof Date) {
        return createValidDate(
            cell.normalized.getFullYear(),
            cell.normalized.getMonth() + 1,
            cell.normalized.getDate()
        ).getTime();
    }
    return null;
}

function matchesFilter(cell, filter, column) {
    if (!filter) {
        return true;
    }

    if (column.type === "text") {
        const needle = (filter.value || "").trim().toLowerCase();
        return !needle || cell.raw.toLowerCase().includes(needle);
    }

    if (column.type === "number" || column.type === "money") {
        const operator = filter.operator || "contains";
        if (operator === "contains") {
            return true;
        }
        if (cell.comparable === null) {
            return false;
        }
        try {
            if (operator === "eq") {
                if (!(filter.value || "").trim()) {
                    return true;
                }
                return cell.comparable === parseStrictNumber(filter.value);
            }
            if (operator === "gt") {
                if (!(filter.value || "").trim()) {
                    return true;
                }
                return cell.comparable > parseStrictNumber(filter.value);
            }
            if (operator === "lt") {
                if (!(filter.value || "").trim()) {
                    return true;
                }
                return cell.comparable < parseStrictNumber(filter.value);
            }
            if (operator === "between") {
                const fromValue = (filter.from || "").trim() ? parseStrictNumber(filter.from) : null;
                const toValue = (filter.to || "").trim() ? parseStrictNumber(filter.to) : null;
                if (fromValue === null && toValue === null) {
                    return true;
                }
                if (fromValue !== null && cell.comparable < fromValue) {
                    return false;
                }
                if (toValue !== null && cell.comparable > toValue) {
                    return false;
                }
                return true;
            }
        } catch {
            return true;
        }
        return true;
    }

    if (column.type === "date" || column.type === "datetime") {
        const fromRaw = (filter.from || "").trim();
        const toRaw = (filter.to || "").trim();
        if (!fromRaw && !toRaw) {
            return true;
        }
        const comparable = getDateOnlyComparable(cell);
        if (comparable === null) {
            return false;
        }
        try {
            const fromValue = fromRaw ? parseDateStrict(fromRaw).getTime() : null;
            const toValue = toRaw ? parseDateStrict(toRaw).getTime() : null;
            if (fromValue !== null && comparable < fromValue) {
                return false;
            }
            if (toValue !== null && comparable > toValue) {
                return false;
            }
            return true;
        } catch {
            return true;
        }
    }

    if (column.type === "boolean") {
        const mode = filter.value || "any";
        if (mode === "any") {
            return true;
        }
        if (mode === "empty") {
            return cell.empty;
        }
        if (cell.empty) {
            return false;
        }
        return mode === "yes" ? cell.normalized === true : cell.normalized === false;
    }

    return true;
}

function compareRows(leftRow, rightRow, column) {
    const leftCell = leftRow.cells[column.name];
    const rightCell = rightRow.cells[column.name];
    if (leftCell.comparable === null && rightCell.comparable === null) {
        return leftRow.index - rightRow.index;
    }
    if (leftCell.comparable === null) {
        return 1;
    }
    if (rightCell.comparable === null) {
        return -1;
    }
    if (column.type === "text") {
        return leftCell.raw.localeCompare(rightCell.raw, "ru", { sensitivity: "base" }) || leftRow.index - rightRow.index;
    }
    if (leftCell.comparable < rightCell.comparable) {
        return -1;
    }
    if (leftCell.comparable > rightCell.comparable) {
        return 1;
    }
    return leftRow.index - rightRow.index;
}

function recomputeTableState() {
    const orderedColumns = getOrderedColumns();
    const visibleColumns = orderedColumns.filter((column) => column.visible);
    const rows = state.dataset.rows.map((row, index) => {
        const cells = Object.fromEntries(
            orderedColumns.map((column) => [column.name, evaluateCell(row.values[column.name], column)])
        );
        return { ...row, index, cells };
    });

    const errors = [];
    rows.forEach((row) => {
        orderedColumns.forEach((column) => {
            const cell = row.cells[column.name];
            if (cell.error) {
                errors.push({
                    row_number: row.row_number,
                    source_row_number: row.source_row_number,
                    column_name: column.name,
                    column_type: column.type,
                    value: cell.raw,
                    reason: cell.error,
                });
            }
        });
    });

    const filteredRows = rows.filter((row) =>
        orderedColumns.every((column) => matchesFilter(row.cells[column.name], state.table.filters[column.name], column))
    );

    let sortedRows = [...filteredRows];
    if (state.table.sort.columnName && state.table.sort.direction) {
        const sortColumn = orderedColumns.find((column) => column.name === state.table.sort.columnName);
        if (sortColumn) {
            sortedRows.sort((left, right) => compareRows(left, right, sortColumn));
            if (state.table.sort.direction === "desc") {
                sortedRows.reverse();
            }
        }
    }

    const rowLimit = Math.max(1, Number(state.table.settings.rowLimit) || 1);
    const displayRows = elements.displayAllToggle.checked ? sortedRows : sortedRows.slice(0, rowLimit);

    state.table.derived = {
        rows,
        filteredRows: sortedRows,
        displayRows,
        visibleColumns,
        orderedColumns,
        errors,
        errorLookup: new Set(errors.map((item) => `${item.row_number}:${item.column_name}`)),
    };
}

function getGridTemplate() {
    return `72px repeat(${Math.max(1, state.table.derived.visibleColumns.length)}, minmax(160px, 1fr))`;
}

function renderDatasetStats() {
    const totalRows = state.dataset.row_count || 0;
    const filteredRows = state.table.derived.filteredRows.length;
    const shownRows = state.table.derived.displayRows.length;
    const visibleColumns = state.table.derived.visibleColumns.length;
    const warnings = state.dataset.warnings.length;
    const errors = state.table.derived.errors.length;

    elements.datasetStats.innerHTML = `
        <div class="stat-chip">
            <span>Строк всего</span>
            <strong>${totalRows}</strong>
        </div>
        <div class="stat-chip">
            <span>После фильтра</span>
            <strong>${filteredRows}</strong>
        </div>
        <div class="stat-chip">
            <span>Показано</span>
            <strong>${shownRows}</strong>
        </div>
        <div class="stat-chip">
            <span>Видимых столбцов</span>
            <strong>${visibleColumns}</strong>
        </div>
        <div class="stat-chip ${warnings ? "warning" : ""}">
            <span>Предупреждения</span>
            <strong>${warnings}</strong>
        </div>
        <div class="stat-chip ${errors ? "danger" : ""}">
            <span>Ошибки</span>
            <strong>${errors}</strong>
        </div>
    `;
}

function getSortIndicator(columnName) {
    if (state.table.sort.columnName !== columnName || !state.table.sort.direction) {
        return "";
    }
    return state.table.sort.direction === "asc" ? "▲" : "▼";
}

function renderHeader() {
    const visibleColumns = state.table.derived.visibleColumns;
    if (!visibleColumns.length) {
        elements.dataTableHeader.innerHTML = "";
        return;
    }

    const gridTemplate = getGridTemplate();
    elements.dataTableHeader.style.gridTemplateColumns = gridTemplate;
    const headerCells = ['<div class="table-cell index-head">#</div>'];
    visibleColumns.forEach((column) => {
        headerCells.push(`
            <button type="button" class="table-cell header-cell table-sort-button" data-action="sort-column" data-column="${escapeHtml(column.name)}">
                <span>${escapeHtml(column.name)}</span>
                <span class="header-tools">
                    <span class="column-badge">${escapeHtml(column.type)}</span>
                    <span class="sort-indicator">${getSortIndicator(column.name)}</span>
                </span>
            </button>
        `);
    });
    elements.dataTableHeader.innerHTML = headerCells.join("");
}

function renderRows(startIndex, endIndex) {
    const visibleColumns = state.table.derived.visibleColumns;
    const gridTemplate = getGridTemplate();
    const rows = state.table.derived.displayRows.slice(startIndex, endIndex);
    elements.dataTableRows.innerHTML = rows
        .map((row) => {
            const cells = [`<div class="table-cell row-index">${row.row_number}</div>`];
            visibleColumns.forEach((column) => {
                const cell = row.cells[column.name];
                const hasError = state.table.derived.errorLookup.has(`${row.row_number}:${column.name}`);
                cells.push(
                    `<div class="table-cell data-cell ${hasError ? "cell-error" : ""}" title="${escapeHtml(cell.display)}">${escapeHtml(cell.display)}</div>`
                );
            });
            return `<div class="table-row" style="grid-template-columns:${gridTemplate}">${cells.join("")}</div>`;
        })
        .join("");
}

function renderVirtualRows() {
    const displayRows = state.table.derived.displayRows;
    const visibleColumns = state.table.derived.visibleColumns;

    if (!visibleColumns.length) {
        elements.dataTableRows.innerHTML = '<div class="table-row empty-table-row"><div class="table-cell data-cell empty-cell">Все столбцы скрыты. Включите хотя бы один столбец во вкладке Столбцы.</div></div>';
        elements.dataTableSpacerTop.style.height = "0px";
        elements.dataTableSpacerBottom.style.height = "0px";
        return;
    }

    if (!displayRows.length) {
        const message = state.dataset.rows.length
            ? "Нет строк для отображения. Проверьте фильтры или лимит строк."
            : "В файле нет строк данных после заголовка.";
        elements.dataTableRows.innerHTML = `<div class="table-row empty-table-row"><div class="table-cell data-cell empty-cell">${escapeHtml(message)}</div></div>`;
        elements.dataTableSpacerTop.style.height = "0px";
        elements.dataTableSpacerBottom.style.height = "0px";
        return;
    }

    const viewportHeight = elements.dataTableViewport.clientHeight || 480;
    const scrollTop = elements.dataTableViewport.scrollTop;
    const visibleCount = Math.ceil(viewportHeight / rowHeight) + rowOverscan * 2;
    const startIndex = Math.max(0, Math.floor(scrollTop / rowHeight) - rowOverscan);
    const endIndex = Math.min(displayRows.length, startIndex + visibleCount);

    elements.dataTableSpacerTop.style.height = `${startIndex * rowHeight}px`;
    elements.dataTableSpacerBottom.style.height = `${(displayRows.length - endIndex) * rowHeight}px`;
    renderRows(startIndex, endIndex);
}

function renderDataset() {
    renderUploadTemplateOptions();
    renderDatasetStats();

    const hasColumns = Boolean(state.dataset.columns.length);
    elements.dataEmptyState.classList.toggle("hidden", hasColumns);
    elements.dataTableShell.classList.toggle("hidden", !hasColumns);
    elements.currentFileName.textContent = state.dataset.file_name || state.ui.last_loaded_file || "Файл не выбран";

    if (!hasColumns) {
        elements.dataEmptyState.textContent = "Выберите шаблон и загрузите CSV-файл, чтобы увидеть таблицу данных.";
        elements.dataTableRows.innerHTML = "";
        elements.dataTableHeader.innerHTML = "";
        elements.dataTableSpacerTop.style.height = "0px";
        elements.dataTableSpacerBottom.style.height = "0px";
        return;
    }

    renderHeader();
    renderVirtualRows();
}

function renderGeneralPanel() {
    const totalRows = state.dataset.row_count || 0;
    const filteredRows = state.table.derived.filteredRows.length;
    elements.generalPanelContent.innerHTML = `
        <div class="panel-list">
            <label class="control-card checkbox-card">
                <input type="checkbox" data-general="hideMoneyCents" ${state.table.settings.hideMoneyCents ? "checked" : ""}>
                <span>Скрыть копейки для денежных столбцов</span>
            </label>
            <label class="control-card checkbox-card">
                <input type="checkbox" data-general="numbersAsIntegers" ${state.table.settings.numbersAsIntegers ? "checked" : ""}>
                <span>Только целые для числовых столбцов</span>
            </label>
            <label class="control-card control-stack">
                <span>Кол-во строк</span>
                <input type="number" min="1" step="1" data-general="rowLimit" value="${state.table.settings.rowLimit}">
            </label>
            <div class="summary-grid">
                <div class="panel-line"><span>Всего строк</span><strong>${totalRows}</strong></div>
                <div class="panel-line"><span>После фильтрации</span><strong>${filteredRows}</strong></div>
                <div class="panel-line"><span>Режим показа</span><strong>${elements.displayAllToggle.checked ? "Все строки" : "По лимиту"}</strong></div>
            </div>
        </div>
    `;
}

function renderColumnsPanel() {
    const columns = state.table.derived.orderedColumns;
    if (!columns.length) {
        elements.columnsPanelContent.innerHTML = '<p class="muted-copy">После загрузки здесь появится список столбцов.</p>';
        return;
    }

    elements.columnsPanelContent.innerHTML = `
        <div class="panel-list">
            ${columns
                .map((column, index) => `
                    <div class="control-card column-manager-row">
                        <label class="checkbox-card inline-checkbox">
                            <input type="checkbox" data-action="toggle-visible" data-column="${escapeHtml(column.name)}" ${column.visible ? "checked" : ""}>
                            <span>${escapeHtml(column.name)}</span>
                        </label>
                        <div class="mini-actions">
                            <button type="button" class="secondary-button small-button" data-action="move-column" data-column="${escapeHtml(column.name)}" data-direction="up" ${index === 0 ? "disabled" : ""}>↑</button>
                            <button type="button" class="secondary-button small-button" data-action="move-column" data-column="${escapeHtml(column.name)}" data-direction="down" ${index === columns.length - 1 ? "disabled" : ""}>↓</button>
                        </div>
                    </div>
                `)
                .join("")}
        </div>
    `;
}

function renderFilterControl(column) {
    const filter = state.table.filters[column.name] || createFilterState(column.type);

    if (column.type === "text") {
        return `
            <div class="filter-controls single-line">
                <input type="text" placeholder="Содержит..." data-action="filter-text" data-column="${escapeHtml(column.name)}" value="${escapeHtml(filter.value || "")}">
            </div>
        `;
    }

    if (column.type === "number" || column.type === "money") {
        return `
            <div class="filter-controls filter-grid">
                <select data-action="filter-number-operator" data-column="${escapeHtml(column.name)}">
                    <option value="contains" ${filter.operator === "contains" ? "selected" : ""}>Без фильтра</option>
                    <option value="eq" ${filter.operator === "eq" ? "selected" : ""}>Равно</option>
                    <option value="gt" ${filter.operator === "gt" ? "selected" : ""}>Больше</option>
                    <option value="lt" ${filter.operator === "lt" ? "selected" : ""}>Меньше</option>
                    <option value="between" ${filter.operator === "between" ? "selected" : ""}>Диапазон</option>
                </select>
                <input type="text" placeholder="Значение" data-action="filter-number-value" data-column="${escapeHtml(column.name)}" data-field="value" value="${escapeHtml(filter.value || "")}">
                <input type="text" placeholder="От" data-action="filter-number-value" data-column="${escapeHtml(column.name)}" data-field="from" value="${escapeHtml(filter.from || "")}">
                <input type="text" placeholder="До" data-action="filter-number-value" data-column="${escapeHtml(column.name)}" data-field="to" value="${escapeHtml(filter.to || "")}">
            </div>
        `;
    }

    if (column.type === "date" || column.type === "datetime") {
        return `
            <div class="filter-controls filter-grid date-grid">
                <input type="date" data-action="filter-date" data-column="${escapeHtml(column.name)}" data-field="from" value="${escapeHtml(filter.from || "")}">
                <input type="date" data-action="filter-date" data-column="${escapeHtml(column.name)}" data-field="to" value="${escapeHtml(filter.to || "")}">
            </div>
        `;
    }

    if (column.type === "boolean") {
        return `
            <div class="filter-controls single-line">
                <select data-action="filter-boolean" data-column="${escapeHtml(column.name)}">
                    <option value="any" ${filter.value === "any" ? "selected" : ""}>Все</option>
                    <option value="yes" ${filter.value === "yes" ? "selected" : ""}>Да</option>
                    <option value="no" ${filter.value === "no" ? "selected" : ""}>Нет</option>
                    <option value="empty" ${filter.value === "empty" ? "selected" : ""}>Пусто</option>
                </select>
            </div>
        `;
    }

    return "";
}

function renderFiltersPanel() {
    const columns = state.table.derived.orderedColumns;
    if (!columns.length) {
        elements.filtersPanelContent.innerHTML = '<p class="muted-copy">После загрузки здесь появятся фильтры по столбцам.</p>';
        return;
    }

    elements.filtersPanelContent.innerHTML = `
        <div class="panel-list">
            ${columns
                .map((column) => `
                    <div class="control-card control-stack">
                        <div class="filter-head">
                            <strong>${escapeHtml(column.name)}</strong>
                            <span>${escapeHtml(column.type)}</span>
                        </div>
                        ${renderFilterControl(column)}
                    </div>
                `)
                .join("")}
        </div>
    `;
}

function renderTypesPanel() {
    const columns = state.table.derived.orderedColumns;
    if (!columns.length) {
        elements.typesPanelContent.innerHTML = '<p class="muted-copy">После загрузки здесь появятся типы столбцов.</p>';
        return;
    }

    elements.typesPanelContent.innerHTML = `
        <div class="panel-list">
            ${columns
                .map(
                    (column) => `
                        <div class="control-card type-row">
                            <div>
                                <strong>${escapeHtml(column.name)}</strong>
                                <span>${column.fromTemplate ? "Тип был задан шаблоном" : "Тип можно сохранить в шаблон"}</span>
                            </div>
                            <select data-action="change-type" data-column="${escapeHtml(column.name)}">
                                ${columnTypes
                                    .map(
                                        (type) => `<option value="${type}" ${type === column.type ? "selected" : ""}>${type}</option>`
                                    )
                                    .join("")}
                            </select>
                        </div>
                    `
                )
                .join("")}
        </div>
    `;
}

function renderProcessingPanel() {
    elements.processingPanelContent.innerHTML = `
        <div class="panel-list">
            <div class="panel-note">Обработка данных начинается в следующем спринте.</div>
            <div class="panel-line"><span>Активный шаблон</span><strong>${escapeHtml(getActiveTemplate()?.name || "Без шаблона")}</strong></div>
            <div class="panel-line"><span>Ошибок после пересчёта</span><strong>${state.table.derived.errors.length}</strong></div>
        </div>
    `;
}

function renderRightPanels() {
    renderGeneralPanel();
    renderColumnsPanel();
    renderFiltersPanel();
    renderTypesPanel();
    renderProcessingPanel();
    updateTemplateColumnCount();
}

function renderErrorsPanel() {
    const warnings = state.dataset.warnings || [];
    const errors = state.table.derived.errors || [];

    if (!warnings.length && !errors.length) {
        elements.issuesSummary.textContent = "Ошибок и предупреждений пока нет.";
        elements.issuesList.innerHTML = "";
        elements.errorsPanelMessage.textContent = "После загрузки CSV здесь появятся предупреждения и ошибки валидации.";
        return;
    }

    elements.issuesSummary.textContent = `Предупреждений: ${warnings.length}. Ошибок: ${errors.length}.`;
    const warningMarkup = warnings.map(
        (warning) => `<div class="issue-item warning"><strong>Предупреждение</strong><span>${escapeHtml(warning)}</span></div>`
    );
    const errorMarkup = errors.map(
        (error) => `
            <div class="issue-item error">
                <strong>Строка ${error.row_number}, столбец ${escapeHtml(error.column_name)}</strong>
                <span>${escapeHtml(error.reason)}. Значение: ${escapeHtml(error.value || "")}</span>
            </div>
        `
    );
    elements.issuesList.innerHTML = [...warningMarkup, ...errorMarkup].join("");
    elements.errorsPanelMessage.textContent = state.dataset.partial
        ? "Загрузка выполнена частично. Проверьте отсутствующие столбцы и ошибки после пересчёта типов."
        : "Ниже показаны предупреждения и ошибки текущего рабочего набора.";
}

function refreshDataViews() {
    recomputeTableState();
    renderDataset();
    renderRightPanels();
    renderErrorsPanel();
}

function updateTemplateInState(updatedTemplate) {
    state.templates = state.templates.map((template) => (template.id === updatedTemplate.id ? updatedTemplate : template));
    renderTemplateList();
    renderUploadTemplateOptions();
    if (String(elements.templateId.value || "") === String(updatedTemplate.id)) {
        fillTemplateForm(updatedTemplate);
    }
}

function buildTypePersistencePayload(template) {
    const typeByName = new Map(state.table.columns.map((column) => [column.name.toLowerCase(), column.type]));
    const existingNames = new Set();
    const columns = [];

    template.columns.forEach((column) => {
        const key = column.name.toLowerCase();
        existingNames.add(key);
        columns.push({
            name: column.name,
            type: typeByName.get(key) || column.type,
        });
    });

    state.table.derived.orderedColumns.forEach((column) => {
        const key = column.name.toLowerCase();
        if (!existingNames.has(key)) {
            columns.push({ name: column.name, type: column.type });
        }
    });

    return {
        name: template.name,
        columns,
    };
}

async function persistCurrentTypes() {
    const activeTemplate = getActiveTemplate();
    if (!activeTemplate) {
        return;
    }

    const payload = buildTypePersistencePayload(activeTemplate);
    try {
        const response = await requestJson(`/api/templates/${activeTemplate.id}`, {
            method: "PUT",
            body: JSON.stringify(payload),
        });
        updateTemplateInState(response.template);
        await loadJournal();
        setStatus("Тип столбца сохранён в шаблон", "success");
    } catch {
        setStatus("Не удалось сохранить тип в шаблон", "error");
    }
}

async function reloadBootstrap(preferredTemplateId = null) {
    const payload = await requestJson("/api/bootstrap");
    state.settings = payload.settings;
    state.templates = payload.templates;
    state.ui = payload.state;
    state.dataset = payload.dataset || emptyDataset();
    state.activeTemplateId = preferredTemplateId ?? state.ui.last_template_id ?? state.templates[0]?.id ?? null;

    syncSettingsForm();
    renderTemplateList();
    fillTemplateForm(getActiveTemplate());
    renderUploadTemplateOptions();
    applyUiState();
    initializeTableModel();
    renderDataset();
    renderRightPanels();
    renderErrorsPanel();
    await loadJournal();
}

async function saveSettings(event) {
    event.preventDefault();
    elements.settingsErrors.textContent = "";

    try {
        const payload = {
            source_folder: elements.settingsForm.source_folder.value.trim(),
            output_folder: elements.settingsForm.output_folder.value.trim(),
            instructions_file: elements.settingsForm.instructions_file.value.trim(),
        };
        const response = await requestJson("/api/settings", {
            method: "PUT",
            body: JSON.stringify(payload),
        });
        state.settings = response.settings;
        setStatus("Настройки сохранены", "success");
        await loadJournal();
    } catch (error) {
        const messages = Object.values(error.errors || {}).join(" ");
        elements.settingsErrors.textContent = messages || "Не удалось сохранить настройки.";
        setStatus("Ошибка сохранения настроек", "error");
    }
}

function scheduleStateSave(patch = {}, options = {}) {
    const notify = options.notify !== false;
    state.ui = { ...state.ui, ...patch };
    applyUiState();

    window.clearTimeout(stateSaveTimer);
    stateSaveTimer = window.setTimeout(async () => {
        try {
            await requestJson("/api/state", {
                method: "PUT",
                body: JSON.stringify(state.ui),
            });
            if (notify) {
                setStatus("Состояние интерфейса сохранено", "success");
            }
        } catch {
            if (notify) {
                setStatus("Ошибка сохранения состояния", "error");
            }
        }
    }, 250);
}

async function saveTemplate(event) {
    event.preventDefault();
    elements.templateErrors.textContent = "";
    const templateId = elements.templateId.value;
    const payload = getTemplateFormPayload();

    try {
        const url = templateId ? `/api/templates/${templateId}` : "/api/templates";
        const method = templateId ? "PUT" : "POST";
        const response = await requestJson(url, {
            method,
            body: JSON.stringify(payload),
        });
        const savedTemplate = response.template;
        await reloadBootstrap(savedTemplate.id);
        setStatus(templateId ? "Шаблон обновлён" : "Шаблон создан", "success");
    } catch (error) {
        elements.templateErrors.textContent = (error.errors || ["Не удалось сохранить шаблон."]).join(" ");
        setStatus("Ошибка сохранения шаблона", "error");
    }
}

async function deleteTemplate() {
    const templateId = elements.templateId.value;
    if (!templateId) {
        return;
    }

    const templateName = elements.templateName.value.trim() || "без названия";
    const confirmed = window.confirm(`Удалить шаблон "${templateName}"?`);
    if (!confirmed) {
        return;
    }

    try {
        await requestJson(`/api/templates/${templateId}`, { method: "DELETE" });
        await reloadBootstrap();
        setStatus("Шаблон удалён", "success");
    } catch {
        elements.templateErrors.textContent = "Не удалось удалить шаблон.";
        setStatus("Ошибка удаления шаблона", "error");
    }
}

async function handleFileSelected(event) {
    const file = event.target.files?.[0];
    if (!file) {
        setStatus("Файл не выбран.", "warning");
        return;
    }

    setStatus(`Выбран файл: ${file.name}`, "info");
    const formData = new FormData();
    formData.append("file", file);
    if (state.activeTemplateId) {
        formData.append("template_id", String(state.activeTemplateId));
    }

    try {
        const response = await requestJson("/api/data/upload", {
            method: "POST",
            body: formData,
        });
        state.dataset = response.dataset || emptyDataset();
        initializeTableModel();
        renderDataset();
        renderRightPanels();
        renderErrorsPanel();
        scheduleStateSave({ last_loaded_file: state.dataset.file_name || "" }, { notify: false });
        await loadJournal();

        const level = state.dataset.partial || state.dataset.warnings.length ? "warning" : "success";
        const message = state.dataset.partial
            ? `Файл ${state.dataset.file_name} загружен частично.`
            : `Данные успешно загружены: ${state.dataset.file_name}.`;
        setStatus(message, level);
        switchTab("analysis");
    } catch (error) {
        state.dataset = emptyDataset();
        resetTableModel(true);
        renderDataset();
        renderRightPanels();
        renderErrorsPanel();
        setStatus(error.error || "Ошибка загрузки файла.", "error");
    } finally {
        elements.csvFileInput.value = "";
    }
}

function toggleSort(columnName) {
    if (state.table.sort.columnName !== columnName) {
        state.table.sort = { columnName, direction: "asc" };
        return;
    }
    if (state.table.sort.direction === "asc") {
        state.table.sort.direction = "desc";
        return;
    }
    if (state.table.sort.direction === "desc") {
        state.table.sort = { columnName: null, direction: null };
        return;
    }
    state.table.sort = { columnName, direction: "asc" };
}

function swapColumns(columnName, direction) {
    const ordered = getOrderedColumns();
    const index = ordered.findIndex((column) => column.name === columnName);
    if (index < 0) {
        return;
    }
    const targetIndex = direction === "up" ? index - 1 : index + 1;
    if (targetIndex < 0 || targetIndex >= ordered.length) {
        return;
    }
    const left = ordered[index];
    const right = ordered[targetIndex];
    const currentPosition = left.position;
    left.position = right.position;
    right.position = currentPosition;
}

function ensureFilter(columnName, type) {
    if (!state.table.filters[columnName]) {
        state.table.filters[columnName] = createFilterState(type);
    }
    return state.table.filters[columnName];
}

function handleDetailsClick(event) {
    const button = event.target.closest("[data-action]");
    if (!button) {
        return;
    }

    const action = button.dataset.action;
    const columnName = button.dataset.column;

    if (action === "move-column") {
        swapColumns(columnName, button.dataset.direction);
        refreshDataViews();
        setStatus("Порядок столбцов обновлён", "success");
        return;
    }

    if (action === "sort-column") {
        toggleSort(columnName);
        refreshDataViews();
        setStatus("Сортировка обновлена", "success");
    }
}

function handleHeaderClick(event) {
    const button = event.target.closest("[data-action='sort-column']");
    if (!button) {
        return;
    }
    toggleSort(button.dataset.column);
    refreshDataViews();
    setStatus("Сортировка обновлена", "success");
}

async function handleDetailsChange(event) {
    const target = event.target;

    if (target.dataset.general === "hideMoneyCents") {
        state.table.settings.hideMoneyCents = target.checked;
        refreshDataViews();
        setStatus("Настройка отображения обновлена", "success");
        return;
    }

    if (target.dataset.general === "numbersAsIntegers") {
        state.table.settings.numbersAsIntegers = target.checked;
        refreshDataViews();
        setStatus("Настройка отображения обновлена", "success");
        return;
    }

    if (target.dataset.general === "rowLimit") {
        state.table.settings.rowLimit = Math.max(1, Number(target.value) || 1);
        refreshDataViews();
        setStatus("Лимит отображения обновлён", "success");
        return;
    }

    if (target.dataset.action === "toggle-visible") {
        const column = getColumnState(target.dataset.column);
        if (!column) {
            return;
        }
        column.visible = target.checked;
        refreshDataViews();
        setStatus("Видимость столбца обновлена", "success");
        return;
    }

    if (target.dataset.action === "change-type") {
        const column = getColumnState(target.dataset.column);
        if (!column) {
            return;
        }
        column.type = target.value;
        state.table.filters[column.name] = createFilterState(column.type);
        refreshDataViews();
        await persistCurrentTypes();
        return;
    }

    if (target.dataset.action === "filter-number-operator") {
        const column = getColumnState(target.dataset.column);
        const filter = ensureFilter(target.dataset.column, column?.type || "number");
        filter.operator = target.value;
        refreshDataViews();
        return;
    }

    if (target.dataset.action === "filter-boolean") {
        const column = getColumnState(target.dataset.column);
        const filter = ensureFilter(target.dataset.column, column?.type || "boolean");
        filter.value = target.value;
        refreshDataViews();
        return;
    }

    if (target.dataset.action === "filter-date") {
        const column = getColumnState(target.dataset.column);
        const filter = ensureFilter(target.dataset.column, column?.type || "date");
        filter[target.dataset.field] = target.value;
        refreshDataViews();
    }
}

function handleDetailsInput(event) {
    const target = event.target;
    if (target.dataset.action === "filter-text") {
        const column = getColumnState(target.dataset.column);
        const filter = ensureFilter(target.dataset.column, column?.type || "text");
        filter.value = target.value;
        refreshDataViews();
        return;
    }

    if (target.dataset.action === "filter-number-value") {
        const column = getColumnState(target.dataset.column);
        const filter = ensureFilter(target.dataset.column, column?.type || "number");
        filter[target.dataset.field] = target.value;
        refreshDataViews();
    }
}

function bindEvents() {
    elements.tabButtons.forEach((button) => {
        button.addEventListener("click", () => switchTab(button.dataset.tab));
    });

    elements.rightTabButtons.forEach((button) => {
        button.addEventListener("click", () => switchRightTab(button.dataset.rightTab));
    });

    elements.detailsPanel.addEventListener("click", handleDetailsClick);
    elements.dataTableHeader.addEventListener("click", handleHeaderClick);
    elements.detailsPanel.addEventListener("change", (event) => {
        void handleDetailsChange(event);
    });
    elements.detailsPanel.addEventListener("input", handleDetailsInput);

    elements.settingsForm.addEventListener("submit", saveSettings);
    elements.templateForm.addEventListener("submit", saveTemplate);
    elements.newTemplateButton.addEventListener("click", () => {
        state.activeTemplateId = null;
        fillTemplateForm(null);
        renderTemplateList();
        renderUploadTemplateOptions();
        renderRightPanels();
        syncStateLabels();
        setStatus("Новый шаблон", "info");
    });
    elements.addColumnButton.addEventListener("click", () => {
        elements.columnsContainer.appendChild(createColumnRow());
        updateTemplateColumnCount();
    });
    elements.deleteTemplateButton.addEventListener("click", deleteTemplate);

    elements.leftPanelWidth.addEventListener("input", (event) => {
        scheduleStateSave({ left_panel_width: Number(event.target.value) });
    });
    elements.rightPanelVisible.addEventListener("change", (event) => {
        scheduleStateSave({ right_panel_visible: event.target.checked });
        setStatus(event.target.checked ? "Панель данных показана." : "Панель данных скрыта.", "info");
    });
    elements.errorsVisible.addEventListener("change", (event) => {
        scheduleStateSave({ errors_visible: event.target.checked });
    });
    elements.errorsHeight.addEventListener("input", (event) => {
        scheduleStateSave({ errors_height: Number(event.target.value) });
    });
    elements.lastLoadedFile.addEventListener("change", (event) => {
        scheduleStateSave({ last_loaded_file: event.target.value.trim() });
    });

    elements.uploadTemplateSelect.addEventListener("change", (event) => {
        const nextId = event.target.value ? Number(event.target.value) : null;
        setActiveTemplate(nextId);
    });
    elements.uploadButton.addEventListener("click", () => elements.csvFileInput.click());
    elements.csvFileInput.addEventListener("change", handleFileSelected);
    elements.exportButton.addEventListener("click", () => {
        if (!state.dataset.rows.length) {
            setStatus("Нет данных для сохранения.", "warning");
            return;
        }
        setStatus("Сохранение в XLSX запланировано на следующий спринт.", "info");
    });
    elements.toggleRightPanelButton.addEventListener("click", () => {
        scheduleStateSave({ right_panel_visible: !state.ui.right_panel_visible });
        setStatus(state.ui.right_panel_visible ? "Панель данных показана." : "Панель данных скрыта.", "info");
    });
    elements.displayAllToggle.addEventListener("change", () => {
        refreshDataViews();
        setStatus(
            elements.displayAllToggle.checked ? "Включено отображение всех строк." : "Включено ограничение по лимиту строк.",
            "info"
        );
    });
    elements.dataTableViewport.addEventListener("scroll", renderVirtualRows);
}

async function init() {
    bindEvents();
    await reloadBootstrap();
    switchRightTab(state.activeRightTab);
    setStatus("Рабочая область готова.", "success");
}

init().catch(() => {
    setStatus("Ошибка инициализации приложения", "error");
    elements.errorsPanelMessage.textContent = "Не удалось загрузить начальные данные приложения.";
});