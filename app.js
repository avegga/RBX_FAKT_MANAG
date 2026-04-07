const state = {
    settings: {},
    templates: [],
    ui: {},
    activeTemplateId: null,
};

const columnTypes = ["text", "number", "money", "date", "datetime", "boolean"];
let stateSaveTimer = null;

const elements = {
    body: document.body,
    tabButtons: Array.from(document.querySelectorAll(".tab-button")),
    tabPanels: Array.from(document.querySelectorAll(".tab-panel")),
    activeTabTitle: document.getElementById("active-tab-title"),
    saveStatus: document.getElementById("save-status"),
    settingsForm: document.getElementById("settings-form"),
    settingsErrors: document.getElementById("settings-errors"),
    stateForm: document.getElementById("state-form"),
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
    errorsPanel: document.getElementById("errors-panel"),
    errorsPanelMessage: document.getElementById("errors-panel-message"),
};

async function requestJson(url, options = {}) {
    const response = await fetch(url, {
        headers: { "Content-Type": "application/json" },
        ...options,
    });

    const data = await response.json().catch(() => ({}));
    if (!response.ok) {
        throw data;
    }

    return data;
}

function setStatus(message, isError = false) {
    elements.saveStatus.textContent = message;
    elements.saveStatus.style.background = isError ? "#fce5e0" : "#d7efeb";
    elements.saveStatus.style.color = isError ? "#b42318" : "#0f5d57";
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

function renderTemplateList() {
    elements.templateList.innerHTML = "";
    elements.templateCount.textContent = String(state.templates.length);

    if (state.templates.length === 0) {
        const item = document.createElement("li");
        item.textContent = "Шаблоны пока не созданы.";
        elements.templateList.appendChild(item);
        elements.columnCount.textContent = "0";
        return;
    }

    state.templates.forEach((template) => {
        const item = document.createElement("li");
        const button = document.createElement("button");
        button.type = "button";
        button.textContent = template.name;
        button.classList.toggle("active", template.id === state.activeTemplateId);
        button.addEventListener("click", () => selectTemplate(template.id));
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
        updateColumnCount();
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
    updateColumnCount();
}

function updateColumnCount() {
    const count = elements.columnsContainer.querySelectorAll(".column-row").length;
    elements.columnCount.textContent = String(count);
}

function selectTemplate(templateId) {
    const template = state.templates.find((item) => item.id === templateId) || null;
    state.activeTemplateId = template ? template.id : null;
    fillTemplateForm(template);
    renderTemplateList();
    syncStateLabels();
    scheduleStateSave({ last_template_id: state.activeTemplateId });
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

    syncStateLabels();
}

function syncStateLabels() {
    const activeTemplate = state.templates.find((item) => item.id === state.activeTemplateId);
    elements.lastTemplateLabel.textContent = activeTemplate ? activeTemplate.name : "Не выбран";
    elements.lastFileLabel.textContent = state.ui.last_loaded_file || "Не задан";
    elements.rightPanelLabel.textContent = state.ui.right_panel_visible ? "Включена" : "Скрыта";
    elements.errorsPanelLabel.textContent = state.ui.errors_visible ? "Включена" : "Скрыта";
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

async function reloadBootstrap(preferredTemplateId = null) {
    const payload = await requestJson("/api/bootstrap");
    state.settings = payload.settings;
    state.templates = payload.templates;
    state.ui = payload.state;
    state.activeTemplateId = preferredTemplateId ?? state.ui.last_template_id ?? state.templates[0]?.id ?? null;

    syncSettingsForm();
    renderTemplateList();
    selectTemplate(state.activeTemplateId);
    applyUiState();
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
        setStatus("Настройки сохранены");
        await loadJournal();
    } catch (error) {
        const messages = Object.values(error.errors || {}).join(" ");
        elements.settingsErrors.textContent = messages || "Не удалось сохранить настройки.";
        setStatus("Ошибка сохранения настроек", true);
    }
}

function scheduleStateSave(patch = {}) {
    state.ui = { ...state.ui, ...patch };
    applyUiState();

    window.clearTimeout(stateSaveTimer);
    stateSaveTimer = window.setTimeout(async () => {
        try {
            await requestJson("/api/state", {
                method: "PUT",
                body: JSON.stringify(state.ui),
            });
            setStatus("Состояние интерфейса сохранено");
        } catch {
            setStatus("Ошибка сохранения состояния", true);
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
        setStatus(templateId ? "Шаблон обновлён" : "Шаблон создан");
    } catch (error) {
        elements.templateErrors.textContent = (error.errors || ["Не удалось сохранить шаблон."]).join(" ");
        setStatus("Ошибка сохранения шаблона", true);
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
        fillTemplateForm(null);
        await reloadBootstrap();
        setStatus("Шаблон удалён");
    } catch {
        elements.templateErrors.textContent = "Не удалось удалить шаблон.";
        setStatus("Ошибка удаления шаблона", true);
    }
}

function bindEvents() {
    elements.tabButtons.forEach((button) => {
        button.addEventListener("click", () => switchTab(button.dataset.tab));
    });

    elements.settingsForm.addEventListener("submit", saveSettings);
    elements.templateForm.addEventListener("submit", saveTemplate);
    elements.newTemplateButton.addEventListener("click", () => {
        state.activeTemplateId = null;
        fillTemplateForm(null);
        renderTemplateList();
        setStatus("Новый шаблон");
    });
    elements.addColumnButton.addEventListener("click", () => {
        elements.columnsContainer.appendChild(createColumnRow());
        updateColumnCount();
    });
    elements.deleteTemplateButton.addEventListener("click", deleteTemplate);

    elements.leftPanelWidth.addEventListener("input", (event) => {
        scheduleStateSave({ left_panel_width: Number(event.target.value) });
    });
    elements.rightPanelVisible.addEventListener("change", (event) => {
        scheduleStateSave({ right_panel_visible: event.target.checked });
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
}

async function init() {
    bindEvents();
    await reloadBootstrap();
    setStatus("Данные загружены");
}

init().catch(() => {
    setStatus("Ошибка инициализации приложения", true);
    elements.errorsPanelMessage.textContent = "Не удалось загрузить начальные данные приложения.";
});
