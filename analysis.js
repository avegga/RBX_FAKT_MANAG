(function () {
    const MAX_ANALYSIS_ROWS = 250;
    const defaultUserState = {
        selected_analysis_type_id: null,
        left_panel_width: 260,
        visual_source_kind: "none",
        view_mode: "editor",
        table_search: "",
        table_sort_column: "",
        table_sort_direction: "asc",
        table_column_widths: {},
        draft_state: null,
    };
    const emptyDataset = () => ({
        source_kind: "none",
        source_mode: "none",
        source_label: "",
        source_file_name: "",
        attached_at: null,
        source_status: "empty",
        columns: [],
        rows: [],
        row_count: 0,
        column_count: 0,
        warnings: [],
    });
    const defaultChart = (position, sourceKind = "none") => ({
        chart_type: "bar",
        source_kind: sourceKind,
        x_field: "",
        y_field: "",
        group_field: "",
        agg_func: "count",
        color: "#b7791f",
        legend: "",
        legend_position: "right",
        pie_label_mode: "legend",
        pie_value_mode: "value",
        pie_center_mode: "total",
        pie_sector_label_mode: "name-percent",
        labels: "",
        comment_title: "",
        comment_text: "",
        annotations: [],
        pie_sector_label_offsets: [],
        is_hidden: false,
        ui_collapsed: false,
        position,
    });

    const state = {
        types: [],
        userState: { ...defaultUserState },
        dataset: emptyDataset(),
        chartPreviews: [],
        viewMode: "editor",
        loading: false,
        message: {
            text: "Выберите тип анализа или подключите данные из вкладки «Загрузка фактов».",
            level: "info",
        },
        draft: {
            selectedTypeId: null,
            name: "",
            charts: [],
        },
        table: {
            search: "",
            sortColumn: "",
            sortDirection: "asc",
            columnWidths: {},
        },
    };

    let widthSaveTimer = null;
    let chartPreviewTimer = null;
    let chartPreviewRequestId = 0;
    let userStateSaveTimer = null;
    let pendingUserStatePatch = null;
    let activeUserStateSavePromise = null;
    let activeTableResize = null;
    let activeAnnotationDrag = null;
    let activePieSectorLabelDrag = null;
    let sheetLayoutFrame = null;

    const elements = {
        shell: document.getElementById("analysis-shell"),
        message: document.getElementById("alert_message"),
        sourceMeta: document.getElementById("analysis-source-meta"),
        uploadFileButton: document.getElementById("btn_upload_file"),
        analysisFileInput: document.getElementById("analysis-file-input"),
        useFactsButton: document.getElementById("btn_use_facts"),
        typeSelect: document.getElementById("select_analysis_type"),
        editorViewButton: document.getElementById("btn_analysis_editor_view"),
        sheetViewButton: document.getElementById("btn_analysis_sheet_view"),
        pairViewButton: document.getElementById("btn_analysis_pair_view"),
        quadViewButton: document.getElementById("btn_analysis_quad_view"),
        createTypeButton: document.getElementById("btn_create_analysis_type"),
        deleteTypeButton: document.getElementById("btn_delete_analysis_type"),
        saveTypeButton: document.getElementById("btn_save_analysis_type"),
        resetButton: document.getElementById("btn_reset_settings"),
        exportReportButton: document.getElementById("btn_export_analysis_report"),
        chartsList: document.getElementById("list_charts"),
        compatibilityNote: document.getElementById("analysis-compatibility-note"),
        addChartButton: document.getElementById("btn_add_chart"),
        searchInput: document.getElementById("input_search"),
        sortResetButton: document.getElementById("panel_sorting"),
        countLabel: document.getElementById("label_count"),
        loadingIndicator: document.getElementById("loading_indicator"),
        emptyState: document.getElementById("empty_state"),
        retryUseFactsButton: document.getElementById("btn_retry_use_facts"),
        mainView: document.getElementById("analysis-main"),
        sheetView: document.getElementById("analysis-sheet-view"),
        sheetTitle: document.getElementById("analysis-sheet-title"),
        sheetDescription: document.getElementById("analysis-sheet-description"),
        sheetGrid: document.getElementById("analysis-sheet-grid"),
        tableWrap: document.getElementById("analysis-table-wrap"),
        tableElement: document.getElementById("data_table"),
        tableColgroup: document.getElementById("analysis-table-colgroup"),
        tableHead: document.getElementById("analysis-table-head"),
        tableBody: document.getElementById("analysis-table-body"),
    };

    if (!elements.shell) {
        return;
    }

    function apiJson(url, options = {}) {
        const isFormData = options.body instanceof FormData;
        const headers = isFormData ? {} : { "Content-Type": "application/json" };
        return fetch(url, {
            headers: { ...headers, ...(options.headers || {}) },
            ...options,
        }).then(async (response) => {
            const data = await response.json().catch(() => ({}));
            if (!response.ok) {
                throw data;
            }
            return data;
        });
    }

    function escapeHtml(value) {
        return String(value ?? "")
            .replaceAll("&", "&amp;")
            .replaceAll("<", "&lt;")
            .replaceAll(">", "&gt;")
            .replaceAll('"', "&quot;");
    }

    function clamp(value, min, max) {
        return Math.min(max, Math.max(min, value));
    }

    function createAnnotationId() {
        return `annotation-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
    }

    function setMessage(text, level = "info") {
        state.message = { text, level };
        renderMessage();
    }

    function loadAnalysisUiPreferences() {
        try {
            const viewMode = localStorage.getItem("analysis-view-mode");
            state.viewMode = viewMode === "sheet" || viewMode === "pair" || viewMode === "quad" ? viewMode : "editor";
            state.userState.view_mode = state.viewMode;
        } catch {
            state.viewMode = "editor";
        }

        try {
            const rawWidths = localStorage.getItem("analysis-table-column-widths");
            const parsed = rawWidths ? JSON.parse(rawWidths) : {};
            state.table.columnWidths = parsed && typeof parsed === "object" ? parsed : {};
            state.userState.table_column_widths = state.table.columnWidths;
        } catch {
            state.table.columnWidths = {};
        }
    }

    function persistAnalysisUiPreferences() {
        try {
            localStorage.setItem("analysis-view-mode", state.viewMode);
            localStorage.setItem("analysis-table-column-widths", JSON.stringify(state.table.columnWidths || {}));
        } catch {}
    }

    function applyPersistedAnalysisUiState() {
        const persistedViewMode = state.userState.view_mode;
        if (persistedViewMode === "sheet" || persistedViewMode === "pair" || persistedViewMode === "quad" || persistedViewMode === "editor") {
            state.viewMode = persistedViewMode;
        }

        const persistedWidths = state.userState.table_column_widths;
        if (persistedWidths && typeof persistedWidths === "object" && !Array.isArray(persistedWidths) && Object.keys(persistedWidths).length) {
            state.table.columnWidths = { ...persistedWidths };
        }

        persistAnalysisUiPreferences();
    }

    function renderMessage() {
        elements.message.textContent = state.message.text;
        elements.message.className = `service-message ${state.message.level} analysis-message`;
    }

    function formatTimestamp(value) {
        if (!value) {
            return "Не подключен";
        }
        const parsed = new Date(value);
        if (Number.isNaN(parsed.getTime())) {
            return String(value);
        }
        return new Intl.DateTimeFormat("ru-RU", {
            dateStyle: "short",
            timeStyle: "short",
        }).format(parsed);
    }

    function getSourceKindLabel(kind) {
        if (kind === "facts") {
            return "Загрузка фактов";
        }
        if (kind === "file") {
            return "Файл анализа";
        }
        return "Не выбран";
    }

    function getSourceModeLabel(mode) {
        if (mode === "export_layer") {
            return "Экспортируемый слой";
        }
        if (mode === "uploaded_file") {
            return "Загруженный файл";
        }
        return "Нет данных";
    }

    function renderSourceMeta() {
        const dataset = state.dataset;
        const fallbackKind = dataset.source_kind === "none" ? state.userState.visual_source_kind : dataset.source_kind;
        elements.sourceMeta.innerHTML = `
            <div class="analysis-source-card"><span>Источник</span><strong>${escapeHtml(getSourceKindLabel(fallbackKind))}</strong></div>
            <div class="analysis-source-card"><span>Режим</span><strong>${escapeHtml(getSourceModeLabel(dataset.source_mode))}</strong></div>
            <div class="analysis-source-card"><span>Набор</span><strong>${escapeHtml(dataset.source_label || "Не подключен")}</strong></div>
            <div class="analysis-source-card"><span>Подключен</span><strong>${escapeHtml(formatTimestamp(dataset.attached_at))}</strong></div>
        `;
    }

    function getAvailableColumns() {
        return new Set((state.dataset.columns || []).map((column) => column.name));
    }

    function getChartCompatibility(chart) {
        const issues = [];
        const availableColumns = getAvailableColumns();
        const activeSourceKind = state.dataset.source_kind;

        if (chart.source_kind && chart.source_kind !== "none" && activeSourceKind !== "none" && chart.source_kind !== activeSourceKind) {
            issues.push(`ожидается источник ${getSourceKindLabel(chart.source_kind).toLowerCase()}`);
        }

        [chart.x_field, chart.y_field, chart.group_field].filter(Boolean).forEach((fieldName) => {
            if (!availableColumns.has(fieldName)) {
                issues.push(`поле '${fieldName}' недоступно в текущем источнике`);
            }
        });

        return {
            ok: issues.length === 0,
            issues,
        };
    }

    function formatPreviewNumber(value) {
        const number = Number(value);
        if (!Number.isFinite(number)) {
            return String(value ?? "");
        }
        return new Intl.NumberFormat("ru-RU", { maximumFractionDigits: 2 }).format(number);
    }

    function formatCompactPreviewNumber(value) {
        const number = Number(value);
        if (!Number.isFinite(number)) {
            return String(value ?? "");
        }

        const absValue = Math.abs(number);
        if (absValue >= 1000000) {
            return `${new Intl.NumberFormat("ru-RU", { maximumFractionDigits: 1 }).format(number / 1000000)} млн`;
        }
        if (absValue >= 1000) {
            return `${new Intl.NumberFormat("ru-RU", { maximumFractionDigits: 1 }).format(number / 1000)} тыс`;
        }
        return new Intl.NumberFormat("ru-RU", { maximumFractionDigits: 1 }).format(number);
    }

    function shortenAxisLabel(value, maxLength = 12) {
        const text = String(value ?? "");
        if (maxLength <= 1) {
            return text.slice(0, 1);
        }
        return text.length > maxLength ? `${text.slice(0, maxLength - 1)}…` : text;
    }

    function getChartRenderOptions(options = {}, chart = null) {
        const compact = Boolean(options.compact);
        const hideWarnings = Boolean(options.hideWarnings);
        const hideHead = Boolean(options.hideHead || compact);
        const hideMeta = Boolean(options.hideMeta || compact);
        const layoutMode = options.layoutMode || "editor";
        let legendPosition = options.legendPosition || chart?.legend_position || "right";

        if (compact) {
            if (chart?.chart_type === "pie") {
                legendPosition = "bottom";
            } else if (legendPosition === "right" || legendPosition === "left") {
                legendPosition = "bottom";
            }
        }

        return {
            compact,
            hideWarnings,
            hideHead,
            hideMeta,
            layoutMode,
            legendPosition,
        };
    }

    function buildCompactAxisLabel(value, maxLength = 7) {
        return shortenAxisLabel(value, maxLength);
    }

    function normalizePieSectorLabelOffsets(rawOffsets) {
        if (!Array.isArray(rawOffsets)) {
            return [];
        }

        return rawOffsets
            .map((item) => {
                const index = Number(item?.index);
                const dx = Number(item?.dx);
                const dy = Number(item?.dy);
                if (!Number.isFinite(index) || index < 0) {
                    return null;
                }
                return {
                    index,
                    dx: clamp(Number.isFinite(dx) ? dx : 0, -100, 100),
                    dy: clamp(Number.isFinite(dy) ? dy : 0, -100, 100),
                };
            })
            .filter(Boolean)
            .slice(0, 24);
    }

    function formatPiePercent(value, total) {
        const numericValue = Number(value || 0);
        const numericTotal = Number(total || 0);
        if (!numericTotal) {
            return "0%";
        }

        const percent = (numericValue / numericTotal) * 100;
        const precision = percent >= 10 || Number.isInteger(percent) ? 0 : 1;
        return `${new Intl.NumberFormat("ru-RU", { maximumFractionDigits: precision }).format(percent)}%`;
    }

    function buildPieValueText(value, total, mode, compact = false) {
        const numericMode = mode || "value";
        const valueText = compact ? formatCompactPreviewNumber(value) : formatPreviewNumber(value);
        const percentText = formatPiePercent(value, total);

        if (numericMode === "percent") {
            return percentText;
        }
        if (numericMode === "value-percent") {
            return `${valueText} (${percentText})`;
        }
        return valueText;
    }

    function buildPieCenterContent(total, chart, compact = false) {
        const centerMode = chart?.pie_center_mode || "total";
        if (centerMode === "hidden") {
            return "";
        }

        return compact ? formatCompactPreviewNumber(total) : formatPreviewNumber(total);
    }

    function buildPieSectorLabelText(category, value, total, mode, compact = false) {
        const sectorMode = mode || "name-percent";
        const categoryText = compact ? shortenAxisLabel(categoryLabel(category), 10) : categoryLabel(category);
        const valueText = formatCompactPreviewNumber(value);
        const percentText = buildPieValueText(value, total, "percent", compact);

        if (sectorMode === "name") {
            return categoryText;
        }
        if (sectorMode === "value") {
            return valueText;
        }
        if (sectorMode === "percent") {
            return percentText;
        }
        if (sectorMode === "value-percent") {
            return `${valueText} (${percentText})`;
        }
        if (sectorMode === "name-value") {
            return `${categoryText} ${valueText}`;
        }
        if (sectorMode === "name-value-percent") {
            return `${categoryText} ${valueText} (${percentText})`;
        }
        return `${categoryText} ${percentText}`;
    }

    function distributePieLabels(items, minGap, minY, maxY) {
        const ordered = items.slice().sort((left, right) => left.y - right.y);
        let previousY = minY - minGap;

        ordered.forEach((item) => {
            item.y = Math.max(item.y, previousY + minGap);
            previousY = item.y;
        });

        if (ordered.length) {
            let nextY = maxY + minGap;
            for (let index = ordered.length - 1; index >= 0; index -= 1) {
                ordered[index].y = Math.min(ordered[index].y, nextY - minGap);
                nextY = ordered[index].y;
            }
        }

        ordered.forEach((item) => {
            item.y = clamp(item.y, minY, maxY);
        });

        return ordered;
    }

    function getEffectivePieLabelMode(preview, chart, options = {}) {
        if (options.compact) {
            return "hidden";
        }

        const requestedMode = chart?.pie_label_mode || "legend";
        if (requestedMode === "legend") {
            return "hidden";
        }

        const values = preview?.series?.[0]?.values || [];
        const total = values.reduce((sum, value) => sum + Number(value || 0), 0);
        const sectorLabelMode = chart?.pie_sector_label_mode || "name-percent";
        const labels = (preview?.categories || []).map((category, index) =>
            buildPieSectorLabelText(category, values[index] || 0, total, sectorLabelMode, false)
        );
        const longestLabelLength = Math.max(0, ...labels.map((label) => String(label).length));
        const totalLabelLength = labels.reduce((sum, label) => sum + String(label).length, 0);

        // Fallback hides only sector labels; legend remains controlled independently.
        if (labels.length > 6 || longestLabelLength > 28 || totalLabelLength > 120) {
            return "hidden";
        }

        return requestedMode === "outside-with-lines" ? "outside-with-lines" : "outside";
    }

    function buildSeriesLegendMarkup(series) {
        if (!series?.length) {
            return "";
        }
        return series.map((item) => ({
            color: item.color || "#b7791f",
            label: item.name || "Серия",
        }));
    }

    function buildBarChartSvg(preview, options = {}) {
        const categories = preview.categories || [];
        const series = preview.series || [];
        const compact = Boolean(options.compact);
        const width = compact ? 320 : 360;
        const height = compact ? 150 : 190;
        const padding = compact ? { top: 10, right: 8, bottom: 24, left: 28 } : { top: 14, right: 10, bottom: 34, left: 38 };
        const plotWidth = width - padding.left - padding.right;
        const plotHeight = height - padding.top - padding.bottom;
        const allValues = series.flatMap((item) => item.values || []);
        const maxValue = Math.max(1, ...allValues, 0);
        const bandWidth = categories.length ? plotWidth / categories.length : plotWidth;
        const seriesCount = Math.max(series.length, 1);
        const barGroupWidth = bandWidth * 0.72;
        const barWidth = Math.max(compact ? 4 : 8, barGroupWidth / seriesCount - (compact ? 2 : 4));
        const labelStep = compact && categories.length > 6 ? Math.ceil(categories.length / 4) : 1;

        const bars = [];
        const labels = [];
        categories.forEach((category, categoryIndex) => {
            const groupX = padding.left + categoryIndex * bandWidth + (bandWidth - barGroupWidth) / 2;
            if (!compact || categoryIndex % labelStep === 0 || categoryIndex === categories.length - 1) {
                labels.push(`<text x="${groupX + barGroupWidth / 2}" y="${height - (compact ? 6 : 10)}" text-anchor="middle">${escapeHtml(compact ? buildCompactAxisLabel(category) : shortenAxisLabel(category))}</text>`);
            }
            series.forEach((item, seriesIndex) => {
                const value = Number(item.values?.[categoryIndex] || 0);
                const barHeight = maxValue ? (value / maxValue) * plotHeight : 0;
                const x = groupX + seriesIndex * (barWidth + (compact ? 2 : 4));
                const y = padding.top + plotHeight - barHeight;
                bars.push(`<rect x="${x}" y="${y}" width="${barWidth}" height="${barHeight}" rx="${compact ? 2 : 4}" fill="${escapeHtml(item.color || "#b7791f")}"></rect>`);
            });
        });

        return `
            <svg viewBox="0 0 ${width} ${height}" class="analysis-chart-svg" aria-label="Столбчатая диаграмма">
                <line x1="${padding.left}" y1="${padding.top}" x2="${padding.left}" y2="${padding.top + plotHeight}" stroke="#c7b9a6" />
                <line x1="${padding.left}" y1="${padding.top + plotHeight}" x2="${width - padding.right}" y2="${padding.top + plotHeight}" stroke="#c7b9a6" />
                <text x="${padding.left - (compact ? 5 : 8)}" y="${padding.top + 8}" text-anchor="end">${escapeHtml(compact ? formatCompactPreviewNumber(maxValue) : formatPreviewNumber(maxValue))}</text>
                <text x="${padding.left - (compact ? 5 : 8)}" y="${padding.top + plotHeight}" text-anchor="end">0</text>
                ${bars.join("")}
                ${labels.join("")}
            </svg>
        `;
    }

    function buildLineChartSvg(preview, options = {}) {
        const categories = preview.categories || [];
        const series = preview.series || [];
        const compact = Boolean(options.compact);
        const width = compact ? 320 : 360;
        const height = compact ? 150 : 190;
        const padding = compact ? { top: 10, right: 8, bottom: 24, left: 28 } : { top: 14, right: 10, bottom: 34, left: 38 };
        const plotWidth = width - padding.left - padding.right;
        const plotHeight = height - padding.top - padding.bottom;
        const allValues = series.flatMap((item) => item.values || []);
        const maxValue = Math.max(1, ...allValues, 0);
        const stepX = categories.length > 1 ? plotWidth / (categories.length - 1) : plotWidth / 2;
        const labelStep = compact && categories.length > 6 ? Math.ceil(categories.length / 4) : 1;

        const polylines = series
            .map((item) => {
                const points = (item.values || []).map((value, index) => {
                    const x = padding.left + (categories.length > 1 ? index * stepX : plotWidth / 2);
                    const y = padding.top + plotHeight - (Number(value || 0) / maxValue) * plotHeight;
                    return `${x},${y}`;
                });
                const dots = (item.values || [])
                    .map((value, index) => {
                        const x = padding.left + (categories.length > 1 ? index * stepX : plotWidth / 2);
                        const y = padding.top + plotHeight - (Number(value || 0) / maxValue) * plotHeight;
                        return `<circle cx="${x}" cy="${y}" r="${compact ? 2.25 : 3.5}" fill="${escapeHtml(item.color || "#b7791f")}"></circle>`;
                    })
                    .join("");
                return `<polyline fill="none" stroke="${escapeHtml(item.color || "#b7791f")}" stroke-width="${compact ? 2 : 2.5}" points="${points.join(" ")}"></polyline>${dots}`;
            })
            .join("");

        const labels = categories
            .map((category, index) => {
                if (compact && index % labelStep !== 0 && index !== categories.length - 1) {
                    return "";
                }
                const x = padding.left + (categories.length > 1 ? index * stepX : plotWidth / 2);
                return `<text x="${x}" y="${height - (compact ? 6 : 10)}" text-anchor="middle">${escapeHtml(compact ? buildCompactAxisLabel(category) : shortenAxisLabel(category))}</text>`;
            })
            .join("");

        return `
            <svg viewBox="0 0 ${width} ${height}" class="analysis-chart-svg" aria-label="Линейный график">
                <line x1="${padding.left}" y1="${padding.top}" x2="${padding.left}" y2="${padding.top + plotHeight}" stroke="#c7b9a6" />
                <line x1="${padding.left}" y1="${padding.top + plotHeight}" x2="${width - padding.right}" y2="${padding.top + plotHeight}" stroke="#c7b9a6" />
                <text x="${padding.left - (compact ? 5 : 8)}" y="${padding.top + 8}" text-anchor="end">${escapeHtml(compact ? formatCompactPreviewNumber(maxValue) : formatPreviewNumber(maxValue))}</text>
                <text x="${padding.left - (compact ? 5 : 8)}" y="${padding.top + plotHeight}" text-anchor="end">0</text>
                ${polylines}
                ${labels}
            </svg>
        `;
    }

    function buildPieChartSvg(preview, chart, options = {}) {
        const categories = preview.categories || [];
        const values = preview.series?.[0]?.values || [];
        const total = values.reduce((sum, value) => sum + Number(value || 0), 0);
        const compact = Boolean(options.compact);
        const isQuad = options.layoutMode === "quad";
        const chartIndex = Number.isFinite(chart?.position) ? chart.position : 0;
        const effectivePieLabelMode = getEffectivePieLabelMode(preview, chart, options);
        const pieValueMode = chart?.pie_value_mode || "value";
        const centerText = buildPieCenterContent(total, chart, compact);
        const sectorLabelMode = chart?.pie_sector_label_mode || "name-percent";
        const labelOffsets = new Map(normalizePieSectorLabelOffsets(chart?.pie_sector_label_offsets).map((item) => [item.index, item]));
        const isSideLegend = options.legendPosition === "right" || options.legendPosition === "left";
        const width = compact
            ? 220
            : isQuad
            ? (effectivePieLabelMode === "hidden" ? 248 : (isSideLegend ? 264 : 308))
                : (effectivePieLabelMode === "hidden" ? 220 : 360);
        const height = compact ? 132 : isQuad ? (isSideLegend ? 156 : 182) : 220;
        const cx = compact
            ? width / 2
            : isQuad
                ? (effectivePieLabelMode === "hidden" ? width / 2 : (isSideLegend ? 96 : 114))
                : (effectivePieLabelMode === "hidden" ? 92 : 118);
        const cy = compact ? 68 : isQuad ? (isSideLegend ? 78 : 96) : 108;
        const radius = compact ? 42 : isQuad ? (effectivePieLabelMode === "hidden" ? 54 : (isSideLegend ? 44 : 56)) : (effectivePieLabelMode === "hidden" ? 62 : 64);

        if (!total) {
            return "";
        }

        let angle = -Math.PI / 2;
        const labelMarks = [];
        const leftLabels = [];
        const rightLabels = [];
        const slices = values
            .map((value, index) => {
                const numericValue = Number(value || 0);
                const ratio = total ? numericValue / total : 0;
                const nextAngle = angle + ratio * Math.PI * 2;
                const x1 = cx + Math.cos(angle) * radius;
                const y1 = cy + Math.sin(angle) * radius;
                const x2 = cx + Math.cos(nextAngle) * radius;
                const y2 = cy + Math.sin(nextAngle) * radius;
                const largeArc = nextAngle - angle > Math.PI ? 1 : 0;
                const color = pickPreviewColor(index, chart.color || "#b7791f");
                const midAngle = angle + (nextAngle - angle) / 2;
                if (!compact && effectivePieLabelMode !== "hidden") {
                    const textRadius = effectivePieLabelMode === "outside-with-lines"
                        ? radius + (isQuad ? (isSideLegend ? 20 : 26) : 38)
                        : radius + (isQuad ? (isSideLegend ? 12 : 18) : 24);
                    const anchor = Math.cos(midAngle) >= 0 ? "start" : "end";
                    const text = buildPieSectorLabelText(categories[index], numericValue, total, sectorLabelMode, false);
                    const shortenedText = shortenAxisLabel(text, isQuad ? (isSideLegend ? 14 : 17) : 22);
                    const approxCharWidth = isQuad ? 5.0 : 6.1;
                    const textWidth = shortenedText.length * approxCharWidth;
                    const edgePadding = isQuad ? 8 : 12;
                    const rawLabelX = cx + Math.cos(midAngle) * textRadius;
                    const labelX = anchor === "start"
                        ? Math.min(rawLabelX, width - textWidth - edgePadding)
                        : Math.max(rawLabelX, textWidth + edgePadding);
                    const labelY = cy + Math.sin(midAngle) * textRadius;
                    const entry = {
                        sectorIndex: index,
                        x: labelX,
                        y: labelY,
                        anchor,
                        text: shortenedText,
                        textWidth,
                        textHeight: isQuad ? 12.4 : 12,
                        withLine: effectivePieLabelMode === "outside-with-lines",
                        lineStartX: cx + Math.cos(midAngle) * (radius - 2),
                        lineStartY: cy + Math.sin(midAngle) * (radius - 2),
                        lineBendX: cx + Math.cos(midAngle) * (radius + (isQuad ? 7 : 10)),
                        lineBendY: cy + Math.sin(midAngle) * (radius + (isQuad ? 7 : 10)),
                    };
                    if (anchor === "start") {
                        rightLabels.push(entry);
                    } else {
                        leftLabels.push(entry);
                    }
                }
                angle = nextAngle;
                return `<path d="M ${cx} ${cy} L ${x1} ${y1} A ${radius} ${radius} 0 ${largeArc} 1 ${x2} ${y2} Z" fill="${escapeHtml(color)}"></path>`;
            })
            .join("");

        if (!compact && effectivePieLabelMode !== "hidden") {
            const minY = isQuad ? (isSideLegend ? 18 : 12) : 16;
            const maxY = isQuad ? (isSideLegend ? height - 52 : height - 12) : height - 18;
            [leftLabels, rightLabels].forEach((items) => {
                distributePieLabels(items, isQuad ? (isSideLegend ? 12 : 10) : 14, minY, maxY).forEach((item) => {
                    const edgePadding = isQuad ? 8 : 12;
                    const minX = item.anchor === "start" ? edgePadding : item.textWidth + edgePadding;
                    const maxX = item.anchor === "start" ? width - item.textWidth - edgePadding : width - edgePadding;
                    const manualOffset = labelOffsets.get(item.sectorIndex);
                    const currentX = clamp(item.x + width * ((manualOffset?.dx || 0) / 100), minX, maxX);
                    const currentY = clamp(item.y + height * ((manualOffset?.dy || 0) / 100), minY, maxY);
                    const dragAttrs = `data-analysis-action="drag-pie-sector-label" data-chart-index="${escapeHtml(String(chartIndex))}" data-sector-index="${escapeHtml(String(item.sectorIndex))}" data-anchor="${escapeHtml(item.anchor)}" data-base-x="${escapeHtml(String(item.x))}" data-base-y="${escapeHtml(String(item.y))}" data-edge-padding="${escapeHtml(String(edgePadding))}" data-text-width="${escapeHtml(String(item.textWidth))}" data-text-height="${escapeHtml(String(item.textHeight))}" data-min-x="${escapeHtml(String(minX))}" data-max-x="${escapeHtml(String(maxX))}" data-min-y="${escapeHtml(String(minY))}" data-max-y="${escapeHtml(String(maxY))}"`;
                    if (item.withLine) {
                        const lineEndX = currentX + (item.anchor === "start" ? (isQuad ? -4 : -6) : (isQuad ? 4 : 6));
                        labelMarks.push(`
                            <g class="analysis-pie-sector-label-group" data-analysis-pie-sector-label-group="true" data-chart-index="${escapeHtml(String(chartIndex))}" data-sector-index="${escapeHtml(String(item.sectorIndex))}" data-anchor="${escapeHtml(item.anchor)}">
                                <polyline class="analysis-pie-sector-label-line" points="${item.lineStartX},${item.lineStartY} ${item.lineBendX},${item.lineBendY} ${lineEndX},${currentY}" fill="none" stroke="#9b8a75" stroke-width="1"></polyline>
                                <text class="analysis-pie-sector-label" x="${currentX}" y="${currentY}" text-anchor="${item.anchor}" dominant-baseline="middle" ${dragAttrs}>${escapeHtml(item.text)}</text>
                            </g>
                        `);
                        return;
                    }
                    labelMarks.push(`
                        <g class="analysis-pie-sector-label-group" data-analysis-pie-sector-label-group="true" data-chart-index="${escapeHtml(String(chartIndex))}" data-sector-index="${escapeHtml(String(item.sectorIndex))}" data-anchor="${escapeHtml(item.anchor)}">
                            <text class="analysis-pie-sector-label" x="${currentX}" y="${currentY}" text-anchor="${item.anchor}" dominant-baseline="middle" ${dragAttrs}>${escapeHtml(item.text)}</text>
                        </g>
                    `);
                });
            });
        }

        return `
            <svg viewBox="0 0 ${width} ${height}" class="analysis-chart-svg" aria-label="Круговая диаграмма">
                ${slices}
                <circle cx="${cx}" cy="${cy}" r="${compact ? 18 : (isQuad ? 20 : 24)}" fill="#fffaf2"></circle>
                ${centerText ? `<text x="${cx}" y="${cy + 4}" text-anchor="middle">${escapeHtml(centerText)}</text>` : ""}
                ${labelMarks.join("")}
            </svg>
        `;
    }

    function categoryLabel(value) {
        return String(value || "Без группы");
    }

    function getLegendItems(preview, chart, options = {}) {
        if (chart.chart_type === "pie") {
            const categories = preview.categories || [];
            const values = preview.series?.[0]?.values || [];
            const compact = Boolean(options.compact);
            const total = values.reduce((sum, value) => sum + Number(value || 0), 0);
            const pieValueMode = chart?.pie_value_mode || "value";
            return categories.map((category, index) => ({
                color: pickPreviewColor(index, chart.color || "#b7791f"),
                label: compact
                    ? `${shortenAxisLabel(category, 10)}: ${buildPieValueText(values[index] || 0, total, pieValueMode, true)}`
                    : `${category}: ${buildPieValueText(values[index] || 0, total, pieValueMode, false)}`,
            }));
        }

        return buildSeriesLegendMarkup(preview.series || []);
    }

    function buildLegendMarkup(items, position) {
        if (!items?.length || position === "hidden") {
            return "";
        }

        return `
            <div class="analysis-chart-legend analysis-chart-legend-${escapeHtml(position)}">
                ${items
                    .map(
                        (item) => `
                            <span class="analysis-chart-legend-item">
                                <span class="analysis-chart-legend-swatch" style="background:${escapeHtml(item.color || "#b7791f")}"></span>
                                <span>${escapeHtml(item.label || item.name || "Серия")}</span>
                            </span>
                        `
                    )
                    .join("")}
            </div>
        `;
    }

    function pickPreviewColor(index, baseColor) {
        const palette = [baseColor || "#b7791f", "#0f766e", "#c2410c", "#2563eb", "#7c3aed", "#be123c"];
        return palette[index % palette.length];
    }

    function buildAggregationTableMarkup(preview) {
        const rows = preview.table_rows || [];
        if (!rows.length) {
            return "";
        }
        const headers = Object.keys(rows[0]);
        return `
            <div class="analysis-preview-table-wrap">
                <table class="analysis-preview-table">
                    <thead><tr>${headers.map((header) => `<th>${escapeHtml(header)}</th>`).join("")}</tr></thead>
                    <tbody>
                        ${rows
                            .slice(0, 8)
                            .map(
                                (row) => `<tr>${headers.map((header) => `<td>${escapeHtml(formatPreviewNumber(row[header]))}</td>`).join("")}</tr>`
                            )
                            .join("")}
                    </tbody>
                </table>
            </div>
        `;
    }

    function buildChartPreviewMarkup(preview, chart, options = {}) {
        const current = preview || { state: "idle", message: "Подготовьте конфигурацию графика." };
        const renderOptions = getChartRenderOptions(options, chart);
        const summary = current.summary || {};
        const warningMarkup = renderOptions.hideWarnings
            ? ""
            : (current.warnings || [])
            .map((warning) => `<div class="analysis-chart-warning">${escapeHtml(warning)}</div>`)
            .join("");

        if (current.state !== "ready") {
            return `
                <div class="analysis-chart-preview is-${escapeHtml(current.state || "idle")}">
                    <div class="analysis-chart-preview-state">${escapeHtml(current.message || "Подготовьте конфигурацию графика.")}</div>
                    ${warningMarkup}
                </div>
            `;
        }

        let visualizationMarkup = "";
        if (chart.chart_type === "bar") {
            visualizationMarkup = buildBarChartSvg(current, renderOptions);
        } else if (chart.chart_type === "line") {
            visualizationMarkup = buildLineChartSvg(current, renderOptions);
        } else if (chart.chart_type === "pie") {
            visualizationMarkup = buildPieChartSvg(current, chart, renderOptions);
        } else {
            visualizationMarkup = buildAggregationTableMarkup(current);
        }
        const legendItems = getLegendItems(current, chart, renderOptions);
        const effectivePieLabelMode = chart.chart_type === "pie"
            ? getEffectivePieLabelMode(current, chart, renderOptions)
            : "hidden";
        const effectiveLegendPosition = renderOptions.legendPosition;
        const legendMarkup = buildLegendMarkup(legendItems, effectiveLegendPosition);
        const isInsideLegend = effectiveLegendPosition === "inside-top-right" || effectiveLegendPosition === "inside-top-left";
        const isLeftLegend = effectiveLegendPosition === "left";
        const isRightLegend = effectiveLegendPosition === "right";
        const isBottomLegend = effectiveLegendPosition === "bottom";
        const sideLegendClass = isLeftLegend ? "has-side-legend has-side-legend-left" : isRightLegend ? "has-side-legend has-side-legend-right" : "";

        const annotationsMarkup = (chart.annotations || [])
            .map(
                (annotation) => `
                    <button
                        type="button"
                        class="analysis-chart-annotation"
                        data-analysis-action="drag-annotation"
                        data-chart-index="${escapeHtml(String(chart.position ?? 0))}"
                        data-annotation-id="${escapeHtml(annotation.id)}"
                        style="left:${Number(annotation.x) || 0}%;top:${Number(annotation.y) || 0}%"
                    >${escapeHtml(annotation.text)}</button>
                `
            )
            .join("");

        return `
            <div class="analysis-chart-preview is-ready ${renderOptions.compact ? "is-compact" : ""}">
                ${renderOptions.hideHead ? "" : `
                <div class="analysis-chart-preview-head">
                    <strong>${escapeHtml(chart.legend || "Предпросмотр графика")}</strong>
                    <span>${escapeHtml((summary.aggregation || chart.agg_func || "count").toUpperCase())}</span>
                </div>`}
                <div class="analysis-chart-preview-body ${sideLegendClass}">
                    ${isLeftLegend ? legendMarkup : ""}
                    <div class="analysis-chart-canvas" data-chart-index="${escapeHtml(String(chart.position ?? 0))}">
                        <div class="analysis-chart-visual">${visualizationMarkup}</div>
                        ${isInsideLegend ? legendMarkup : ""}
                        <div class="analysis-chart-annotations-layer">${annotationsMarkup}</div>
                    </div>
                    ${isRightLegend ? legendMarkup : ""}
                </div>
                ${isBottomLegend ? legendMarkup : ""}
                ${renderOptions.hideMeta ? "" : `<div class="analysis-chart-preview-meta">Точек: ${escapeHtml(String(summary.points || 0))}${summary.skipped_rows ? ` • Пропущено строк: ${escapeHtml(String(summary.skipped_rows))}` : ""}</div>`}
                ${warningMarkup}
            </div>
        `;
    }

    function buildSheetCardMarkup(chart, preview, index, options = {}) {
        const layoutMode = options.layoutMode || "sheet";
        const isPair = layoutMode === "pair";
        const isQuad = layoutMode === "quad";
        const title = escapeHtml(getChartDisplayTitle(chart, index));
        const compatibility = getChartCompatibility(chart);
        return `
            <article class="analysis-sheet-card ${isPair ? "is-pair-card" : ""} ${isQuad ? "is-quad-card" : ""} ${compatibility.ok ? "" : "is-incompatible"}">
                <div class="analysis-sheet-card-head">
                    <div>
                        <h4>${title}</h4>
                        <div class="analysis-sheet-card-meta">${escapeHtml(chart.chart_type)} • ${escapeHtml(getSourceKindLabel(chart.source_kind || "none"))}</div>
                    </div>
                    <span class="analysis-sheet-card-status ${compatibility.ok ? "" : "warning"}">${compatibility.ok ? "Готов" : "Нужна настройка"}</span>
                </div>
                <div class="analysis-sheet-card-body">
                    ${buildChartPreviewMarkup(preview, chart, (isPair || isQuad)
                        ? { layoutMode, hideWarnings: true, hideHead: true, hideMeta: true }
                        : { layoutMode, compact: true, hideWarnings: true, hideHead: true, hideMeta: true })}
                </div>
            </article>
        `;
    }

    function sanitizeExportFileName(value, fallback) {
        const normalized = String(value || "")
            .trim()
            .replace(/[^\w.-]+/g, "_")
            .replace(/^_+|_+$/g, "");
        return normalized || fallback;
    }

    function getChartDisplayTitle(chart, index) {
        return (chart.legend || "").trim() || `График_${index + 1}`;
    }

    function downloadTextFile(content, fileName, mimeType) {
        const blob = new Blob([content], { type: mimeType });
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = fileName;
        document.body.appendChild(link);
        link.click();
        link.remove();
        window.setTimeout(() => URL.revokeObjectURL(url), 0);
    }

    function buildCsvContent(rows) {
        if (!rows.length) {
            return "";
        }
        const headers = Object.keys(rows[0] || {});
        const escapeCsv = (value) => `"${String(value ?? "").replaceAll('"', '""')}"`;
        return [
            headers.map(escapeCsv).join(";"),
            ...rows.map((row) => headers.map((header) => escapeCsv(row[header])).join(";")),
        ].join("\r\n");
    }

    function buildExportStyles() {
        return `
            body { font-family: "Segoe UI", Tahoma, Geneva, Verdana, sans-serif; color: #2c241b; margin: 24px; }
            .report-head { margin-bottom: 20px; }
            .report-head h1 { margin: 0 0 8px; font-size: 24px; }
            .report-meta { display: grid; gap: 6px; margin-bottom: 16px; color: #6f6458; font-size: 13px; }
            .report-section { margin-bottom: 28px; border: 1px solid rgba(74, 56, 38, 0.14); border-radius: 14px; padding: 16px; background: #fffaf2; }
            .analysis-chart-preview { display: flex; flex-direction: column; gap: 8px; padding: 10px; border-radius: 10px; border: 1px solid rgba(74, 56, 38, 0.12); background: linear-gradient(180deg, rgba(255, 255, 255, 0.92), rgba(248, 242, 232, 0.88)); }
            .analysis-chart-preview-state { padding: 14px 12px; border-radius: 8px; text-align: center; color: #6f6458; background: rgba(255, 255, 255, 0.72); border: 1px dashed rgba(74, 56, 38, 0.16); }
            .analysis-chart-preview-head, .analysis-chart-preview-meta { display: flex; align-items: center; justify-content: space-between; gap: 8px; font-size: 12px; }
            .analysis-chart-preview-head strong { font-size: 13px; }
            .analysis-chart-preview-head span, .analysis-chart-preview-meta, .report-meta { color: #6f6458; }
            .analysis-chart-preview-body.has-side-legend { display: grid; grid-template-columns: minmax(150px, 180px) minmax(0, 1fr); gap: 10px; align-items: start; }
            .analysis-chart-preview-body.has-side-legend-right { grid-template-columns: minmax(0, 1fr) minmax(150px, 180px); }
            .analysis-chart-canvas { position: relative; min-height: 180px; }
            .analysis-chart-svg { width: 100%; height: auto; display: block; }
            .analysis-chart-svg text { fill: #6f6458; font-size: 10px; font-family: "Segoe UI", Tahoma, Geneva, Verdana, sans-serif; }
            .analysis-chart-legend { display: flex; flex-wrap: wrap; gap: 8px 12px; margin-top: 8px; }
            .analysis-chart-legend-left, .analysis-chart-legend-right, .analysis-chart-legend-bottom, .analysis-chart-legend-inside-top-right, .analysis-chart-legend-inside-top-left { padding: 8px 10px; border-radius: 10px; background: rgba(255, 250, 242, 0.94); border: 1px solid rgba(74, 56, 38, 0.12); }
            .analysis-chart-legend-left, .analysis-chart-legend-right { flex-direction: column; align-items: flex-start; margin-top: 0; }
            .analysis-chart-legend-bottom { margin-top: 10px; }
            .analysis-chart-legend-inside-top-right, .analysis-chart-legend-inside-top-left { position: absolute; top: 8px; max-width: 44%; z-index: 2; }
            .analysis-chart-legend-inside-top-right { right: 8px; }
            .analysis-chart-legend-inside-top-left { left: 8px; }
            .analysis-chart-legend-item { display: inline-flex; align-items: center; gap: 6px; font-size: 12px; color: #6f6458; }
            .analysis-chart-legend-swatch { width: 10px; height: 10px; border-radius: 999px; flex: 0 0 auto; }
            .analysis-preview-table-wrap { overflow: auto; border: 1px solid rgba(74, 56, 38, 0.12); border-radius: 8px; background: #fff; }
            .analysis-preview-table { width: 100%; border-collapse: collapse; font-size: 12px; }
            .analysis-preview-table th, .analysis-preview-table td { padding: 7px 8px; border-bottom: 1px solid rgba(74, 56, 38, 0.1); text-align: left; }
            .analysis-preview-table th { background: #f8f2e8; }
            .analysis-chart-warning { padding: 8px 10px; border-radius: 8px; background: rgba(217, 119, 6, 0.1); color: #b45309; font-size: 12px; }
            .analysis-chart-comment-block { margin-top: 14px; padding: 12px 14px; border-radius: 10px; background: rgba(15, 118, 110, 0.06); border: 1px solid rgba(15, 118, 110, 0.14); }
            .analysis-chart-comment-block h3 { margin: 0 0 8px; font-size: 14px; }
            .analysis-chart-comment-block p { margin: 0; line-height: 1.5; }
            @media print { body { margin: 12mm; } .report-section { break-inside: avoid; } }
        `;
    }

    function buildChartMetaMarkup(chart, preview, index) {
        return `
            <div class="report-meta">
                <div><strong>График:</strong> ${escapeHtml(getChartDisplayTitle(chart, index))}</div>
                <div><strong>Тип:</strong> ${escapeHtml(chart.chart_type)}</div>
                <div><strong>Источник:</strong> ${escapeHtml(state.dataset.source_label || getSourceKindLabel(state.dataset.source_kind || "none"))}</div>
                <div><strong>Режим:</strong> ${escapeHtml(getSourceModeLabel(state.dataset.source_mode || "none"))}</div>
                <div><strong>Сформирован:</strong> ${escapeHtml(formatTimestamp(new Date().toISOString()))}</div>
                <div><strong>Точек:</strong> ${escapeHtml(String(preview?.summary?.points || 0))}</div>
            </div>
        `;
    }

    function buildChartCommentMarkup(chart) {
        const title = String(chart.comment_title || "").trim();
        const text = String(chart.comment_text || "").trim();
        if (!title && !text) {
            return "";
        }

        return `
            <section class="analysis-chart-comment-block">
                ${title ? `<h3>${escapeHtml(title)}</h3>` : ""}
                ${text ? `<p>${escapeHtml(text).replaceAll("\n", "<br>")}</p>` : ""}
            </section>
        `;
    }

    function buildChartExportSection(chart, preview, index) {
        return `
            <section class="report-section">
                ${buildChartMetaMarkup(chart, preview, index)}
                ${buildChartPreviewMarkup(preview, chart)}
                ${buildChartCommentMarkup(chart)}
            </section>
        `;
    }

    function buildExportDocument(title, sections) {
        return `<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${escapeHtml(title)}</title>
    <style>${buildExportStyles()}</style>
</head>
<body>
    <header class="report-head">
        <h1>${escapeHtml(title)}</h1>
        <div class="report-meta">
            <div><strong>Источник:</strong> ${escapeHtml(state.dataset.source_label || getSourceKindLabel(state.dataset.source_kind || "none"))}</div>
            <div><strong>Тип источника:</strong> ${escapeHtml(getSourceKindLabel(state.dataset.source_kind || "none"))}</div>
            <div><strong>Дата формирования:</strong> ${escapeHtml(formatTimestamp(new Date().toISOString()))}</div>
        </div>
    </header>
    ${sections.join("\n")}
</body>
</html>`;
    }

    function isPreviewExportable(preview) {
        return preview && preview.state === "ready";
    }

    function saveChartPreview(index) {
        const chart = state.draft.charts[index];
        const preview = state.chartPreviews[index];
        if (!chart || !isPreviewExportable(preview)) {
            setMessage("Сохранение доступно только для корректно построенного графика.", "warning");
            return;
        }

        const title = getChartDisplayTitle(chart, index);
        const html = buildExportDocument(title, [buildChartExportSection(chart, preview, index)]);
        const fileName = `${sanitizeExportFileName(title, `analysis_chart_${index + 1}`)}.html`;
        downloadTextFile(html, fileName, "text/html;charset=utf-8");
        setMessage(`График сохранен в локальный HTML-файл: ${fileName}.`, "success");
    }

    function printChartPreview(index) {
        const chart = state.draft.charts[index];
        const preview = state.chartPreviews[index];
        if (!chart || !isPreviewExportable(preview)) {
            setMessage("Печать доступна только для корректно построенного графика.", "warning");
            return;
        }

        const title = getChartDisplayTitle(chart, index);
        const html = buildExportDocument(title, [buildChartExportSection(chart, preview, index)]);
        const printWindow = window.open("", "_blank", "width=1200,height=800");
        if (!printWindow) {
            setMessage("Браузер заблокировал окно печати. Разрешите всплывающие окна и повторите попытку.", "error");
            return;
        }

        printWindow.document.open();
        printWindow.document.write(html);
        printWindow.document.close();
        printWindow.focus();
        window.setTimeout(() => {
            printWindow.print();
        }, 200);
    }

    async function exportChartTable(index) {
        const chart = state.draft.charts[index];
        const preview = state.chartPreviews[index];
        if (!chart || !isPreviewExportable(preview) || !preview.table_rows?.length) {
            setMessage("Для текущего графика нет агрегированной таблицы для выгрузки.", "warning");
            return;
        }

        const title = getChartDisplayTitle(chart, index);
        const headers = Object.keys(preview.table_rows[0] || {});
        try {
            const response = await apiJson("/api/analysis/export-table", {
                method: "POST",
                body: JSON.stringify({
                    columns: headers,
                    rows: preview.table_rows,
                    file_stem: sanitizeExportFileName(title, `analysis_table_${index + 1}`),
                    sheet_title: title,
                }),
            });
            setMessage(`Таблица Анализа сохранена: ${response.file_name}.`, "success");
        } catch (error) {
            const fileName = `${sanitizeExportFileName(title, `analysis_table_${index + 1}`)}.csv`;
            downloadTextFile(`\ufeff${buildCsvContent(preview.table_rows)}`, fileName, "text/csv;charset=utf-8");
            setMessage(`${error.error || "Не удалось сохранить XLSX."} Таблица выгружена локально как CSV: ${fileName}.`, "warning");
        }
    }

    function exportVisibleChartsReport() {
        const sections = state.draft.charts
            .map((chart, index) => ({ chart, preview: state.chartPreviews[index], index }))
            .filter((item) => !item.chart.is_hidden && isPreviewExportable(item.preview))
            .map((item) => buildChartExportSection(item.chart, item.preview, item.index));

        if (!sections.length) {
            setMessage("Для экспорта отчета нет готовых видимых графиков.", "warning");
            return;
        }

        const reportTitle = state.draft.name ? `Отчет_${state.draft.name}` : "analysis_report";
        const fileName = `${sanitizeExportFileName(reportTitle, "analysis_report")}.html`;
        downloadTextFile(buildExportDocument(reportTitle, sections), fileName, "text/html;charset=utf-8");
        setMessage(`Отчет по видимым графикам сохранен: ${fileName}.`, "success");
    }

    function applyBackendAnalysisWidth() {
        const width = Number(state.userState.left_panel_width || 260);
        document.documentElement.style.setProperty("--analysis-left-width", `${width}px`);
        try {
            localStorage.setItem("analysis-left-width", String(width));
        } catch {}
    }

    function getSelectedType() {
        return state.types.find((item) => item.id === state.draft.selectedTypeId) || null;
    }

    function syncDraftFromSelection() {
        const current = state.types.find((item) => item.id === state.userState.selected_analysis_type_id) || null;
        const persistedDraft = state.userState.draft_state;

        if (
            current
            && persistedDraft
            && Number(persistedDraft.analysis_type_id) === Number(current.id)
            && Array.isArray(persistedDraft.charts)
        ) {
            state.draft.selectedTypeId = current.id;
            state.draft.name = persistedDraft.name || current.name;
            state.draft.charts = persistedDraft.charts
                .slice()
                .sort((left, right) => left.position - right.position)
                .map((chart, index) => ({ ...chart, ui_collapsed: false, position: index }));
            return;
        }

        state.draft.selectedTypeId = current ? current.id : null;
        state.draft.name = current ? current.name : "";
        state.draft.charts = current
            ? current.charts
                .slice()
                .sort((left, right) => left.position - right.position)
                .map((chart, index) => ({ ...chart, ui_collapsed: false, position: index }))
            : [];
    }

    function toPersistedChart(chart, index) {
        return {
            id: chart.id,
            chart_type: chart.chart_type,
            source_kind: chart.source_kind,
            x_field: chart.x_field,
            y_field: chart.y_field,
            group_field: chart.group_field,
            agg_func: chart.agg_func,
            color: chart.color,
            legend: chart.legend,
            legend_position: chart.legend_position,
            pie_label_mode: chart.pie_label_mode,
            pie_value_mode: chart.pie_value_mode,
            pie_center_mode: chart.pie_center_mode,
            pie_sector_label_mode: chart.pie_sector_label_mode,
            labels: chart.labels,
            comment_title: chart.comment_title,
            comment_text: chart.comment_text,
            annotations: (chart.annotations || []).map((annotation) => ({
                id: annotation.id,
                text: annotation.text,
                x: Number(annotation.x) || 0,
                y: Number(annotation.y) || 0,
            })),
            pie_sector_label_offsets: normalizePieSectorLabelOffsets(chart.pie_sector_label_offsets),
            is_hidden: chart.is_hidden,
            position: index,
        };
    }

    function getChartAt(index) {
        return state.draft.charts[index] || null;
    }

    function addAnnotation(index) {
        const chart = getChartAt(index);
        if (!chart) {
            return;
        }

        chart.annotations = [
            ...(chart.annotations || []),
            { id: createAnnotationId(), text: `Надпись ${((chart.annotations || []).length || 0) + 1}`, x: 8, y: 8 },
        ];
        state.userState.draft_state = buildDraftStatePayload();
        scheduleUserStateSave({ draft_state: state.userState.draft_state });
        renderAll({ refreshPreviews: false });
    }

    function removeAnnotation(index, annotationId) {
        const chart = getChartAt(index);
        if (!chart) {
            return;
        }

        chart.annotations = (chart.annotations || []).filter((annotation) => annotation.id !== annotationId);
        state.userState.draft_state = buildDraftStatePayload();
        scheduleUserStateSave({ draft_state: state.userState.draft_state });
        renderAll({ refreshPreviews: false });
    }

    function updateAnnotationText(index, annotationId, text) {
        const chart = getChartAt(index);
        const annotation = chart?.annotations?.find((item) => item.id === annotationId);
        if (!annotation) {
            return;
        }

        annotation.text = String(text || "").trimStart().slice(0, 160);
        state.userState.draft_state = buildDraftStatePayload();
    }

    function setAnnotationPosition(index, annotationId, x, y) {
        const chart = getChartAt(index);
        const annotation = chart?.annotations?.find((item) => item.id === annotationId);
        if (!annotation) {
            return;
        }

        annotation.x = clamp(x, 0, 92);
        annotation.y = clamp(y, 0, 92);
        state.userState.draft_state = buildDraftStatePayload();
    }

    function setPieSectorLabelOffset(index, sectorIndex, dx, dy) {
        const chart = getChartAt(index);
        if (!chart) {
            return;
        }

        const offsets = normalizePieSectorLabelOffsets(chart.pie_sector_label_offsets);
        const nextOffset = {
            index: sectorIndex,
            dx: Math.round(clamp(dx, -100, 100) * 100) / 100,
            dy: Math.round(clamp(dy, -100, 100) * 100) / 100,
        };
        const existingIndex = offsets.findIndex((item) => item.index === sectorIndex);
        if (existingIndex >= 0) {
            offsets[existingIndex] = nextOffset;
        } else {
            offsets.push(nextOffset);
        }
        chart.pie_sector_label_offsets = offsets;
        state.userState.draft_state = buildDraftStatePayload();
    }

    function resetPieSectorLabelOffsets(index) {
        const chart = getChartAt(index);
        if (!chart) {
            return;
        }

        chart.pie_sector_label_offsets = [];
        state.userState.draft_state = buildDraftStatePayload();
        scheduleUserStateSave({ draft_state: state.userState.draft_state });
        renderAll({ refreshPreviews: false });
    }

    function syncAnnotationNodes(index, annotationId, x, y, text = null) {
        document.querySelectorAll(`[data-analysis-action='drag-annotation'][data-chart-index='${index}'][data-annotation-id='${annotationId}']`).forEach((node) => {
            if (x !== null && x !== undefined) {
                node.style.left = `${x}%`;
            }
            if (y !== null && y !== undefined) {
                node.style.top = `${y}%`;
            }
            if (text !== null) {
                node.textContent = text;
            }
        });
    }

    function syncPieSectorLabelNodes(index, sectorIndex, x, y) {
        document.querySelectorAll(`[data-analysis-pie-sector-label-group='true'][data-chart-index='${index}'][data-sector-index='${sectorIndex}']`).forEach((group) => {
            const textNode = group.querySelector(".analysis-pie-sector-label");
            const lineNode = group.querySelector(".analysis-pie-sector-label-line");
            const anchor = group.dataset.anchor || textNode?.dataset.anchor || "start";
            if (textNode) {
                textNode.setAttribute("x", String(x));
                textNode.setAttribute("y", String(y));
            }
            if (lineNode) {
                const points = (lineNode.getAttribute("points") || "").trim().split(/\s+/);
                if (points.length === 3) {
                    const lineEndX = x + (anchor === "start" ? -4 : 4);
                    points[2] = `${lineEndX},${y}`;
                    lineNode.setAttribute("points", points.join(" "));
                }
            }
        });
    }

    function startAnnotationDrag(index, annotationId, event) {
        if (event.button !== 0) {
            return;
        }

        const handle = event.target.closest("[data-analysis-action='drag-annotation']");
        const canvas = handle?.closest(".analysis-chart-canvas");
        if (!handle || !canvas) {
            return;
        }

        event.preventDefault();
        const rect = canvas.getBoundingClientRect();
        activeAnnotationDrag = {
            chartIndex: index,
            annotationId,
            rect,
        };
        window.addEventListener("mousemove", handleAnnotationDrag);
        window.addEventListener("mouseup", stopAnnotationDrag);
    }

    function startPieSectorLabelDrag(index, sectorIndex, event) {
        if (event.button !== 0) {
            return;
        }

        const handle = event.target.closest("[data-analysis-action='drag-pie-sector-label']");
        const svg = handle?.closest(".analysis-chart-svg");
        if (!handle || !svg) {
            return;
        }

        event.preventDefault();
        const rect = svg.getBoundingClientRect();
        const viewBox = svg.viewBox?.baseVal;
        if (!viewBox || !rect.width || !rect.height) {
            return;
        }

        activePieSectorLabelDrag = {
            chartIndex: index,
            sectorIndex,
            rect,
            viewBoxWidth: viewBox.width,
            viewBoxHeight: viewBox.height,
            baseX: Number(handle.dataset.baseX) || 0,
            baseY: Number(handle.dataset.baseY) || 0,
            anchor: handle.dataset.anchor || "start",
            textWidth: Number(handle.dataset.textWidth) || 56,
            textHeight: Number(handle.dataset.textHeight) || 12,
            edgePadding: Number(handle.dataset.edgePadding) || 8,
            minX: Number(handle.dataset.minX),
            maxX: Number(handle.dataset.maxX),
            minY: Number(handle.dataset.minY),
            maxY: Number(handle.dataset.maxY),
        };
        window.addEventListener("mousemove", handlePieSectorLabelDrag);
        window.addEventListener("mouseup", stopPieSectorLabelDrag);
    }

    function handleAnnotationDrag(event) {
        if (!activeAnnotationDrag) {
            return;
        }

        const { chartIndex, annotationId, rect } = activeAnnotationDrag;
        const x = ((event.clientX - rect.left) / Math.max(rect.width, 1)) * 100;
        const y = ((event.clientY - rect.top) / Math.max(rect.height, 1)) * 100;
        setAnnotationPosition(chartIndex, annotationId, x, y);
        syncAnnotationNodes(chartIndex, annotationId, clamp(x, 0, 92), clamp(y, 0, 92));
    }

    function handlePieSectorLabelDrag(event) {
        if (!activePieSectorLabelDrag) {
            return;
        }

        const {
            chartIndex,
            sectorIndex,
            rect,
            viewBoxWidth,
            viewBoxHeight,
            baseX,
            baseY,
            anchor,
            textWidth,
            textHeight,
            edgePadding,
            minX,
            maxX,
            minY,
            maxY,
        } = activePieSectorLabelDrag;
        const rawX = ((event.clientX - rect.left) / Math.max(rect.width, 1)) * viewBoxWidth;
        const rawY = ((event.clientY - rect.top) / Math.max(rect.height, 1)) * viewBoxHeight;
        const clampedMinX = Number.isFinite(minX) ? minX : (anchor === "start" ? edgePadding : textWidth + edgePadding);
        const clampedMaxX = Number.isFinite(maxX) ? maxX : (anchor === "start" ? viewBoxWidth - textWidth - edgePadding : viewBoxWidth - edgePadding);
        const clampedMinY = Number.isFinite(minY) ? minY : textHeight;
        const clampedMaxY = Number.isFinite(maxY) ? maxY : (viewBoxHeight - edgePadding);
        const x = clamp(rawX, clampedMinX, clampedMaxX);
        const y = clamp(rawY, clampedMinY, clampedMaxY);
        const dx = ((x - baseX) / Math.max(viewBoxWidth, 1)) * 100;
        const dy = ((y - baseY) / Math.max(viewBoxHeight, 1)) * 100;
        setPieSectorLabelOffset(chartIndex, sectorIndex, dx, dy);
        syncPieSectorLabelNodes(chartIndex, sectorIndex, x, y);
    }

    function stopAnnotationDrag() {
        if (!activeAnnotationDrag) {
            return;
        }

        activeAnnotationDrag = null;
        scheduleUserStateSave({ draft_state: state.userState.draft_state });
        window.removeEventListener("mousemove", handleAnnotationDrag);
        window.removeEventListener("mouseup", stopAnnotationDrag);
    }

    function stopPieSectorLabelDrag() {
        if (!activePieSectorLabelDrag) {
            return;
        }

        activePieSectorLabelDrag = null;
        scheduleUserStateSave({ draft_state: state.userState.draft_state });
        window.removeEventListener("mousemove", handlePieSectorLabelDrag);
        window.removeEventListener("mouseup", stopPieSectorLabelDrag);
    }

    function buildDraftStatePayload() {
        if (!state.draft.selectedTypeId) {
            return null;
        }
        return {
            analysis_type_id: state.draft.selectedTypeId,
            name: state.draft.name,
            charts: state.draft.charts.map((chart, index) => toPersistedChart(chart, index)),
        };
    }

    function queueUserStateSavePatch(patch) {
        if (!patch || typeof patch !== "object") {
            return;
        }

        pendingUserStatePatch = {
            ...(pendingUserStatePatch || {}),
            ...patch,
        };
    }

    function handleUserStateSaveError(error) {
        const message = error?.error || "Не удалось автоматически сохранить состояние вкладки Анализ.";
        setMessage(message, "warning");
    }

    async function flushUserStateSave(options = {}) {
        const keepalive = options.keepalive === true;
        window.clearTimeout(userStateSaveTimer);
        userStateSaveTimer = null;

        if (activeUserStateSavePromise) {
            await activeUserStateSavePromise;
        }

        if (!pendingUserStatePatch || !Object.keys(pendingUserStatePatch).length) {
            return null;
        }

        const patch = pendingUserStatePatch;
        pendingUserStatePatch = null;

        const request = apiJson("/api/analysis/state", {
            method: "PUT",
            keepalive,
            body: JSON.stringify(patch),
        });

        activeUserStateSavePromise = request;

        let response;
        try {
            response = await request;
            state.userState = { ...defaultUserState, ...(response.user_state || state.userState) };
            applyPersistedAnalysisUiState();
        } catch (error) {
            pendingUserStatePatch = {
                ...patch,
                ...(pendingUserStatePatch || {}),
            };
            throw error;
        } finally {
            if (activeUserStateSavePromise === request) {
                activeUserStateSavePromise = null;
            }
        }

        if (pendingUserStatePatch && !keepalive) {
            return flushUserStateSave(options);
        }

        return response;
    }

    function scheduleUserStateSave(patch, delay = 180) {
        queueUserStateSavePatch(patch);
        window.clearTimeout(userStateSaveTimer);
        userStateSaveTimer = window.setTimeout(() => {
            void flushUserStateSave().catch(handleUserStateSaveError);
        }, delay);
    }

    function renderTypeSelect() {
        const selectedId = state.draft.selectedTypeId ? String(state.draft.selectedTypeId) : "";
        elements.typeSelect.innerHTML = [
            '<option value="">Тип анализа</option>',
            ...state.types.map((item) => `<option value="${item.id}">${escapeHtml(item.name)}</option>`),
        ].join("");
        elements.typeSelect.value = selectedId;
        elements.deleteTypeButton.disabled = !state.draft.selectedTypeId;
        elements.saveTypeButton.disabled = !state.draft.selectedTypeId;
    }

    function getColumnOptionsMarkup(selectedValue) {
        const options = ['<option value="">Не выбрано</option>'];
        state.dataset.columns.forEach((column) => {
            options.push(
                `<option value="${escapeHtml(column.name)}" ${column.name === selectedValue ? "selected" : ""}>${escapeHtml(column.name)}</option>`
            );
        });
        return options.join("");
    }

    function renderViewMode() {
        const isPresentationView = state.viewMode === "sheet" || state.viewMode === "pair" || state.viewMode === "quad";
        const isSheet = state.viewMode === "sheet";
        const isPair = state.viewMode === "pair";
        const isQuad = state.viewMode === "quad";
        elements.shell?.classList.toggle("analysis-presentation-mode", isPresentationView);
        elements.mainView?.classList.toggle("hidden", isPresentationView);
        elements.sheetView?.classList.toggle("hidden", !isPresentationView);
        elements.sheetView?.setAttribute("data-layout", isPair ? "pair" : isQuad ? "quad" : "sheet");
        elements.sheetGrid?.setAttribute("data-layout", isPair ? "pair" : isQuad ? "quad" : "sheet");
        elements.editorViewButton?.classList.toggle("active", state.viewMode === "editor");
        elements.sheetViewButton?.classList.toggle("active", isSheet);
        elements.pairViewButton?.classList.toggle("active", isPair);
        elements.quadViewButton?.classList.toggle("active", isQuad);
        if (elements.sheetTitle) {
            elements.sheetTitle.textContent = isPair ? "Лист 1*1" : isQuad ? "Лист 4" : "Лист 2x2";
        }
        if (elements.sheetDescription) {
            elements.sheetDescription.textContent = isPair
                ? "Отдельный экран для одновременного просмотра двух графиков слева и справа."
                : isQuad
                    ? "Экран из четырех секторов: до четырех графиков одновременно по правилам расширенного рендера, близким к Листу 1*1."
                    : "Отдельный экран для одновременного просмотра до четырех графиков.";
        }
        updateSheetViewportLayout();
    }

    function updateSheetViewportLayout() {
        if (!elements.sheetView) {
            return;
        }

        if (state.viewMode !== "sheet" && state.viewMode !== "pair" && state.viewMode !== "quad") {
            elements.sheetView.style.removeProperty("height");
            elements.sheetView.style.removeProperty("max-height");
            elements.sheetView.style.removeProperty("--analysis-sheet-grid-height");
            elements.sheetView.style.removeProperty("--analysis-sheet-row-height");
            return;
        }

        const viewRect = elements.sheetView.getBoundingClientRect();
        const viewportHeight = window.innerHeight || document.documentElement.clientHeight || 0;
        const availableHeight = Math.max(360, Math.floor(viewportHeight - viewRect.top - 16));
        elements.sheetView.style.height = `${availableHeight}px`;
        elements.sheetView.style.maxHeight = `${availableHeight}px`;

        const headHeight = elements.sheetView.querySelector(".analysis-sheet-head")?.offsetHeight || 0;
        const computedStyle = window.getComputedStyle(elements.sheetView);
        const gapValue = Number.parseFloat(computedStyle.rowGap || computedStyle.gap || "0") || 0;
        const gridHeight = Math.max(220, Math.floor(availableHeight - headHeight - gapValue));
        const gridGap = Number.parseFloat(window.getComputedStyle(elements.sheetGrid || elements.sheetView).gap || "0") || 0;
        const rowHeight = state.viewMode === "pair"
            ? Math.max(220, gridHeight)
            : state.viewMode === "quad"
                ? Math.max(180, Math.floor((gridHeight - gridGap) / 2))
                : Math.max(140, Math.floor((gridHeight - gridGap) / 2));

        elements.sheetView.style.setProperty("--analysis-sheet-grid-height", `${gridHeight}px`);
        elements.sheetView.style.setProperty("--analysis-sheet-row-height", `${rowHeight}px`);
    }

    function scheduleSheetViewportLayout() {
        if (sheetLayoutFrame !== null) {
            cancelAnimationFrame(sheetLayoutFrame);
        }

        sheetLayoutFrame = requestAnimationFrame(() => {
            sheetLayoutFrame = null;
            updateSheetViewportLayout();
        });
    }

    function renderSheetView() {
        if (!elements.sheetGrid) {
            return;
        }

        const isPair = state.viewMode === "pair";
        const isQuad = state.viewMode === "quad";
        const maxCards = isPair ? 2 : 4;
        const cards = state.draft.charts.slice(0, maxCards);
        if (!cards.length) {
            elements.sheetGrid.innerHTML = `<div class="analysis-sheet-empty">${isPair
                ? "Добавьте до двух графиков, чтобы увидеть экран 1*1 с двумя карточками слева и справа."
                : isQuad
                    ? "Добавьте до четырех графиков, чтобы увидеть экран Лист 4 с четырьмя секторами и расширенным рендером."
                    : "Добавьте до четырех графиков, чтобы увидеть общий лист 2x2."}</div>`;
            scheduleSheetViewportLayout();
            return;
        }

        const markup = [];
        cards.forEach((chart, index) => {
            const preview = state.chartPreviews[index] || { state: "idle", message: "Подготовьте конфигурацию графика." };
            if (chart.is_hidden) {
                markup.push(`
                    <article class="analysis-sheet-card is-hidden-card">
                        <div class="analysis-sheet-card-head">
                            <div>
                                <h4>${escapeHtml(getChartDisplayTitle(chart, index))}</h4>
                                <div class="analysis-sheet-card-meta">График скрыт пользователем</div>
                            </div>
                        </div>
                        <div class="analysis-sheet-placeholder">Этот график скрыт и не показывается на листе.</div>
                    </article>
                `);
                return;
            }
            markup.push(buildSheetCardMarkup(chart, preview, index, { layoutMode: state.viewMode }));
        });

        while (markup.length < maxCards) {
            markup.push(`<div class="analysis-sheet-placeholder">${isPair ? "Свободное место для графика" : "Свободное место для графика"}</div>`);
        }

        elements.sheetGrid.innerHTML = markup.join("");
        scheduleSheetViewportLayout();
    }

    function getAnalysisColumnWidth(columnName) {
        const fallback = columnName === "__row_index__" ? 72 : 190;
        return Math.max(72, Number(state.table.columnWidths[columnName]) || fallback);
    }

    function renderTableColgroup(columns) {
        if (!elements.tableColgroup) {
            return;
        }

        const widths = [
            `<col style="width:${getAnalysisColumnWidth("__row_index__")}px">`,
            ...columns.map((column) => `<col style="width:${getAnalysisColumnWidth(column.name)}px">`),
        ];
        elements.tableColgroup.innerHTML = widths.join("");

        const totalWidth = getAnalysisColumnWidth("__row_index__") + columns.reduce((sum, column) => sum + getAnalysisColumnWidth(column.name), 0);
        if (elements.tableElement) {
            elements.tableElement.style.width = `${totalWidth}px`;
            elements.tableElement.style.minWidth = `${totalWidth}px`;
        }
    }

    function renderCharts() {
        if (!state.draft.charts.length) {
            elements.chartsList.innerHTML = '<div class="analysis-empty-list">В этом типе анализа пока нет графиков.</div>';
            elements.compatibilityNote.textContent = state.dataset.columns.length
                ? "Текущий источник данных готов к привязке карточек графиков."
                : "Пока источник данных не подключен, поля карточек графиков будут пустыми.";
            elements.addChartButton.disabled = false;
            return;
        }

        const incompatibleCount = state.draft.charts.filter((chart) => !getChartCompatibility(chart).ok).length;
        elements.compatibilityNote.textContent = incompatibleCount
            ? `Найдены несовместимые карточки графиков: ${incompatibleCount}. Переназначьте источник или поля.`
            : "Все карточки графиков совместимы с текущим источником данных.";
        elements.addChartButton.disabled = state.draft.charts.length >= 4;
        elements.chartsList.innerHTML = state.draft.charts
            .map((chart, index) => {
                const chartNumber = index + 1;
                const compatibility = getChartCompatibility(chart);
                const preview = state.chartPreviews[index] || null;
                const hasComment = Boolean(String(chart.comment_title || "").trim() || String(chart.comment_text || "").trim());
                const paramsCollapsed = Boolean(chart.ui_collapsed);
                const annotations = chart.annotations || [];
                const pieLabelOffsets = normalizePieSectorLabelOffsets(chart.pie_sector_label_offsets);
                return `
                    <article class="chart-card ${chart.is_hidden ? "is-hidden" : ""} ${compatibility.ok ? "" : "is-incompatible"}" id="chart_card_${chartNumber}" data-chart-index="${index}">
                        <div class="chart-card-head">
                            <div>
                                <div class="chart-card-title">График ${chartNumber}</div>
                                <div class="chart-card-status ${compatibility.ok ? "" : "warning"}">${compatibility.ok ? "Совместим" : "Требует настройки"}</div>
                                <div class="chart-card-comment-state ${hasComment ? "filled" : "empty"}">${hasComment ? "Комментарий заполнен" : "Комментарий не заполнен"}</div>
                            </div>
                            <div class="chart-card-actions">
                                <button type="button" class="secondary-button small-button" data-analysis-action="save-chart" data-chart-index="${index}">HTML</button>
                                <button type="button" class="secondary-button small-button" data-analysis-action="print-chart" data-chart-index="${index}">Печать</button>
                                <button type="button" class="secondary-button small-button" data-analysis-action="export-table" data-chart-index="${index}">XLSX</button>
                                <button type="button" class="secondary-button small-button" data-analysis-action="add-annotation" data-chart-index="${index}">Добавить надпись</button>
                                ${chart.chart_type === "pie" ? `<button type="button" class="secondary-button small-button" data-analysis-action="reset-pie-sector-labels" data-chart-index="${index}" ${pieLabelOffsets.length ? "" : "disabled"}>Сбросить подписи pie</button>` : ""}
                                <button type="button" class="secondary-button small-button" data-analysis-action="toggle-params" data-chart-index="${index}">${paramsCollapsed ? "Показать параметры" : "Свернуть параметры"}</button>
                                <button type="button" class="secondary-button small-button" data-analysis-action="move-up" data-chart-index="${index}" ${index === 0 ? "disabled" : ""}>↑</button>
                                <button type="button" class="secondary-button small-button" data-analysis-action="move-down" data-chart-index="${index}" ${index === state.draft.charts.length - 1 ? "disabled" : ""}>↓</button>
                                <button type="button" class="secondary-button small-button" id="btn_hide_chart_${chartNumber}" data-analysis-action="toggle-hidden" data-chart-index="${index}">${chart.is_hidden ? "Показать график" : "Скрыть график"}</button>
                                <button type="button" class="danger-button small-button" id="btn_delete_chart_${chartNumber}" data-analysis-action="delete-chart" data-chart-index="${index}">Удалить</button>
                            </div>
                        </div>
                        <div class="chart-card-summary">${compatibility.ok ? `Источник: ${escapeHtml(getSourceKindLabel(chart.source_kind || "none"))}.` : escapeHtml(compatibility.issues.join("; "))}</div>
                        <div class="chart-card-summary chart-card-params-state">${paramsCollapsed ? "Параметры графика свернуты." : "Параметры графика развернуты."}</div>
                        <div class="chart-card-grid ${paramsCollapsed ? "is-collapsed" : ""}">
                            <label>
                                <span>Тип графика</span>
                                <select id="select_chart_type_${chartNumber}" data-analysis-field="chart_type" data-chart-index="${index}">
                                    <option value="bar" ${chart.chart_type === "bar" ? "selected" : ""}>Столбиковая диаграмма</option>
                                    <option value="line" ${chart.chart_type === "line" ? "selected" : ""}>Линейный график</option>
                                    <option value="pie" ${chart.chart_type === "pie" ? "selected" : ""}>Круговая диаграмма</option>
                                    <option value="table" ${chart.chart_type === "table" ? "selected" : ""}>Таблица-агрегация</option>
                                </select>
                            </label>
                            <label>
                                <span>Источник</span>
                                <select id="select_source_${chartNumber}" data-analysis-field="source_kind" data-chart-index="${index}">
                                    <option value="none" ${chart.source_kind === "none" ? "selected" : ""}>Не выбран</option>
                                    <option value="facts" ${chart.source_kind === "facts" ? "selected" : ""}>Загрузка фактов</option>
                                    <option value="file" ${chart.source_kind === "file" ? "selected" : ""}>Файл анализа</option>
                                </select>
                            </label>
                            <label>
                                <span>X</span>
                                <select id="select_x_${chartNumber}" data-analysis-field="x_field" data-chart-index="${index}">${getColumnOptionsMarkup(chart.x_field)}</select>
                            </label>
                            <label>
                                <span>Y</span>
                                <select id="select_y_${chartNumber}" data-analysis-field="y_field" data-chart-index="${index}">${getColumnOptionsMarkup(chart.y_field)}</select>
                            </label>
                            <label>
                                <span>Group</span>
                                <select id="select_group_${chartNumber}" data-analysis-field="group_field" data-chart-index="${index}">${getColumnOptionsMarkup(chart.group_field)}</select>
                            </label>
                            <label>
                                <span>Agg</span>
                                <select id="select_agg_${chartNumber}" data-analysis-field="agg_func" data-chart-index="${index}">
                                    <option value="count" ${chart.agg_func === "count" ? "selected" : ""}>count</option>
                                    <option value="sum" ${chart.agg_func === "sum" ? "selected" : ""}>sum</option>
                                    <option value="avg" ${chart.agg_func === "avg" ? "selected" : ""}>avg</option>
                                    <option value="min" ${chart.agg_func === "min" ? "selected" : ""}>min</option>
                                    <option value="max" ${chart.agg_func === "max" ? "selected" : ""}>max</option>
                                </select>
                            </label>
                            <label>
                                <span>Цвет</span>
                                <input id="input_color_${chartNumber}" type="text" value="${escapeHtml(chart.color)}" data-analysis-field="color" data-chart-index="${index}">
                            </label>
                            <label class="chart-card-wide">
                                <span>Легенда</span>
                                <input id="input_legend_${chartNumber}" type="text" value="${escapeHtml(chart.legend)}" data-analysis-field="legend" data-chart-index="${index}">
                            </label>
                            <label>
                                <span>Положение легенды</span>
                                <select id="select_legend_position_${chartNumber}" data-analysis-field="legend_position" data-chart-index="${index}">
                                    <option value="right" ${chart.legend_position === "right" ? "selected" : ""}>справа</option>
                                    <option value="left" ${chart.legend_position === "left" ? "selected" : ""}>слева</option>
                                    <option value="bottom" ${chart.legend_position === "bottom" ? "selected" : ""}>снизу</option>
                                    <option value="inside-top-right" ${chart.legend_position === "inside-top-right" ? "selected" : ""}>внутри справа сверху</option>
                                    <option value="inside-top-left" ${chart.legend_position === "inside-top-left" ? "selected" : ""}>внутри слева сверху</option>
                                    <option value="hidden" ${chart.legend_position === "hidden" ? "selected" : ""}>скрыта</option>
                                </select>
                            </label>
                            <label>
                                <span>Подписи секторов pie</span>
                                <select id="select_pie_label_mode_${chartNumber}" data-analysis-field="pie_label_mode" data-chart-index="${index}">
                                    <option value="legend" ${chart.pie_label_mode === "legend" ? "selected" : ""}>без подписей у секторов</option>
                                    <option value="outside" ${chart.pie_label_mode === "outside" ? "selected" : ""}>рядом с секторами</option>
                                    <option value="outside-with-lines" ${chart.pie_label_mode === "outside-with-lines" ? "selected" : ""}>рядом с секторами + линии</option>
                                </select>
                            </label>
                            <label>
                                <span>Значения pie</span>
                                <select id="select_pie_value_mode_${chartNumber}" data-analysis-field="pie_value_mode" data-chart-index="${index}">
                                    <option value="value" ${chart.pie_value_mode === "value" ? "selected" : ""}>значение</option>
                                    <option value="percent" ${chart.pie_value_mode === "percent" ? "selected" : ""}>доля, %</option>
                                    <option value="value-percent" ${chart.pie_value_mode === "value-percent" ? "selected" : ""}>значение + доля, %</option>
                                </select>
                            </label>
                            <label>
                                <span>Центр pie</span>
                                <select id="select_pie_center_mode_${chartNumber}" data-analysis-field="pie_center_mode" data-chart-index="${index}">
                                    <option value="total" ${chart.pie_center_mode === "total" ? "selected" : ""}>сумма</option>
                                    <option value="hidden" ${chart.pie_center_mode === "hidden" ? "selected" : ""}>скрыть</option>
                                </select>
                            </label>
                            <label>
                                <span>Подпись сектора pie</span>
                                <select id="select_pie_sector_label_mode_${chartNumber}" data-analysis-field="pie_sector_label_mode" data-chart-index="${index}">
                                    <option value="name" ${chart.pie_sector_label_mode === "name" ? "selected" : ""}>только название</option>
                                    <option value="value" ${chart.pie_sector_label_mode === "value" ? "selected" : ""}>только значение</option>
                                    <option value="percent" ${chart.pie_sector_label_mode === "percent" ? "selected" : ""}>только доля, %</option>
                                    <option value="value-percent" ${chart.pie_sector_label_mode === "value-percent" ? "selected" : ""}>значение + доля, %</option>
                                    <option value="name-value" ${chart.pie_sector_label_mode === "name-value" ? "selected" : ""}>название + значение</option>
                                    <option value="name-percent" ${chart.pie_sector_label_mode === "name-percent" ? "selected" : ""}>название + доля, %</option>
                                    <option value="name-value-percent" ${chart.pie_sector_label_mode === "name-value-percent" ? "selected" : ""}>название + значение + доля, %</option>
                                </select>
                            </label>
                            <label class="chart-card-wide">
                                <span>Подписи</span>
                                <input id="input_labels_${chartNumber}" type="text" value="${escapeHtml(chart.labels)}" data-analysis-field="labels" data-chart-index="${index}">
                            </label>
                            <label class="chart-card-wide">
                                <span>Заголовок комментария</span>
                                <input id="input_comment_title_${chartNumber}" type="text" value="${escapeHtml(chart.comment_title || "")}" data-analysis-field="comment_title" data-chart-index="${index}">
                            </label>
                            <label class="chart-card-wide">
                                <span>Комментарий</span>
                                <textarea id="input_comment_text_${chartNumber}" rows="4" data-analysis-field="comment_text" data-chart-index="${index}">${escapeHtml(chart.comment_text || "")}</textarea>
                            </label>
                            <div class="chart-card-wide annotation-editor">
                                <div class="annotation-editor-head">
                                    <span>Надписи на графике</span>
                                    <strong>${annotations.length}</strong>
                                </div>
                                <div class="annotation-editor-list">
                                    ${annotations.length
                                        ? annotations
                                            .map(
                                                (annotation) => `
                                                    <div class="annotation-editor-row">
                                                        <input
                                                            type="text"
                                                            value="${escapeHtml(annotation.text)}"
                                                            data-analysis-annotation-text="true"
                                                            data-chart-index="${index}"
                                                            data-annotation-id="${escapeHtml(annotation.id)}"
                                                            placeholder="Текст надписи"
                                                        >
                                                        <button type="button" class="secondary-button small-button" data-analysis-action="remove-annotation" data-chart-index="${index}" data-annotation-id="${escapeHtml(annotation.id)}">Удалить</button>
                                                    </div>
                                                `
                                            )
                                            .join("")
                                        : '<div class="annotation-editor-empty">Добавьте надпись, затем перетащите ее мышью по графику.</div>'}
                                </div>
                            </div>
                            ${chart.chart_type === "pie" ? `
                                <div class="chart-card-wide annotation-editor pie-label-editor">
                                    <div class="annotation-editor-head">
                                        <span>Подписи секторов pie</span>
                                        <strong>${pieLabelOffsets.length}</strong>
                                    </div>
                                    <div class="annotation-editor-empty">Подписи секторов можно перетаскивать прямо по графику. Они сохраняются и остаются внутри рамки SVG.</div>
                                </div>
                            ` : ""}
                        </div>
                        ${buildChartPreviewMarkup(preview, chart)}
                        ${hasComment ? `<div class="chart-card-comment-preview">${buildChartCommentMarkup(chart)}</div>` : ""}
                    </article>
                `;
            })
            .join("");
    }

    function getPreparedTableRows() {
        const searchNeedle = state.table.search.trim().toLowerCase();
        const columns = state.dataset.columns || [];
        const rows = (state.dataset.rows || []).slice();

        let filteredRows = rows;
        if (searchNeedle) {
            filteredRows = rows.filter((row) =>
                columns.some((column) => String((row.values || {})[column.name] ?? "").toLowerCase().includes(searchNeedle))
            );
        }

        if (state.table.sortColumn) {
            const columnName = state.table.sortColumn;
            const direction = state.table.sortDirection === "desc" ? -1 : 1;
            filteredRows.sort((left, right) => {
                const leftValue = String((left.values || {})[columnName] ?? "");
                const rightValue = String((right.values || {})[columnName] ?? "");
                return leftValue.localeCompare(rightValue, "ru", { sensitivity: "base", numeric: true }) * direction;
            });
        }

        return filteredRows;
    }

    function renderTable() {
        const dataset = state.dataset;
        const columns = dataset.columns || [];
        const rows = getPreparedTableRows();
        const visibleRows = rows.slice(0, MAX_ANALYSIS_ROWS);

        elements.loadingIndicator.classList.toggle("hidden", !state.loading);

        if (state.loading) {
            elements.emptyState.classList.add("hidden");
            elements.tableWrap.classList.add("hidden");
            return;
        }

        if (!columns.length) {
            elements.emptyState.classList.remove("hidden");
            elements.tableWrap.classList.add("hidden");
            elements.emptyState.innerHTML = `
                <p>${state.userState.visual_source_kind === "facts"
                    ? "Источник выбран, но активных данных экспортируемого слоя сейчас нет. Повторите подключение данных из вкладки «Загрузка фактов»."
                    : state.userState.visual_source_kind === "file"
                        ? "Файл анализа пока не загружен или источник пуст. Загрузите файл повторно."
                        : "Источник данных не выбран. Подключите данные из вкладки «Загрузка фактов» или загрузите файл анализа."}</p>
                <button type="button" id="btn_retry_use_facts" class="secondary-button">Повторить подключение данных</button>
            `;
            document.getElementById("btn_retry_use_facts")?.addEventListener("click", () => {
                if (state.userState.visual_source_kind === "file") {
                    elements.analysisFileInput?.click();
                    return;
                }
                void useFactsDataset();
            });
            elements.countLabel.textContent = "Строк: 0 | Столбцов: 0";
            return;
        }

        if (!rows.length) {
            elements.emptyState.classList.remove("hidden");
            elements.tableWrap.classList.add("hidden");
            elements.emptyState.innerHTML = "<p>После поиска в таблице анализа не осталось строк. Измените строку поиска или сбросьте сортировку.</p>";
            elements.countLabel.textContent = `Строк: 0 | Столбцов: ${columns.length}`;
            return;
        }

        elements.emptyState.classList.add("hidden");
        elements.tableWrap.classList.remove("hidden");
        elements.countLabel.textContent = `Строк: ${rows.length}${rows.length > visibleRows.length ? ` (показано ${visibleRows.length})` : ""} | Столбцов: ${columns.length}`;

        renderTableColgroup(columns);

        elements.tableHead.innerHTML = `
            <tr>
                <th>#</th>
                ${columns
                    .map(
                        (column) => `
                            <th>
                                <div class="analysis-th-shell">
                                    <button type="button" data-analysis-action="sort-column" data-column="${escapeHtml(column.name)}">
                                        ${escapeHtml(column.name)}${state.table.sortColumn === column.name ? ` ${state.table.sortDirection === "asc" ? "▲" : "▼"}` : ""}
                                    </button>
                                    <span class="analysis-col-resizer" data-analysis-action="resize-column" data-column="${escapeHtml(column.name)}"></span>
                                </div>
                            </th>
                        `
                    )
                    .join("")}
            </tr>
        `;

        elements.tableBody.innerHTML = visibleRows
            .map(
                (row, index) => `
                    <tr>
                        <td>${row.row_number || index + 1}</td>
                        ${columns.map((column) => `<td>${escapeHtml((row.values || {})[column.name] ?? "")}</td>`).join("")}
                    </tr>
                `
            )
            .join("");
    }

    function startTableColumnResize(columnName, event) {
        if (event.button !== 0) {
            return;
        }

        event.preventDefault();
        event.stopPropagation();
        activeTableResize = {
            columnName,
            startX: event.clientX,
            startWidth: getAnalysisColumnWidth(columnName),
        };
        window.addEventListener("mousemove", handleTableColumnResize);
        window.addEventListener("mouseup", stopTableColumnResize);
    }

    function handleTableColumnResize(event) {
        if (!activeTableResize) {
            return;
        }

        const nextWidth = Math.max(120, activeTableResize.startWidth + (event.clientX - activeTableResize.startX));
        state.table.columnWidths[activeTableResize.columnName] = nextWidth;
        renderTableColgroup(state.dataset.columns || []);
    }

    function stopTableColumnResize() {
        if (!activeTableResize) {
            return;
        }

        activeTableResize = null;
        persistAnalysisUiPreferences();
        state.userState.table_column_widths = { ...state.table.columnWidths };
        scheduleUserStateSave({ table_column_widths: state.userState.table_column_widths });
        window.removeEventListener("mousemove", handleTableColumnResize);
        window.removeEventListener("mouseup", stopTableColumnResize);
    }

    function scheduleChartPreviewRefresh() {
        window.clearTimeout(chartPreviewTimer);
        chartPreviewTimer = window.setTimeout(() => {
            void refreshChartPreviews();
        }, 120);
    }

    async function refreshChartPreviews() {
        const requestId = ++chartPreviewRequestId;
        const charts = state.draft.charts.slice();

        if (!charts.length) {
            state.chartPreviews = [];
            renderCharts();
            return;
        }

        state.chartPreviews = charts.map((chart) =>
            chart.is_hidden
                ? { state: "hidden", message: "График скрыт пользователем.", warnings: [] }
                : { state: "loading", message: "Подготовка графика...", warnings: [] }
        );
        renderCharts();

        const previews = await Promise.all(
            charts.map(async (chart) => {
                if (chart.is_hidden) {
                    return { state: "hidden", message: "График скрыт пользователем.", warnings: [] };
                }
                try {
                    const response = await apiJson("/api/analysis/chart-preview", {
                        method: "POST",
                        body: JSON.stringify({ chart }),
                    });
                    return response.preview || { state: "error", message: "Не удалось подготовить график.", warnings: [] };
                } catch (error) {
                    return {
                        state: "error",
                        message: error.error || "Не удалось построить график.",
                        warnings: [],
                    };
                }
            })
        );

        if (requestId !== chartPreviewRequestId) {
            return;
        }

        state.chartPreviews = previews;
        renderCharts();
    }

    function renderAll(options = {}) {
        renderViewMode();
        renderMessage();
        renderSourceMeta();
        renderTypeSelect();
        renderCharts();
        renderTable();
        renderSheetView();
        if (options.refreshPreviews !== false) {
            scheduleChartPreviewRefresh();
        }
    }

    async function loadBootstrap() {
        state.loading = true;
        renderAll({ refreshPreviews: false });
        try {
            const payload = await apiJson("/api/analysis/bootstrap");
            state.types = payload.types || [];
            state.userState = { ...defaultUserState, ...(payload.user_state || {}) };
            state.dataset = payload.dataset || emptyDataset();
            applyPersistedAnalysisUiState();
            applyBackendAnalysisWidth();
            syncDraftFromSelection();
            state.table.search = state.userState.table_search || "";
            state.table.sortColumn = state.userState.table_sort_column || "";
            state.table.sortDirection = state.userState.table_sort_direction || "asc";
            elements.searchInput.value = state.table.search;
            setMessage(
                state.dataset.columns.length
                    ? `Подключён источник анализа: ${state.dataset.source_label || "данные доступны"}.`
                    : "Выберите тип анализа, подключите данные из вкладки «Загрузка фактов» или загрузите файл анализа.",
                state.dataset.columns.length ? "success" : "info"
            );
        } catch (error) {
            setMessage(error.error || "Не удалось загрузить состояние вкладки Анализ.", "error");
        } finally {
            state.loading = false;
            renderAll();
        }
    }

    async function saveUserState(patch) {
        queueUserStateSavePatch(patch);
        return flushUserStateSave();
    }

    async function useFactsDataset() {
        state.loading = true;
        renderTable();
        try {
            const payload = await apiJson("/api/analysis/use-facts", {
                method: "POST",
                body: JSON.stringify({}),
            });
            state.dataset = payload.dataset || emptyDataset();
            state.userState.visual_source_kind = "facts";
            scheduleUserStateSave({ visual_source_kind: "facts" });
            setMessage(`Экспортируемый слой из вкладки «Загрузка фактов» подключён: ${state.dataset.source_label || "без имени файла"}.`, "success");
        } catch (error) {
            setMessage(error.error || "Не удалось подключить данные из вкладки «Загрузка фактов».", "error");
        } finally {
            state.loading = false;
            renderAll();
        }
    }

    async function uploadAnalysisFile(file) {
        if (!file) {
            return;
        }

        state.loading = true;
        renderTable();
        const formData = new FormData();
        formData.append("file", file);

        try {
            const payload = await apiJson("/api/analysis/upload", {
                method: "POST",
                body: formData,
            });
            state.dataset = payload.dataset || emptyDataset();
            state.userState.visual_source_kind = "file";
            scheduleUserStateSave({ visual_source_kind: "file" });
            setMessage(`Файл анализа загружен: ${state.dataset.source_file_name || state.dataset.source_label || "без имени"}.`, "success");
        } catch (error) {
            setMessage(error.error || "Не удалось загрузить файл анализа.", "error");
        } finally {
            state.loading = false;
            if (elements.analysisFileInput) {
                elements.analysisFileInput.value = "";
            }
            renderAll();
        }
    }

    async function createAnalysisType() {
        const name = window.prompt("Введите имя нового типа анализа:", "Новый тип анализа");
        if (!name) {
            return;
        }

        try {
            const response = await apiJson("/api/analysis/types", {
                method: "POST",
                body: JSON.stringify({ name, charts: [] }),
            });
            state.userState.selected_analysis_type_id = response.analysis_type.id;
            state.userState.draft_state = {
                analysis_type_id: response.analysis_type.id,
                name: response.analysis_type.name,
                charts: [],
            };
            await saveUserState({
                selected_analysis_type_id: response.analysis_type.id,
                draft_state: state.userState.draft_state,
            });
            await loadBootstrap();
            setMessage(`Тип анализа «${response.analysis_type.name}» создан.`, "success");
        } catch (error) {
            setMessage((error.errors || [error.error || "Не удалось создать тип анализа."]).join(" "), "error");
        }
    }

    async function deleteSelectedAnalysisType() {
        const current = getSelectedType();
        if (!current) {
            setMessage("Сначала выберите тип анализа для удаления.", "warning");
            return;
        }

        if (!window.confirm(`Удалить тип анализа «${current.name}»?`)) {
            return;
        }

        try {
            await apiJson(`/api/analysis/types/${current.id}`, { method: "DELETE" });
            state.userState.selected_analysis_type_id = null;
            state.userState.draft_state = null;
            await saveUserState({ selected_analysis_type_id: null, draft_state: null });
            await loadBootstrap();
            setMessage(`Тип анализа «${current.name}» удалён.`, "success");
        } catch (error) {
            setMessage(error.error || "Не удалось удалить тип анализа.", "error");
        }
    }

    async function saveSelectedAnalysisType() {
        if (!state.draft.selectedTypeId) {
            setMessage("Сначала создайте или выберите тип анализа.", "warning");
            return;
        }

        try {
            await apiJson(`/api/analysis/types/${state.draft.selectedTypeId}`, {
                method: "PUT",
                body: JSON.stringify({
                    name: state.draft.name,
                    charts: state.draft.charts.map((chart, index) => toPersistedChart(chart, index)),
                }),
            });
            state.userState.draft_state = buildDraftStatePayload();
            await saveUserState({ draft_state: state.userState.draft_state });
            await loadBootstrap();
            setMessage("Тип анализа сохранён.", "success");
        } catch (error) {
            setMessage((error.errors || [error.error || "Не удалось сохранить тип анализа."]).join(" "), "error");
        }
    }

    function resetDraftSettings() {
        if (!state.draft.selectedTypeId) {
            state.table.search = "";
            state.table.sortColumn = "";
            state.table.sortDirection = "asc";
            elements.searchInput.value = "";
            scheduleUserStateSave({ table_search: "", table_sort_column: "", table_sort_direction: "asc", draft_state: null });
            renderTable();
            return;
        }

        state.draft.charts = [];
        state.table.search = "";
        state.table.sortColumn = "";
        state.table.sortDirection = "asc";
        elements.searchInput.value = "";
        state.userState.draft_state = buildDraftStatePayload();
        scheduleUserStateSave({
            table_search: "",
            table_sort_column: "",
            table_sort_direction: "asc",
            draft_state: state.userState.draft_state,
        });
        setMessage("Настройки текущего типа анализа сброшены локально. Сохраните тип анализа, если хотите записать изменения в БД.", "warning");
        renderAll();
    }

    function addChart() {
        if (state.draft.charts.length >= 4) {
            setMessage("Во вкладке «Анализ» допускается не более 4 графиков.", "warning");
            return;
        }

        state.draft.charts.push(defaultChart(state.draft.charts.length, state.dataset.source_kind !== "none" ? state.dataset.source_kind : "none"));
        state.userState.draft_state = buildDraftStatePayload();
        scheduleUserStateSave({ draft_state: state.userState.draft_state });
        setMessage("Карточка графика добавлена. Сохраните тип анализа, чтобы записать изменения в БД.", "success");
        renderAll();
    }

    function updateChartField(index, field, value) {
        const chart = state.draft.charts[index];
        if (!chart) {
            return;
        }

        chart[field] = value;
        state.userState.draft_state = buildDraftStatePayload();
    }

    function reorderCharts(fromIndex, toIndex) {
        if (toIndex < 0 || toIndex >= state.draft.charts.length) {
            return;
        }
        const [chart] = state.draft.charts.splice(fromIndex, 1);
        state.draft.charts.splice(toIndex, 0, chart);
        state.draft.charts = state.draft.charts.map((item, index) => ({ ...item, position: index }));
        state.userState.draft_state = buildDraftStatePayload();
        scheduleUserStateSave({ draft_state: state.userState.draft_state });
        renderAll();
    }

    function bindEvents() {
        elements.uploadFileButton.addEventListener("click", () => {
            elements.analysisFileInput?.click();
        });
        elements.analysisFileInput?.addEventListener("change", (event) => {
            void uploadAnalysisFile(event.target.files?.[0]);
        });
        elements.useFactsButton.addEventListener("click", () => {
            void useFactsDataset();
        });
        elements.retryUseFactsButton?.addEventListener("click", () => {
            void useFactsDataset();
        });
        elements.createTypeButton.addEventListener("click", () => {
            void createAnalysisType();
        });
        elements.deleteTypeButton.addEventListener("click", () => {
            void deleteSelectedAnalysisType();
        });
        elements.saveTypeButton.addEventListener("click", () => {
            void saveSelectedAnalysisType();
        });
        elements.resetButton.addEventListener("click", resetDraftSettings);
        elements.exportReportButton?.addEventListener("click", exportVisibleChartsReport);
        elements.addChartButton.addEventListener("click", addChart);
        elements.editorViewButton?.addEventListener("click", () => {
            state.viewMode = "editor";
            state.userState.view_mode = state.viewMode;
            persistAnalysisUiPreferences();
            void saveUserState({ view_mode: state.viewMode }).catch(handleUserStateSaveError);
            renderViewMode();
        });
        elements.sheetViewButton?.addEventListener("click", () => {
            state.viewMode = "sheet";
            state.userState.view_mode = state.viewMode;
            persistAnalysisUiPreferences();
            void saveUserState({ view_mode: state.viewMode }).catch(handleUserStateSaveError);
            renderViewMode();
            renderSheetView();
        });
        elements.pairViewButton?.addEventListener("click", () => {
            state.viewMode = "pair";
            state.userState.view_mode = state.viewMode;
            persistAnalysisUiPreferences();
            void saveUserState({ view_mode: state.viewMode }).catch(handleUserStateSaveError);
            renderViewMode();
            renderSheetView();
        });
        elements.quadViewButton?.addEventListener("click", () => {
            state.viewMode = "quad";
            state.userState.view_mode = state.viewMode;
            persistAnalysisUiPreferences();
            void saveUserState({ view_mode: state.viewMode }).catch(handleUserStateSaveError);
            renderViewMode();
            renderSheetView();
        });
        elements.typeSelect.addEventListener("change", async (event) => {
            const nextId = event.target.value ? Number(event.target.value) : null;
            state.userState.selected_analysis_type_id = nextId;
            await saveUserState({ selected_analysis_type_id: nextId });
            syncDraftFromSelection();
            setMessage(nextId ? "Тип анализа переключён." : "Тип анализа не выбран.", "info");
            renderAll();
        });
        elements.searchInput.addEventListener("input", (event) => {
            state.table.search = event.target.value || "";
            scheduleUserStateSave({ table_search: state.table.search });
            renderTable();
        });
        window.addEventListener("resize", scheduleSheetViewportLayout);
        window.addEventListener("orientationchange", scheduleSheetViewportLayout);
        elements.sortResetButton.addEventListener("click", () => {
            state.table.sortColumn = "";
            state.table.sortDirection = "asc";
            scheduleUserStateSave({ table_sort_column: "", table_sort_direction: "asc" });
            renderTable();
            setMessage("Сортировка таблицы анализа сброшена.", "info");
        });
        elements.tableHead.addEventListener("click", (event) => {
            if (event.target.closest("[data-analysis-action='resize-column']")) {
                return;
            }
            const actionNode = event.target.closest("[data-analysis-action='sort-column']");
            if (!actionNode) {
                return;
            }
            const column = actionNode.dataset.column;
            if (state.table.sortColumn === column) {
                state.table.sortDirection = state.table.sortDirection === "asc" ? "desc" : "asc";
            } else {
                state.table.sortColumn = column;
                state.table.sortDirection = "asc";
            }
            scheduleUserStateSave({ table_sort_column: state.table.sortColumn, table_sort_direction: state.table.sortDirection });
            renderTable();
        });
        elements.tableHead.addEventListener("mousedown", (event) => {
            const resizer = event.target.closest("[data-analysis-action='resize-column']");
            if (!resizer) {
                return;
            }
            startTableColumnResize(resizer.dataset.column, event);
        });
        elements.chartsList.addEventListener("click", (event) => {
            const actionNode = event.target.closest("[data-analysis-action]");
            if (!actionNode) {
                return;
            }

            const action = actionNode.dataset.analysisAction;
            const index = Number(actionNode.dataset.chartIndex);
            if (Number.isNaN(index)) {
                return;
            }

            if (action === "delete-chart") {
                state.draft.charts.splice(index, 1);
                state.draft.charts = state.draft.charts.map((item, position) => ({ ...item, position }));
                state.userState.draft_state = buildDraftStatePayload();
                scheduleUserStateSave({ draft_state: state.userState.draft_state });
                renderAll();
                setMessage("Карточка графика удалена. Сохраните тип анализа, чтобы записать изменения в БД.", "success");
                return;
            }
            if (action === "save-chart") {
                saveChartPreview(index);
                return;
            }
            if (action === "add-annotation") {
                addAnnotation(index);
                return;
            }
            if (action === "reset-pie-sector-labels") {
                resetPieSectorLabelOffsets(index);
                return;
            }
            if (action === "print-chart") {
                printChartPreview(index);
                return;
            }
            if (action === "export-table") {
                void exportChartTable(index);
                return;
            }
            if (action === "remove-annotation") {
                removeAnnotation(index, actionNode.dataset.annotationId);
                return;
            }
            if (action === "toggle-params") {
                const chart = state.draft.charts[index];
                if (!chart) {
                    return;
                }
                chart.ui_collapsed = !chart.ui_collapsed;
                renderCharts();
                return;
            }
            if (action === "toggle-hidden") {
                state.draft.charts[index].is_hidden = !state.draft.charts[index].is_hidden;
                state.userState.draft_state = buildDraftStatePayload();
                scheduleUserStateSave({ draft_state: state.userState.draft_state });
                renderAll();
                return;
            }
            if (action === "move-up") {
                reorderCharts(index, index - 1);
                return;
            }
            if (action === "move-down") {
                reorderCharts(index, index + 1);
                return;
            }
        });
        elements.chartsList.addEventListener("change", (event) => {
            const target = event.target;
            const field = target.dataset.analysisField;
            const index = Number(target.dataset.chartIndex);
            if (!field || Number.isNaN(index)) {
                return;
            }
            updateChartField(index, field, target.type === "checkbox" ? target.checked : target.value);
            scheduleUserStateSave({ draft_state: state.userState.draft_state });
            renderAll();
        });
        elements.chartsList.addEventListener("input", (event) => {
            const target = event.target;
            if (target.dataset.analysisAnnotationText) {
                updateAnnotationText(Number(target.dataset.chartIndex), target.dataset.annotationId, target.value);
                syncAnnotationNodes(Number(target.dataset.chartIndex), target.dataset.annotationId, null, null, target.value);
                scheduleUserStateSave({ draft_state: state.userState.draft_state });
                return;
            }
            const field = target.dataset.analysisField;
            const index = Number(target.dataset.chartIndex);
            if (!field || Number.isNaN(index)) {
                return;
            }
            updateChartField(index, field, target.value);
            scheduleUserStateSave({ draft_state: state.userState.draft_state });
        });
        elements.shell.addEventListener("mousedown", (event) => {
            const pieLabelHandle = event.target.closest("[data-analysis-action='drag-pie-sector-label']");
            if (pieLabelHandle) {
                startPieSectorLabelDrag(Number(pieLabelHandle.dataset.chartIndex), Number(pieLabelHandle.dataset.sectorIndex), event);
                return;
            }
            const handle = event.target.closest("[data-analysis-action='drag-annotation']");
            if (!handle) {
                return;
            }

            startAnnotationDrag(Number(handle.dataset.chartIndex), handle.dataset.annotationId, event);
        });
        window.addEventListener("analysis:left-width-changed", (event) => {
            const width = Number(event.detail?.width || 260);
            window.clearTimeout(widthSaveTimer);
            widthSaveTimer = window.setTimeout(() => {
                void saveUserState({ left_panel_width: width }).catch(handleUserStateSaveError);
            }, 250);
        });
        document.addEventListener("visibilitychange", () => {
            if (document.visibilityState !== "hidden") {
                return;
            }
            void flushUserStateSave({ keepalive: true }).catch(() => {});
        });
    }

    loadAnalysisUiPreferences();
    bindEvents();
    void loadBootstrap();
})();