/**
 * static/js/excel_mapper.js
 * Interaktivní vizuální mapovač datových polí do Excelu (.xlsx)
 * s podporou kontextového menu, figurek (badges), prioritních indexů a dotykového ovládání.
 */

// Globální stav mapovače
const ExcelMapper = {
    schema: null,
    settings: {},
    availableFiles: [],
    currentFile: '',
    currentSheet: '',
    sheetData: null,
    sheetNames: [],
    selectedCell: null,
    activeContextMenuCell: null,
    touchTimer: null,
    isDirty: false,

    async init() {
        console.log('🚀 Inicializace ExcelMapper...');
        try {
            await this.loadSchema();
            await this.loadSettings();
            await this.loadFiles();
            this.setupGlobalEvents();
            this.renderToolbar();
            if (this.availableFiles.length > 0) {
                // Výchozí soubor
                const defaultFile = this.availableFiles.includes('Hodiny_Cap.xlsx') 
                    ? 'Hodiny_Cap.xlsx' 
                    : this.availableFiles[0];
                await this.selectFile(defaultFile);
            }
        } catch (err) {
            console.error('Chyba při inicializaci ExcelMapper:', err);
            this.showToast('Chyba při načítání mapovače: ' + err.message, 'error');
        }
    },

    async loadSchema() {
        const res = await fetch('/api/mapping-schema');
        if (!res.ok) throw new Error('Nepodařilo se načíst schéma polí');
        this.schema = await res.json();
    },

    async loadSettings() {
        const res = await fetch('/api/settings');
        if (!res.ok) throw new Error('Nepodařilo se načíst nastavení');
        this.settings = await res.json() || {};
        this.ensureSettingsStructure();
    },

    async loadFiles() {
        const res = await fetch('/api/files');
        if (!res.ok) throw new Error('Nepodařilo se načíst seznam souborů');
        const data = await res.json();
        this.availableFiles = data.files || [];
    },

    ensureSettingsStructure() {
        if (!this.schema) return;
        for (const [catKey, catVal] of Object.entries(this.schema)) {
            if (!this.settings[catKey]) {
                this.settings[catKey] = {};
            }
            for (const fieldKey of Object.keys(catVal.fields)) {
                if (!Array.isArray(this.settings[catKey][fieldKey])) {
                    this.settings[catKey][fieldKey] = [];
                }
            }
        }
    },

    setupGlobalEvents() {
        // Zavření kontextového menu při kliknutí mimo
        document.addEventListener('click', (e) => {
            const menu = document.getElementById('mapper-context-menu');
            if (menu && !menu.contains(e.target) && !e.target.closest('.grid-cell')) {
                this.hideContextMenu();
            }
        });

        // Klávesa Escape pro zavření menu
        document.addEventListener('keydown', (e) => {
            if (e.key === 'Escape') {
                this.hideContextMenu();
            }
        });
    },

    renderToolbar() {
        const container = document.getElementById('excel-mapper-root');
        if (!container) return;

        container.innerHTML = `
            <div class="mapper-container">
                <!-- Ovládací panel -->
                <div class="mapper-toolbar">
                    <div class="mapper-toolbar-group">
                        <label for="mapper-file-select" class="form-label-modern mb-0" style="font-weight: 600;">
                            📂 Cílový Excel soubor:
                        </label>
                        <select id="mapper-file-select" class="form-control-modern py-1 px-3" style="width: auto; min-width: 220px;" onchange="ExcelMapper.onFileChange(this.value)">
                            ${this.availableFiles.map(f => `<option value="${f}" ${f === this.currentFile ? 'selected' : ''}>${f}</option>`).join('')}
                        </select>
                    </div>

                    <div class="mapper-toolbar-group">
                        <button type="button" class="btn-modern btn-secondary btn-sm" onclick="ExcelMapper.resetToDefaults()">
                            🔄 Obnovit výchozí šablonu
                        </button>
                        <button type="button" class="btn-modern btn-danger btn-sm" onclick="ExcelMapper.clearCurrentSheetMappings()">
                            🧹 Vyčistit list
                        </button>
                        <button type="button" id="mapper-save-btn" class="btn-modern btn-primary py-2 px-4" onclick="ExcelMapper.saveSettings()">
                            💾 Uložit Excel mapování
                        </button>
                    </div>
                </div>

                <!-- Nápověda -->
                <div class="glass-card py-2 px-3 mb-0" style="background: rgba(var(--text-accent-rgb), 0.06); border-left: 4px solid var(--text-accent);">
                    <small>
                        💡 <strong>Jak mapovat:</strong> Klikněte <em>pravým tlačítkem</em> (nebo <em>dlouhým stiskem</em> na mobilu / kliknutím) na buňku pro otevření nabídky polí. Značka s tvarem a barvou označuje přiřazené pole. Vícenásobná přiřazení jsou automaticky číslována jako záložní (#1, #2...).
                    </small>
                </div>

                <!-- Tabový přepínač listů a Mřížka -->
                <div class="mapper-workbench">
                    <div id="mapper-sheet-tabs" class="mapper-sheet-tabs"></div>
                    <div id="mapper-grid-wrapper" class="mapper-grid-wrapper">
                        <div class="text-center py-5 text-muted">Načítám tabulku...</div>
                    </div>
                </div>

                <!-- Legenda a přehled -->
                <div class="mapper-legend-card">
                    <div class="d-flex justify-content-between align-items-center">
                        <h3 class="m-0" style="font-size: 1.1rem;">📊 Přehled a legenda polí</h3>
                        <span id="mapper-total-count" class="badge-modern badge-primary">0 namapovaných buněk</span>
                    </div>
                    <div id="mapper-legend-grid" class="mapper-legend-grid"></div>
                </div>
            </div>

            <!-- Plovoucí kontextové menu -->
            <div id="mapper-context-menu" class="mapper-context-menu"></div>
        `;
    },

    async onFileChange(filename) {
        if (!filename) return;
        await this.selectFile(filename);
    },

    async selectFile(filename) {
        this.currentFile = filename;
        const fileSelect = document.getElementById('mapper-file-select');
        if (fileSelect) fileSelect.value = filename;

        try {
            const res = await fetch(`/api/sheets/${encodeURIComponent(filename)}`);
            if (!res.ok) throw new Error('Nepodařilo se načíst listy');
            const data = await res.json();
            this.sheetNames = data.sheets || [];

            this.renderSheetTabs();

            if (this.sheetNames.length > 0) {
                // Zachovat list nebo zvolit první
                const targetSheet = this.sheetNames.includes(this.currentSheet) 
                    ? this.currentSheet 
                    : this.sheetNames[0];
                await this.selectSheet(targetSheet);
            }
        } catch (err) {
            console.error(err);
            this.showToast('Chyba při načítání souboru: ' + err.message, 'error');
        }
    },

    renderSheetTabs() {
        const tabsContainer = document.getElementById('mapper-sheet-tabs');
        if (!tabsContainer) return;

        tabsContainer.innerHTML = this.sheetNames.map(sheet => `
            <div class="sheet-tab-item ${sheet === this.currentSheet ? 'active' : ''}" onclick="ExcelMapper.selectSheet('${sheet}')">
                <span>📄</span>
                <span>${sheet}</span>
            </div>
        `).join('');
    },

    async selectSheet(sheetName) {
        this.currentSheet = sheetName;
        this.renderSheetTabs();

        const gridWrapper = document.getElementById('mapper-grid-wrapper');
        if (gridWrapper) {
            gridWrapper.innerHTML = '<div class="text-center py-5 text-muted">⏳ Načítám obsah listu...</div>';
        }

        try {
            const res = await fetch(`/api/sheet_content/${encodeURIComponent(this.currentFile)}/${encodeURIComponent(sheetName)}`);
            if (!res.ok) throw new Error('Nepodařilo se načíst obsah listu');
            this.sheetData = await res.json();
            this.renderGrid();
            this.renderLegend();
        } catch (err) {
            console.error(err);
            if (gridWrapper) {
                gridWrapper.innerHTML = `<div class="text-center py-5 text-danger">Nepodařilo se načíst data listu: ${err.message}</div>`;
            }
        }
    },

    /**
     * Vrací všechna mapování pro konkrétní buňku (např. 'A8') na aktuálním souboru a listu
     */
    getMappingsForCell(cellAddress) {
        const results = [];
        if (!this.settings || !this.schema) return results;

        for (const [catKey, catVal] of Object.entries(this.schema)) {
            const catSettings = this.settings[catKey] || {};
            for (const [fieldKey, fieldVal] of Object.entries(catVal.fields)) {
                const locations = catSettings[fieldKey] || [];
                locations.forEach((loc, idx) => {
                    if (loc.file === this.currentFile && loc.sheet === this.currentSheet && loc.cell === cellAddress) {
                        results.push({
                            categoryKey: catKey,
                            fieldKey: fieldKey,
                            categoryMeta: catVal,
                            fieldMeta: fieldVal,
                            locationIndex: idx,
                            priority: idx + 1,
                            isFallback: idx > 0
                        });
                    }
                });
            }
        }
        return results;
    },

    renderGrid() {
        const gridWrapper = document.getElementById('mapper-grid-wrapper');
        if (!gridWrapper || !this.sheetData) return;

        const { data, rows, cols } = this.sheetData;
        let html = '<table class="excel-mapper-table"><thead><tr><th class="corner-header"></th>';

        // Hlavičky sloupců (A, B, C, ...)
        for (let col = 1; col <= cols; col++) {
            const colLetter = this.getColumnLetter(col);
            html += `<th>${colLetter}</th>`;
        }
        html += '</tr></thead><tbody>';

        // Řádky tabulky
        for (let r = 0; r < rows; r++) {
            const rowNumber = r + 1;
            html += `<tr><th class="row-header">${rowNumber}</th>`;

            for (let c = 0; c < cols; c++) {
                const colLetter = this.getColumnLetter(c + 1);
                const cellAddress = `${colLetter}${rowNumber}`;
                const cellVal = (data[r] && data[r][c] !== undefined) ? data[r][c] : '';
                const mappings = this.getMappingsForCell(cellAddress);
                const hasMapping = mappings.length > 0;

                // Vykreslení figurek (badges)
                let badgesHtml = '';
                if (hasMapping) {
                    badgesHtml = '<div class="mapper-badges-container">';
                    mappings.forEach(m => {
                        const shapeClass = `shape-${m.categoryMeta.shape || 'circle'}`;
                        const badgeText = m.fieldMeta.badge || m.fieldMeta.label.substring(0, 4);
                        const indexBadge = m.priority > 1 ? `<span class="mapper-badge-index">#${m.priority}</span>` : '';
                        badgesHtml += `
                            <span class="mapper-figure-badge ${shapeClass}" style="background-color: ${m.fieldMeta.color};" title="${m.categoryMeta.label} > ${m.fieldMeta.label} (Priorita #${m.priority})">
                                ${m.categoryMeta.icon} ${badgeText}${indexBadge}
                            </span>
                        `;
                    });
                    badgesHtml += '</div>';
                }

                const safeText = this.escapeHtml(cellVal);
                const isSelected = this.selectedCell === cellAddress;

                html += `
                    <td class="grid-cell ${hasMapping ? 'cell-has-mapping' : ''} ${isSelected ? 'cell-selected' : ''}" 
                        data-cell="${cellAddress}" 
                        data-value="${safeText}"
                        oncontextmenu="ExcelMapper.handleCellContextMenu(event, '${cellAddress}')"
                        onclick="ExcelMapper.handleCellClick(event, '${cellAddress}')"
                        ontouchstart="ExcelMapper.handleTouchStart(event, '${cellAddress}')"
                        ontouchend="ExcelMapper.handleTouchEnd(event, '${cellAddress}')"
                        title="${cellAddress}${cellVal ? ': ' + safeText : ''}">
                        <div class="cell-content-box">
                            <span class="cell-value-text">${safeText}</span>
                            ${badgesHtml}
                        </div>
                    </td>
                `;
            }
            html += '</tr>';
        }
        html += '</tbody></table>';

        gridWrapper.innerHTML = html;
    },

    getColumnLetter(colIndex) {
        let letter = '';
        while (colIndex > 0) {
            const mod = (colIndex - 1) % 26;
            letter = String.fromCharCode(65 + mod) + letter;
            colIndex = Math.floor((colIndex - mod) / 26);
        }
        return letter;
    },

    escapeHtml(str) {
        if (!str) return '';
        return String(str)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#039;');
    },

    handleCellClick(event, cellAddress) {
        this.selectedCell = cellAddress;
        this.highlightSelectedCell(cellAddress);
        // Na desktopu levý klik může také otevřít menu pokud uživatel preferuje kliknutí
        // Zde otevřeme menu přímo pro maximální pohodlí
        this.openContextMenu(event.clientX, event.clientY, cellAddress);
    },

    handleCellContextMenu(event, cellAddress) {
        event.preventDefault();
        this.selectedCell = cellAddress;
        this.highlightSelectedCell(cellAddress);
        this.openContextMenu(event.clientX, event.clientY, cellAddress);
    },

    handleTouchStart(event, cellAddress) {
        this.touchTimer = setTimeout(() => {
            const touch = event.touches[0];
            this.selectedCell = cellAddress;
            this.highlightSelectedCell(cellAddress);
            this.openContextMenu(touch.clientX, touch.clientY, cellAddress);
        }, 500);
    },

    handleTouchEnd(event, cellAddress) {
        if (this.touchTimer) {
            clearTimeout(this.touchTimer);
            this.touchTimer = null;
        }
    },

    highlightSelectedCell(cellAddress) {
        document.querySelectorAll('.excel-mapper-table td.grid-cell').forEach(td => {
            if (td.dataset.cell === cellAddress) {
                td.classList.add('cell-selected');
            } else {
                td.classList.remove('cell-selected');
            }
        });
    },

    openContextMenu(clientX, clientY, cellAddress) {
        this.activeContextMenuCell = cellAddress;
        const menu = document.getElementById('mapper-context-menu');
        if (!menu || !this.schema) return;

        const cellMappings = this.getMappingsForCell(cellAddress);
        const mappedFieldKeys = cellMappings.map(m => `${m.categoryKey}.${m.fieldKey}`);

        let groupsHtml = '';
        for (const [catKey, catVal] of Object.entries(this.schema)) {
            let fieldsHtml = '';
            for (const [fieldKey, fieldVal] of Object.entries(catVal.fields)) {
                const isAssignedToThis = mappedFieldKeys.includes(`${catKey}.${fieldKey}`);
                const currentLocations = (this.settings[catKey] && this.settings[catKey][fieldKey]) || [];
                const totalAssigned = currentLocations.length;

                fieldsHtml += `
                    <button type="button" class="menu-item-btn ${isAssignedToThis ? 'active-mapping' : ''}" 
                            data-search-text="${catVal.label} ${fieldVal.label}"
                            onclick="ExcelMapper.toggleMapping('${catKey}', '${fieldKey}', '${cellAddress}')">
                        <div class="menu-item-left">
                            <span class="menu-item-color-dot" style="background-color: ${fieldVal.color};"></span>
                            <span>${fieldVal.label}</span>
                        </div>
                        <div>
                            ${isAssignedToThis ? '✅' : (totalAssigned > 0 ? `<span class="badge badge-secondary" style="font-size:0.7rem;">+${totalAssigned}</span>` : '')}
                        </div>
                    </button>
                `;
            }

            groupsHtml += `
                <div class="menu-category-group">
                    <div class="menu-category-title">
                        <span>${catVal.icon}</span>
                        <span>${catVal.label}</span>
                    </div>
                    <div class="menu-fields-list">
                        ${fieldsHtml}
                    </div>
                </div>
            `;
        }

        let removeButtonHtml = '';
        if (cellMappings.length > 0) {
            removeButtonHtml = `
                <div class="menu-actions-divider"></div>
                <button type="button" class="menu-remove-btn" onclick="ExcelMapper.removeCellMappings('${cellAddress}')">
                    🗑️ Odebrat všechna přiřazení z buňky
                </button>
            `;
        }

        menu.innerHTML = `
            <div class="menu-header">
                <div>
                    <strong>Přiřadit pole k buňce</strong>
                </div>
                <span class="menu-cell-badge">${cellAddress}</span>
            </div>
            <input type="text" class="menu-search-input" placeholder="🔍 Filtrovat pole..." onkeyup="ExcelMapper.filterContextMenu(this.value)">
            <div class="menu-body" style="max-height: 280px; overflow-y: auto;">
                ${groupsHtml}
            </div>
            ${removeButtonHtml}
        `;

        menu.style.display = 'block';

        // Pozicování menu v rámci viewportu
        const menuWidth = 310;
        const menuHeight = menu.offsetHeight || 350;
        let left = clientX;
        let top = clientY;

        if (left + menuWidth > window.innerWidth - 20) {
            left = window.innerWidth - menuWidth - 20;
        }
        if (top + menuHeight > window.innerHeight - 20) {
            top = window.innerHeight - menuHeight - 20;
        }

        menu.style.left = `${Math.max(10, left)}px`;
        menu.style.top = `${Math.max(10, top)}px`;
    },

    filterContextMenu(query) {
        const cleanQuery = query.toLowerCase().trim();
        document.querySelectorAll('#mapper-context-menu .menu-item-btn').forEach(btn => {
            const text = (btn.dataset.searchText || '').toLowerCase();
            if (!cleanQuery || text.includes(cleanQuery)) {
                btn.style.display = 'flex';
            } else {
                btn.style.display = 'none';
            }
        });
    },

    hideContextMenu() {
        const menu = document.getElementById('mapper-context-menu');
        if (menu) {
            menu.style.display = 'none';
        }
        this.activeContextMenuCell = null;
    },

    /**
     * Přepne přiřazení pole k dané buňce
     */
    toggleMapping(categoryKey, fieldKey, cellAddress) {
        this.ensureSettingsStructure();
        const locations = this.settings[categoryKey][fieldKey];
        const existingIndex = locations.findIndex(loc => 
            loc.file === this.currentFile && loc.sheet === this.currentSheet && loc.cell === cellAddress
        );

        if (existingIndex >= 0) {
            // Odebrat z této buňky
            locations.splice(existingIndex, 1);
            this.showToast(`Odebráno pole z buňky ${cellAddress}`, 'info');
        } else {
            // Přidat jako další (záložní / novou) lokaci
            locations.push({
                file: this.currentFile,
                sheet: this.currentSheet,
                cell: cellAddress
            });
            const priority = locations.length;
            const priorityText = priority > 1 ? ` (jako záložní #${priority})` : '';
            this.showToast(`Přiřazeno k ${cellAddress}${priorityText}`, 'success');
        }

        this.isDirty = true;
        this.hideContextMenu();
        this.renderGrid();
        this.renderLegend();
    },

    /**
     * Odebere všechna přiřazení z jedné buňky
     */
    removeCellMappings(cellAddress) {
        this.ensureSettingsStructure();
        let removedCount = 0;

        for (const catKey of Object.keys(this.schema)) {
            for (const fieldKey of Object.keys(this.schema[catKey].fields)) {
                const list = this.settings[catKey][fieldKey] || [];
                const filtered = list.filter(loc => 
                    !(loc.file === this.currentFile && loc.sheet === this.currentSheet && loc.cell === cellAddress)
                );
                if (filtered.length < list.length) {
                    removedCount += (list.length - filtered.length);
                    this.settings[catKey][fieldKey] = filtered;
                }
            }
        }

        this.isDirty = true;
        this.hideContextMenu();
        this.renderGrid();
        this.renderLegend();
        this.showToast(`Odebráno ${removedCount} mapování z buňky ${cellAddress}`, 'info');
    },

    /**
     * Vyčistí všechna mapování pro aktuální list
     */
    clearCurrentSheetMappings() {
        if (!confirm(`Opravdu chcete odebrat všechna mapování pro list "${this.currentSheet}" v souboru "${this.currentFile}"?`)) {
            return;
        }

        this.ensureSettingsStructure();
        let count = 0;
        for (const catKey of Object.keys(this.schema)) {
            for (const fieldKey of Object.keys(this.schema[catKey].fields)) {
                const list = this.settings[catKey][fieldKey] || [];
                const filtered = list.filter(loc => 
                    !(loc.file === this.currentFile && loc.sheet === this.currentSheet)
                );
                count += (list.length - filtered.length);
                this.settings[catKey][fieldKey] = filtered;
            }
        }

        this.isDirty = true;
        this.renderGrid();
        this.renderLegend();
        this.showToast(`Vyčištěno ${count} mapování na listu ${this.currentSheet}`, 'success');
    },

    /**
     * Obnoví výchozí mapování ze serveru
     */
    async resetToDefaults() {
        if (!confirm('Opravdu chcete obnovit výchozí šablonu mapování? Všechny neuložené i vlastní změny budou nahrazeny továrním nastavením.')) {
            return;
        }

        try {
            const res = await fetch('/api/settings/defaults');
            if (!res.ok) throw new Error('Nepodařilo se načíst výchozí šablonu');
            this.settings = await res.json();
            this.ensureSettingsStructure();
            this.isDirty = true;
            this.renderGrid();
            this.renderLegend();
            this.showToast('Výchozí šablona byla načtena. Nezapomeňte ji uložit.', 'success');
        } catch (err) {
            console.error(err);
            this.showToast('Chyba při obnovení: ' + err.message, 'error');
        }
    },

    renderLegend() {
        const legendGrid = document.getElementById('mapper-legend-grid');
        const countBadge = document.getElementById('mapper-total-count');
        if (!legendGrid || !this.schema) return;

        let totalMappings = 0;
        let html = '';

        for (const [catKey, catVal] of Object.entries(this.schema)) {
            let fieldsHtml = '';
            for (const [fieldKey, fieldVal] of Object.entries(catVal.fields)) {
                const locations = (this.settings[catKey] && this.settings[catKey][fieldKey]) || [];
                const currentSheetLocations = locations.filter(loc => 
                    loc.file === this.currentFile && loc.sheet === this.currentSheet
                );
                totalMappings += currentSheetLocations.length;

                const cellsSummary = currentSheetLocations.map(l => l.cell).join(', ') || 'nenamapováno';

                fieldsHtml += `
                    <div class="legend-field-item">
                        <div class="legend-field-name">
                            <span class="menu-item-color-dot" style="background-color: ${fieldVal.color};"></span>
                            <span>${fieldVal.label}</span>
                        </div>
                        <span class="legend-field-count ${currentSheetLocations.length > 0 ? 'has-items' : ''}" title="${cellsSummary}">
                            ${currentSheetLocations.length > 0 ? cellsSummary : '0'}
                        </span>
                    </div>
                `;
            }

            html += `
                <div class="legend-category-box">
                    <h4>
                        <span>${catVal.icon}</span>
                        <span>${catVal.label}</span>
                    </h4>
                    <div class="legend-field-list">
                        ${fieldsHtml}
                    </div>
                </div>
            `;
        }

        legendGrid.innerHTML = html;
        if (countBadge) {
            countBadge.textContent = `${totalMappings} namapovaných buněk na listu`;
        }
    },

    async saveSettings() {
        const saveBtn = document.getElementById('mapper-save-btn');
        if (saveBtn) {
            saveBtn.disabled = true;
            saveBtn.innerHTML = '⏳ Ukládám...';
        }

        try {
            const res = await fetch('/api/settings', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(this.settings)
            });

            if (!res.ok) {
                const errData = await res.json();
                throw new Error(errData.error || 'Chyba serveru při ukládání');
            }

            const result = await res.json();
            if (result.success) {
                this.isDirty = false;
                this.showToast('✅ Excel mapování bylo úspěšně uloženo!', 'success');
            } else {
                throw new Error(result.error || 'Uložení se nezdařilo');
            }
        } catch (err) {
            console.error('Chyba při ukládání:', err);
            this.showToast('❌ Chyba při ukládání: ' + err.message, 'error');
        } finally {
            if (saveBtn) {
                saveBtn.disabled = false;
                saveBtn.innerHTML = '💾 Uložit Excel mapování';
            }
        }
    },

    showToast(message, type = 'info') {
        if (typeof window.showToast === 'function') {
            window.showToast(message, type);
        } else {
            console.log(`[Toast ${type}]: ${message}`);
            alert(message);
        }
    }
};

// Automatické spuštění po načtení stránky
document.addEventListener('DOMContentLoaded', () => {
    if (document.getElementById('excel-mapper-root')) {
        ExcelMapper.init();
    }
});
