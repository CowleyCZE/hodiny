/**
 * JavaScript logika pro stránku dynamických nastavení XLSX s rozšířenou podporou
 */

// Globální proměnné
let currentSettings = {};
let availableFiles = [];
let currentModal = {
    dataType: null,
    fieldKey: null,
    locationIndex: null,
    selectedCell: null
};

// Přednastavené údaje, které lze konfigurovat podle kategorií
const CONFIGURABLE_DATA_TYPES = {
    'weekly_time': {
        label: 'Týdenní evidence',
        description: 'Konfigurace pro ukládání týdenních časových záznamů',
        fields: {
            'employee_name': {
                label: 'Jméno zaměstnance',
                description: 'Místo, kam se ukládá jméno zaměstnance'
            },
            'date': {
                label: 'Datum',
                description: 'Místo, kam se ukládá datum záznamu'
            },
            'start_time': {
                label: 'Čas začátku',
                description: 'Místo, kam se ukládá čas začátku práce'
            },
            'end_time': {
                label: 'Čas konce',
                description: 'Místo, kam se ukládá čas konce práce'
            },
            'lunch_duration': {
                label: 'Doba oběda',
                description: 'Místo, kam se ukládá doba oběda'
            },
            'total_hours': {
                label: 'Celkové hodiny',
                description: 'Místo, kam se ukládají celkové odpracované hodiny'
            }
        }
    },
    'advances': {
        label: 'Zálohy a půjčky',
        description: 'Konfigurace pro ukládání záloh a půjček zaměstnanců',
        fields: {
            'employee_name': {
                label: 'Jméno zaměstnance',
                description: 'Místo, kam se ukládá jméno zaměstnance pro zálohy'
            },
            'amount_eur': {
                label: 'Částka EUR',
                description: 'Místo, kam se ukládá částka v eurech'
            },
            'amount_czk': {
                label: 'Částka CZK',
                description: 'Místo, kam se ukládá částka v korunách'
            },
            'date': {
                label: 'Datum zálohy',
                description: 'Místo, kam se ukládá datum zálohy'
            },
            'option_type': {
                label: 'Typ zálohy',
                description: 'Místo, kam se ukládá typ/kategorie zálohy'
            }
        }
    },
    'monthly_time': {
        label: 'Měsíční evidence',
        description: 'Konfigurace pro ukládání měsíčních časových záznamů (Hodiny2025)',
        fields: {
            'employee_name': {
                label: 'Jméno zaměstnance',
                description: 'Místo, kam se ukládá jméno zaměstnance v měsíční evidenci'
            },
            'date': {
                label: 'Datum',
                description: 'Místo, kam se ukládá datum v měsíční evidenci'
            },
            'start_time': {
                label: 'Čas začátku',
                description: 'Místo, kam se ukládá čas začátku v měsíční evidenci'
            },
            'end_time': {
                label: 'Čas konce',
                description: 'Místo, kam se ukládá čas konce v měsíční evidenci'
            },
            'lunch_hours': {
                label: 'Hodiny oběda',
                description: 'Místo, kam se ukládá doba oběda v hodinách'
            },
            'total_hours': {
                label: 'Celkové hodiny',
                description: 'Místo, kam se ukládají celkové hodiny v měsíční evidenci'
            },
            'overtime': {
                label: 'Přesčasy',
                description: 'Místo, kam se ukládají přesčasy'
            },
            'num_employees': {
                label: 'Počet zaměstnanců',
                description: 'Místo, kam se ukládá počet zaměstnanců'
            },
            'total_all_employees': {
                label: 'Celkové hodiny všech',
                description: 'Místo, kam se ukládá součet hodin všech zaměstnanců'
            }
        }
    },
    'projects': {
        label: 'Projekty',
        description: 'Konfigurace pro ukládání informací o projektech',
        fields: {
            'project_name': {
                label: 'Název projektu',
                description: 'Místo, kam se ukládá název projektu'
            },
            'start_date': {
                label: 'Datum začátku',
                description: 'Místo, kam se ukládá datum začátku projektu'
            },
            'end_date': {
                label: 'Datum konce',
                description: 'Místo, kam se ukládá datum konce projektu'
            }
        }
    }
};

// Inicializace po načtení stránky
document.addEventListener('DOMContentLoaded', function() {
    console.log('Nastaveni.js se načítá s rozšířenou podporou...');
    initializeSettingsPage();
    setupEventListeners();
});

/**
 * Inicializuje stránku nastavení
 */
async function initializeSettingsPage() {
    try {
        // Načte aktuální nastavení ze serveru
        await loadCurrentSettings();
        
        // Načte seznam dostupných Excel souborů
        await loadAvailableFiles();
        
        // Vygeneruje UI pro všechny kategorie
        generateSettingsUI();
        
        console.log('Nastavení úspěšně inicializována');
    } catch (error) {
        console.error('Chyba při inicializaci nastavení:', error);
        showErrorMessage('Chyba při načítání nastavení: ' + error.message);
    }
}

/**
 * Načte aktuální nastavení ze serveru
 */
async function loadCurrentSettings() {
    try {
        const response = await fetch('/api/settings');
        if (!response.ok) {
            throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }
        
        currentSettings = await response.json();
        console.log('Načtena nastavení:', currentSettings);
        
        // Ujisti se, že všechny kategorie existují
        for (const dataType of Object.keys(CONFIGURABLE_DATA_TYPES)) {
            if (!currentSettings[dataType]) {
                currentSettings[dataType] = {};
            }
            
            // Ujisti se, že všechna pole existují jako prázdné pole
            for (const fieldKey of Object.keys(CONFIGURABLE_DATA_TYPES[dataType].fields)) {
                if (!currentSettings[dataType][fieldKey]) {
                    currentSettings[dataType][fieldKey] = [];
                }
            }
        }
        
    } catch (error) {
        console.error('Chyba při načítání nastavení:', error);
        currentSettings = {};
        // Inicializuj prázdnou strukturu
        for (const dataType of Object.keys(CONFIGURABLE_DATA_TYPES)) {
            currentSettings[dataType] = {};
            for (const fieldKey of Object.keys(CONFIGURABLE_DATA_TYPES[dataType].fields)) {
                currentSettings[dataType][fieldKey] = [];
            }
        }
    }
}

/**
 * Načte seznam dostupných Excel souborů
 */
async function loadAvailableFiles() {
    try {
        const response = await fetch('/api/files');
        if (!response.ok) {
            throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }
        
        const data = await response.json();
        availableFiles = data.files || [];
        console.log('Načteny soubory:', availableFiles);
    } catch (error) {
        console.error('Chyba při načítání souborů:', error);
        availableFiles = [];
    }
}

/**
 * Vygeneruje UI pro všechny kategorie nastavení
 */
function generateSettingsUI() {
    const container = document.getElementById('settings-container');
    container.innerHTML = '';
    
    for (const [dataType, typeConfig] of Object.entries(CONFIGURABLE_DATA_TYPES)) {
        const categoryElement = createCategoryElement(dataType, typeConfig);
        container.appendChild(categoryElement);
    }
}

/**
 * Vytvoří element pro jednu kategorii dat
 */
function createCategoryElement(dataType, typeConfig) {
    const categoryDiv = document.createElement('div');
    categoryDiv.className = 'setting-item';
    categoryDiv.innerHTML = `
        <div class="setting-header">
            ${typeConfig.label}
            <div style="font-weight: normal; font-size: 0.9em; color: #6c757d;">
                ${typeConfig.description}
            </div>
        </div>
        <div id="fields-${dataType}"></div>
    `;
    
    const fieldsContainer = categoryDiv.querySelector(`#fields-${dataType}`);
    
    for (const [fieldKey, fieldConfig] of Object.entries(typeConfig.fields)) {
        const fieldElement = createFieldElement(dataType, fieldKey, fieldConfig);
        fieldsContainer.appendChild(fieldElement);
    }
    
    return categoryDiv;
}

/**
 * Vytvoří element pro jedno pole s možností více lokací
 */
function createFieldElement(dataType, fieldKey, fieldConfig) {
    const fieldDiv = document.createElement('div');
    fieldDiv.className = 'field-container';
    fieldDiv.style.marginBottom = '15px';
    fieldDiv.style.padding = '15px';
    fieldDiv.style.border = '1px solid #e9ecef';
    fieldDiv.style.borderRadius = '5px';
    fieldDiv.style.backgroundColor = '#ffffff';
    
    const locations = currentSettings[dataType][fieldKey] || [];
    
    fieldDiv.innerHTML = `
        <div style="font-weight: bold; margin-bottom: 10px;">
            ${fieldConfig.label}
            <span style="font-weight: normal; color: #6c757d; font-size: 0.9em;">
                - ${fieldConfig.description}
            </span>
        </div>
        <div id="locations-${dataType}-${fieldKey}"></div>
        <button type="button" class="btn btn-secondary btn-sm" 
                onclick="addNewLocation('${dataType}', '${fieldKey}')" 
                style="margin-top: 10px;">
            ➕ Přidat lokaci
        </button>
    `;
    
    const locationsContainer = fieldDiv.querySelector(`#locations-${dataType}-${fieldKey}`);
    
    if (locations.length === 0) {
        // Přidej alespoň jednu prázdnou lokaci
        addLocationElement(locationsContainer, dataType, fieldKey, 0, {});
    } else {
        locations.forEach((location, index) => {
            addLocationElement(locationsContainer, dataType, fieldKey, index, location);
        });
    }
    
    return fieldDiv;
}

/**
 * Přidá element pro jednu lokaci
 */
function addLocationElement(container, dataType, fieldKey, index, location) {
    const locationDiv = document.createElement('div');
    locationDiv.className = 'setting-row';
    locationDiv.style.marginBottom = '10px';
    locationDiv.style.padding = '10px';
    locationDiv.style.backgroundColor = '#f8f9fa';
    locationDiv.style.borderRadius = '3px';
    
    const currentFile = location.file || '';
    const currentSheet = location.sheet || '';
    const currentCell = location.cell || '';
    
    locationDiv.innerHTML = `
        <div class="form-group">
            <label>Excel soubor:</label>
            <select id="file-${dataType}-${fieldKey}-${index}" onchange="onFileChange('${dataType}', '${fieldKey}', ${index})">
                <option value="">-- Vyberte soubor --</option>
                ${availableFiles.map(file => 
                    `<option value="${file}" ${file === currentFile ? 'selected' : ''}>${file}</option>`
                ).join('')}
            </select>
        </div>
        
        <div class="form-group">
            <label>List:</label>
            <select id="sheet-${dataType}-${fieldKey}-${index}" onchange="onSheetChange('${dataType}', '${fieldKey}', ${index})" disabled>
                <option value="">-- Vyberte list --</option>
                ${currentSheet ? `<option value="${currentSheet}" selected>${currentSheet}</option>` : ''}
            </select>
        </div>
        
        <div class="form-group">
            <label>Buňka:</label>
            <input type="text" id="cell-${dataType}-${fieldKey}-${index}" value="${currentCell}" readonly
                   style="background-color: #e9ecef; cursor: pointer;" 
                   placeholder="Klikněte pro výběr">
        </div>
        
        <button type="button" class="btn btn-primary" 
                onclick="showSheetModal('${dataType}', '${fieldKey}', ${index})"
                ${!currentFile || !currentSheet ? 'disabled' : ''}>
            🔍 Zobrazit list
        </button>
        
        <button type="button" class="btn btn-secondary" 
                onclick="removeLocation('${dataType}', '${fieldKey}', ${index})">
            🗑️ Odebrat
        </button>
    `;
    
    container.appendChild(locationDiv);
    
    // Pokud má soubor, načti jeho listy
    if (currentFile) {
        loadSheetsForFile(currentFile, dataType, fieldKey, index);
    }
}

/**
 * Přidá novou lokaci pro pole
 */
function addNewLocation(dataType, fieldKey) {
    const container = document.querySelector(`#locations-${dataType}-${fieldKey}`);
    const currentLocations = container.children.length;
    addLocationElement(container, dataType, fieldKey, currentLocations, {});
}

/**
 * Odebere lokaci
 */
function removeLocation(dataType, fieldKey, index) {
    const container = document.querySelector(`#locations-${dataType}-${fieldKey}`);
    const locationElements = container.children;
    
    if (locationElements.length <= 1) {
        // Nech alespoň jednu lokaci, ale vymaž ji
        addLocationElement(container, dataType, fieldKey, 0, {});
        if (locationElements.length > 1) {
            locationElements[1].remove();
        }
    } else {
        // Odstraň vybranou lokaci a přečísluj zbytek
        Array.from(locationElements).forEach(element => element.remove());
        
        // Znovu vygeneruj s aktualizovanými indexy
        const currentLocations = currentSettings[dataType][fieldKey] || [];
        currentLocations.splice(index, 1);
        currentLocations.forEach((location, newIndex) => {
            addLocationElement(container, dataType, fieldKey, newIndex, location);
        });
        
        if (currentLocations.length === 0) {
            addLocationElement(container, dataType, fieldKey, 0, {});
        }
    }
}

/**
 * Handler pro změnu souboru
 */
async function onFileChange(dataType, fieldKey, index) {
    const fileSelect = document.getElementById(`file-${dataType}-${fieldKey}-${index}`);
    const sheetSelect = document.getElementById(`sheet-${dataType}-${fieldKey}-${index}`);
    const cellInput = document.getElementById(`cell-${dataType}-${fieldKey}-${index}`);
    const showButton = fileSelect.parentElement.parentElement.querySelector('button[onclick*="showSheetModal"]');
    
    // Reset dependent fields
    sheetSelect.innerHTML = '<option value="">-- Vyberte list --</option>';
    sheetSelect.disabled = true;
    cellInput.value = '';
    showButton.disabled = true;
    
    if (fileSelect.value) {
        await loadSheetsForFile(fileSelect.value, dataType, fieldKey, index);
    }
}

/**
 * Handler pro změnu listu
 */
function onSheetChange(dataType, fieldKey, index) {
    const sheetSelect = document.getElementById(`sheet-${dataType}-${fieldKey}-${index}`);
    const cellInput = document.getElementById(`cell-${dataType}-${fieldKey}-${index}`);
    const showButton = sheetSelect.parentElement.parentElement.querySelector('button[onclick*="showSheetModal"]');
    
    cellInput.value = '';
    
    if (sheetSelect.value) {
        showButton.disabled = false;
    } else {
        showButton.disabled = true;
    }
}

/**
 * Načte listy pro vybraný soubor
 */
async function loadSheetsForFile(filename, dataType, fieldKey, index) {
    try {
        const response = await fetch(`/api/sheets/${encodeURIComponent(filename)}`);
        if (!response.ok) {
            throw new Error(`Nepodařilo se načíst listy: ${response.statusText}`);
        }
        
        const data = await response.json();
        const sheetSelect = document.getElementById(`sheet-${dataType}-${fieldKey}-${index}`);
        const currentSheet = sheetSelect.value;
        
        sheetSelect.innerHTML = '<option value="">-- Vyberte list --</option>';
        data.sheets.forEach(sheetName => {
            const option = document.createElement('option');
            option.value = sheetName;
            option.textContent = sheetName;
            if (sheetName === currentSheet) {
                option.selected = true;
            }
            sheetSelect.appendChild(option);
        });
        
        sheetSelect.disabled = false;
        
        // Pokud byl list vybrán, aktivuj tlačítko
        if (sheetSelect.value) {
            const showButton = sheetSelect.parentElement.parentElement.querySelector('button[onclick*="showSheetModal"]');
            showButton.disabled = false;
        }
        
    } catch (error) {
        console.error('Chyba při načítání listů:', error);
        showErrorMessage('Chyba při načítání listů: ' + error.message);
    }
}

/**
 * Zobrazí modal s obsahem listu
 */
async function showSheetModal(dataType, fieldKey, index) {
    const fileSelect = document.getElementById(`file-${dataType}-${fieldKey}-${index}`);
    const sheetSelect = document.getElementById(`sheet-${dataType}-${fieldKey}-${index}`);
    
    if (!fileSelect.value || !sheetSelect.value) {
        showErrorMessage('Nejprve vyberte soubor a list');
        return;
    }
    
    currentModal = {
        dataType: dataType,
        fieldKey: fieldKey,
        locationIndex: index,
        selectedCell: null
    };
    
    try {
        const response = await fetch(`/api/sheet_content/${encodeURIComponent(fileSelect.value)}/${encodeURIComponent(sheetSelect.value)}`);
        if (!response.ok) {
            throw new Error(`Nepodařilo se načíst obsah listu: ${response.statusText}`);
        }
        
        const data = await response.json();
        
        document.getElementById('modal-title').textContent = `Výběr buňky - ${fileSelect.value}/${sheetSelect.value}`;
        generateSheetTable(data.data, data.rows, data.cols);
        
        document.getElementById('sheet-modal').style.display = 'block';
        
    } catch (error) {
        console.error('Chyba při načítání obsahu listu:', error);
        showErrorMessage('Chyba při načítání obsahu listu: ' + error.message);
    }
}

/**
 * Vygeneruje tabulku s obsahem listu
 */
function generateSheetTable(data, rows, cols) {
    const modalBody = document.getElementById('modal-body');
    
    let tableHTML = '<table class="sheet-table"><thead><tr><th></th>';
    
    // Hlavička sloupců (A, B, C, ...)
    for (let col = 1; col <= cols; col++) {
        const colLetter = String.fromCharCode(64 + col);
        tableHTML += `<th>${colLetter}</th>`;
    }
    tableHTML += '</tr></thead><tbody>';
    
    // Řádky s daty
    for (let row = 0; row < rows; row++) {
        tableHTML += `<tr><th>${row + 1}</th>`;
        for (let col = 0; col < cols; col++) {
            const cellValue = (data[row] && data[row][col]) || '';
            const cellAddress = String.fromCharCode(65 + col) + (row + 1);
            tableHTML += `<td onclick="selectCell('${cellAddress}')" data-cell="${cellAddress}">${cellValue}</td>`;
        }
        tableHTML += '</tr>';
    }
    tableHTML += '</tbody></table>';
    
    modalBody.innerHTML = tableHTML;
}

/**
 * Vybere buňku v tabulce
 */
function selectCell(cellAddress) {
    // Odstraň předchozí výběr
    document.querySelectorAll('.sheet-table td.selected').forEach(cell => {
        cell.classList.remove('selected');
    });
    
    // Přidej výběr k nové buňce
    const selectedCell = document.querySelector(`[data-cell="${cellAddress}"]`);
    if (selectedCell) {
        selectedCell.classList.add('selected');
        currentModal.selectedCell = cellAddress;
        
        // Zobraz informace o vybrané buňce
        document.getElementById('selected-cell-address').textContent = cellAddress;
        document.getElementById('selected-cell-info').style.display = 'block';
        
        // Automaticky nastavit buňku a zavřít modal po krátké pauze
        setTimeout(() => {
            setCellAndCloseModal();
        }, 500);
    }
}

/**
 * Nastaví vybranou buňku a zavře modal
 */
function setCellAndCloseModal() {
    if (currentModal.selectedCell) {
        const cellInput = document.getElementById(`cell-${currentModal.dataType}-${currentModal.fieldKey}-${currentModal.locationIndex}`);
        cellInput.value = currentModal.selectedCell;
        
        closeModal();
        showSuccessMessage(`Buňka ${currentModal.selectedCell} byla vybrána pro ${currentModal.fieldKey}`);
    }
}

/**
 * Zavře modal
 */
function closeModal() {
    document.getElementById('sheet-modal').style.display = 'none';
    document.getElementById('selected-cell-info').style.display = 'none';
    currentModal = {
        dataType: null,
        fieldKey: null,
        locationIndex: null,
        selectedCell: null
    };
}

/**
 * Shromáždí všechna nastavení z formuláře
 */
function collectAllSettings() {
    const settings = {};
    
    for (const dataType of Object.keys(CONFIGURABLE_DATA_TYPES)) {
        settings[dataType] = {};
        
        for (const fieldKey of Object.keys(CONFIGURABLE_DATA_TYPES[dataType].fields)) {
            settings[dataType][fieldKey] = [];
            
            const container = document.querySelector(`#locations-${dataType}-${fieldKey}`);
            if (container) {
                Array.from(container.children).forEach((locationElement, index) => {
                    const fileSelect = locationElement.querySelector(`select[id*="file"]`);
                    const sheetSelect = locationElement.querySelector(`select[id*="sheet"]`);
                    const cellInput = locationElement.querySelector(`input[id*="cell"]`);
                    
                    if (fileSelect && sheetSelect && cellInput && 
                        fileSelect.value && sheetSelect.value && cellInput.value) {
                        settings[dataType][fieldKey].push({
                            file: fileSelect.value,
                            sheet: sheetSelect.value,
                            cell: cellInput.value
                        });
                    }
                });
            }
        }
    }
    
    return settings;
}

/**
 * Uloží všechna nastavení
 */
async function saveAllSettings() {
    try {
        const settings = collectAllSettings();
        console.log('Ukládání nastavení:', settings);
        
        const response = await fetch('/api/settings', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(settings)
        });
        
        if (!response.ok) {
            const errorData = await response.json();
            throw new Error(errorData.error || `HTTP ${response.status}: ${response.statusText}`);
        }
        
        const result = await response.json();
        
        if (result.success) {
            showSuccessMessage('Všechna nastavení byla úspěšně uložena!');
            currentSettings = settings;
        } else {
            throw new Error(result.error || 'Neznámá chyba při ukládání');
        }
        
    } catch (error) {
        console.error('Chyba při ukládání nastavení:', error);
        showErrorMessage('Chyba při ukládání nastavení: ' + error.message);
    }
}

/**
 * Nastaví event listenery
 */
function setupEventListeners() {
    // Save all button
    document.getElementById('save-all-btn').addEventListener('click', saveAllSettings);
    
    // Modal close handlers
    document.querySelector('.close').addEventListener('click', closeModal);
    document.getElementById('sheet-modal').addEventListener('click', function(e) {
        if (e.target === this) {
            closeModal();
        }
    });
    
    // ESC key to close modal
    document.addEventListener('keydown', function(e) {
        if (e.key === 'Escape') {
            closeModal();
        }
    });
}

/**
 * Zobrazí chybovou zprávu
 */
function showErrorMessage(message) {
    console.error(message);
    alert('Chyba: ' + message);
}

/**
 * Zobrazí úspěšnou zprávu  
 */
function showSuccessMessage(message) {
    console.log(message);
    // Můžeme přidat toast notifikaci nebo podobně
    alert('Úspěch: ' + message);
}
async function initializeSettingsPage() {
    try {
        // Načte současná nastavení a dostupné soubory
        await Promise.all([
            loadCurrentSettings(),
            loadAvailableFiles()
        ]);
        
        // Vykreslí formulář
        renderSettingsForm();
        
    } catch (error) {
        console.error('Chyba při inicializaci:', error);
        showMessage('Chyba při načítání dat.', 'error');
    }
}

/**
 * Načte současná nastavení z API
 */
async function loadCurrentSettings() {
    try {
        const response = await fetch('/api/settings');
        if (response.ok) {
            currentSettings = await response.json();
            console.log('Načtená nastavení:', currentSettings);
        } else {
            console.warn('Žádná nastavení nenalezena, používám výchozí');
            currentSettings = {};
        }
    } catch (error) {
        console.error('Chyba při načítání nastavení:', error);
        currentSettings = {};
    }
}

/**
 * Načte seznam dostupných Excel souborů
 */
async function loadAvailableFiles() {
    try {
        const response = await fetch('/api/files');
        const data = await response.json();
        if (response.ok) {
            availableFiles = data.files || [];
            console.log('Dostupné soubory:', availableFiles);
        } else {
            throw new Error(data.error || 'Chyba při načítání souborů');
        }
    } catch (error) {
        console.error('Chyba při načítání souborů:', error);
        availableFiles = [];
        showMessage('Chyba při načítání Excel souborů.', 'error');
    }
}

/**
 * Vykreslí formulář s nastavením
 */
function renderSettingsForm() {
    const container = document.getElementById('settings-container');
    container.innerHTML = '';
    
    // Pro každý konfigurovatelný údaj vytvoř řádek
    Object.keys(CONFIGURABLE_FIELDS).forEach(fieldKey => {
        const field = CONFIGURABLE_FIELDS[fieldKey];
        const setting = currentSettings[fieldKey] || {};
        
        const settingDiv = document.createElement('div');
        settingDiv.className = 'setting-item';
        settingDiv.innerHTML = `
            <div class="setting-header">
                ${field.label}
                <small style="display: block; font-weight: normal; color: #6c757d;">
                    ${field.description}
                </small>
            </div>
            <div class="setting-row">
                <div class="form-group">
                    <label for="file-${fieldKey}">Excel soubor:</label>
                    <select id="file-${fieldKey}" data-field="${fieldKey}">
                        <option value="">-- Vyberte soubor --</option>
                        ${availableFiles.map(file => 
                            `<option value="${file}" ${setting.file === file ? 'selected' : ''}>${file}</option>`
                        ).join('')}
                    </select>
                </div>
                <div class="form-group">
                    <label for="sheet-${fieldKey}">List:</label>
                    <select id="sheet-${fieldKey}" data-field="${fieldKey}" disabled>
                        <option value="">-- Nejprve vyberte soubor --</option>
                    </select>
                </div>
                <div class="form-group">
                    <label>Buňka:</label>
                    <input type="text" id="cell-${fieldKey}" readonly placeholder="-- Nevybráno --" 
                           value="${setting.cell || ''}" style="background: #f8f9fa;">
                </div>
                <button class="btn btn-primary" id="preview-${fieldKey}" data-field="${fieldKey}" disabled>
                    👁️ Zobrazit list
                </button>
                <button class="btn btn-secondary" id="clear-${fieldKey}" data-field="${fieldKey}">
                    🗑️ Vymazat
                </button>
            </div>
            <div class="current-setting">
                Aktuálně: ${setting.file && setting.sheet && setting.cell ? 
                    `${setting.file} → ${setting.sheet} → ${setting.cell}` : 
                    'Nenamapováno'}
            </div>
        `;
        
        container.appendChild(settingDiv);
        
        // Pokud je soubor vybrán, načti listy
        if (setting.file) {
            loadSheetsForFile(fieldKey, setting.file, setting.sheet);
        }
    });
}

/**
 * Nastavuje event listenery
 */
function setupEventListeners() {
    // Při změně souboru načti listy
    document.addEventListener('change', function(e) {
        if (e.target.id && e.target.id.startsWith('file-')) {
            const fieldKey = e.target.dataset.field;
            const filename = e.target.value;
            
            if (filename) {
                loadSheetsForFile(fieldKey, filename);
            } else {
                clearSheetAndCell(fieldKey);
            }
        }
        
        // Při změně listu aktivuj tlačítko náhledu
        if (e.target.id && e.target.id.startsWith('sheet-')) {
            const fieldKey = e.target.dataset.field;
            const sheetSelect = e.target;
            const previewBtn = document.getElementById(`preview-${fieldKey}`);
            
            if (sheetSelect.value) {
                previewBtn.disabled = false;
            } else {
                previewBtn.disabled = true;
                // Vymaž buňku pokud není list vybrán
                document.getElementById(`cell-${fieldKey}`).value = '';
            }
        }
    });
    
    // Tlačítka pro zobrazení listu
    document.addEventListener('click', function(e) {
        if (e.target.id && e.target.id.startsWith('preview-')) {
            const fieldKey = e.target.dataset.field;
            showSheetPreview(fieldKey);
        }
        
        // Tlačítka pro vymazání
        if (e.target.id && e.target.id.startsWith('clear-')) {
            const fieldKey = e.target.dataset.field;
            clearFieldSetting(fieldKey);
        }
    });
    
    // Tlačítko pro uložení všech nastavení
    document.getElementById('save-all-btn').addEventListener('click', saveAllSettings);
    
    // Modal ovládání
    setupModalEventListeners();
}

/**
 * Načte listy pro zadaný soubor
 */
async function loadSheetsForFile(fieldKey, filename, selectedSheet = null) {
    const sheetSelect = document.getElementById(`sheet-${fieldKey}`);
    const previewBtn = document.getElementById(`preview-${fieldKey}`);
    
    try {
        sheetSelect.innerHTML = '<option value="">Načítám...</option>';
        sheetSelect.disabled = true;
        previewBtn.disabled = true;
        
        const response = await fetch(`/api/sheets/${encodeURIComponent(filename)}`);
        const data = await response.json();
        
        if (response.ok) {
            const sheets = data.sheets || [];
            sheetSelect.innerHTML = '<option value="">-- Vyberte list --</option>';
            
            sheets.forEach(sheet => {
                const option = document.createElement('option');
                option.value = sheet;
                option.textContent = sheet;
                if (selectedSheet === sheet) {
                    option.selected = true;
                }
                sheetSelect.appendChild(option);
            });
            
            sheetSelect.disabled = false;
            
            // Pokud je list vybrán, aktivuj tlačítko náhledu
            if (selectedSheet && sheets.includes(selectedSheet)) {
                previewBtn.disabled = false;
            }
            
        } else {
            throw new Error(data.error || 'Chyba při načítání listů');
        }
    } catch (error) {
        console.error('Chyba při načítání listů:', error);
        sheetSelect.innerHTML = '<option value="">Chyba při načítání</option>';
        showMessage('Chyba při načítání listů ze souboru.', 'error');
    }
}

/**
 * Vymaže nastavení listu a buňky
 */
function clearSheetAndCell(fieldKey) {
    const sheetSelect = document.getElementById(`sheet-${fieldKey}`);
    const cellInput = document.getElementById(`cell-${fieldKey}`);
    const previewBtn = document.getElementById(`preview-${fieldKey}`);
    
    sheetSelect.innerHTML = '<option value="">-- Nejprve vyberte soubor --</option>';
    sheetSelect.disabled = true;
    cellInput.value = '';
    previewBtn.disabled = true;
}

/**
 * Vymaže celé nastavení pole
 */
function clearFieldSetting(fieldKey) {
    document.getElementById(`file-${fieldKey}`).value = '';
    clearSheetAndCell(fieldKey);
    showMessage(`Nastavení pro "${CONFIGURABLE_FIELDS[fieldKey].label}" bylo vymazáno.`, 'info');
}

/**
 * Zobrazí náhled listu pro výběr buňky
 */
async function showSheetPreview(fieldKey) {
    const filename = document.getElementById(`file-${fieldKey}`).value;
    const sheetname = document.getElementById(`sheet-${fieldKey}`).value;
    
    if (!filename || !sheetname) {
        showMessage('Nejprve vyberte soubor a list.', 'warning');
        return;
    }
    
    currentModal.fieldKey = fieldKey;
    currentModal.selectedCell = null;
    
    try {
        const modal = document.getElementById('sheet-modal');
        const modalTitle = document.getElementById('modal-title');
        const modalBody = document.getElementById('modal-body');
        const cellInfo = document.getElementById('selected-cell-info');
        
        modalTitle.textContent = `${filename} - ${sheetname}`;
        modalBody.innerHTML = '<p>Načítám obsah listu...</p>';
        cellInfo.style.display = 'none';
        modal.style.display = 'block';
        
        const response = await fetch(`/api/sheet_content/${encodeURIComponent(filename)}/${encodeURIComponent(sheetname)}`);
        const data = await response.json();
        
        if (response.ok) {
            renderSheetTable(data.data, data.rows, data.cols);
        } else {
            throw new Error(data.error || 'Chyba při načítání obsahu');
        }
        
    } catch (error) {
        console.error('Chyba při zobrazení náhledu:', error);
        showMessage('Chyba při načítání náhledu listu.', 'error');
        document.getElementById('sheet-modal').style.display = 'none';
    }
}

/**
 * Vykreslí tabulku s obsahem listu
 */
function renderSheetTable(data, rows, cols) {
    const modalBody = document.getElementById('modal-body');
    
    let tableHTML = '<table class="sheet-table"><thead><tr><th></th>';
    
    // Hlavička se sloupci (A, B, C, ...)
    for (let col = 1; col <= cols; col++) {
        const columnLetter = String.fromCharCode(64 + col); // A=65, B=66, ...
        tableHTML += `<th>${columnLetter}</th>`;
    }
    tableHTML += '</tr></thead><tbody>';
    
    // Řádky s daty
    for (let row = 0; row < rows; row++) {
        tableHTML += `<tr><th>${row + 1}</th>`;
        for (let col = 0; col < cols; col++) {
            const cellValue = data[row] && data[row][col] ? data[row][col] : '';
            const cellAddress = String.fromCharCode(65 + col) + (row + 1);
            tableHTML += `<td data-address="${cellAddress}" title="Buňka ${cellAddress}">${cellValue}</td>`;
        }
        tableHTML += '</tr>';
    }
    tableHTML += '</tbody></table>';
    
    modalBody.innerHTML = tableHTML;
    
    // Přidej event listenery pro klikání na buňky
    modalBody.querySelectorAll('td[data-address]').forEach(cell => {
        cell.addEventListener('click', function() {
            selectCell(this);
        });
    });
}

/**
 * Vybere buňku v tabulce
 */
function selectCell(cellElement) {
    // Odstraň předchozí výběr
    document.querySelectorAll('.sheet-table td.selected').forEach(cell => {
        cell.classList.remove('selected');
    });
    
    // Označ novou buňku
    cellElement.classList.add('selected');
    
    const cellAddress = cellElement.dataset.address;
    currentModal.selectedCell = cellAddress;
    
    // Zobraz informace o vybrané buňce
    const cellInfo = document.getElementById('selected-cell-info');
    const cellAddressSpan = document.getElementById('selected-cell-address');
    cellAddressSpan.textContent = cellAddress;
    cellInfo.style.display = 'block';
    
    // Automaticky zavři modal a nastav buňku po 1 sekundě
    setTimeout(() => {
        closeModalAndSetCell();
    }, 1000);
}

/**
 * Zavře modal a nastaví vybranou buňku
 */
function closeModalAndSetCell() {
    if (currentModal.fieldKey && currentModal.selectedCell) {
        const cellInput = document.getElementById(`cell-${currentModal.fieldKey}`);
        cellInput.value = currentModal.selectedCell;
        
        showMessage(`Buňka ${currentModal.selectedCell} byla vybrána pro "${CONFIGURABLE_FIELDS[currentModal.fieldKey].label}".`, 'success');
    }
    
    document.getElementById('sheet-modal').style.display = 'none';
    currentModal = { fieldKey: null, selectedCell: null };
}

/**
 * Nastavuje event listenery pro modal
 */
function setupModalEventListeners() {
    const modal = document.getElementById('sheet-modal');
    const closeBtn = modal.querySelector('.close');
    
    // Zavření křížkem
    closeBtn.addEventListener('click', function() {
        modal.style.display = 'none';
    });
    
    // Zavření kliknutím mimo modal
    window.addEventListener('click', function(e) {
        if (e.target === modal) {
            modal.style.display = 'none';
        }
    });
}

/**
 * Uloží všechna nastavení
 */
async function saveAllSettings() {
    const saveBtn = document.getElementById('save-all-btn');
    const originalText = saveBtn.textContent;
    
    try {
        saveBtn.textContent = '💾 Ukládám...';
        saveBtn.disabled = true;
        
        // Sestaví objekt s nastavením
        const settingsToSave = {};
        
        Object.keys(CONFIGURABLE_FIELDS).forEach(fieldKey => {
            const file = document.getElementById(`file-${fieldKey}`).value;
            const sheet = document.getElementById(`sheet-${fieldKey}`).value;
            const cell = document.getElementById(`cell-${fieldKey}`).value;
            
            if (file && sheet && cell) {
                settingsToSave[fieldKey] = {
                    file: file,
                    sheet: sheet,
                    cell: cell
                };
            }
        });
        
        // Odešle na server
        const response = await fetch('/api/settings', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(settingsToSave)
        });
        
        const result = await response.json();
        
        if (response.ok) {
            currentSettings = settingsToSave;
            showMessage('Všechna nastavení byla úspěšně uložena!', 'success');
            renderSettingsForm(); // Obnovit zobrazení
        } else {
            throw new Error(result.error || 'Chyba při ukládání');
        }
        
    } catch (error) {
        console.error('Chyba při ukládání:', error);
        showMessage('Chyba při ukládání nastavení.', 'error');
    } finally {
        saveBtn.textContent = originalText;
        saveBtn.disabled = false;
    }
}

/**
 * Zobrazí zprávu uživateli
 */
function showMessage(message, type = 'info') {
    // Najdi nebo vytvoř kontainer pro zprávy
    let messagesContainer = document.querySelector('.flash-messages');
    if (!messagesContainer) {
        messagesContainer = document.createElement('ul');
        messagesContainer.className = 'flash-messages';
        messagesContainer.setAttribute('role', 'alert');
        messagesContainer.setAttribute('aria-live', 'polite');
        
        const mainContainer = document.querySelector('.settings-container');
        mainContainer.insertBefore(messagesContainer, mainContainer.firstChild.nextSibling);
    }
    
    // Vytvoř novou zprávu
    const messageItem = document.createElement('li');
    messageItem.className = `flash-message ${type}`;
    messageItem.textContent = message;
    
    // Přidej na začátek
    messagesContainer.insertBefore(messageItem, messagesContainer.firstChild);
    
    // Automaticky odstraň po 5 sekundách
    setTimeout(() => {
        if (messageItem.parentNode) {
            messageItem.parentNode.removeChild(messageItem);
        }
    }, 5000);
}