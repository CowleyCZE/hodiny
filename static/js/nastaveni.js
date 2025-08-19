/**
 * JavaScript logika pro stránku dynamických nastavení XLSX
 */

// Globální proměnné
let currentSettings = {};
let availableFiles = [];
let currentModal = {
    fieldKey: null,
    selectedCell: null
};

// Přednastavené údaje, které lze konfigurovat
const CONFIGURABLE_FIELDS = {
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
};

// Inicializace po načtení stránky
document.addEventListener('DOMContentLoaded', function() {
    console.log('Nastaveni.js se načítá...');
    initializeSettingsPage();
    setupEventListeners();
});

/**
 * Inicializuje stránku nastavení
 */
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