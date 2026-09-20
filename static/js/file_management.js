/**
 * static/js/file_management.js
 * Správa souborů XLSX v nastavení (Tab 3)
 */

let fmAvailableFiles = [];

document.addEventListener('DOMContentLoaded', () => {
    initFileManagement();
});

async function initFileManagement() {
    const container = document.getElementById('file-management-container');
    if (!container) return;

    try {
        const res = await fetch('/api/files');
        if (!res.ok) throw new Error('Chyba při načítání souborů');
        const data = await res.json();
        fmAvailableFiles = data.files || [];
        renderFileManagementList();
    } catch (err) {
        console.error('Chyba při inicializaci správy souborů:', err);
    }
}

function renderFileManagementList() {
    const container = document.getElementById('file-management-container');
    if (!container) return;

    if (fmAvailableFiles.length === 0) {
        container.innerHTML = '<div class="text-center py-4 text-muted">Žádné XLSX soubory nebyly nalezeny.</div>';
        return;
    }

    container.innerHTML = '';
    fmAvailableFiles.forEach(filename => {
        const item = document.createElement('div');
        item.className = 'file-item';
        item.innerHTML = `
            <div class="file-info">
                <span class="file-icon">📄</span>
                <span class="file-name" id="fm-filename-${filename}">${filename}</span>
                <input type="text" class="rename-input form-control-modern py-1 px-2" id="fm-rename-${filename}" value="${filename.replace('.xlsx', '')}" style="display: none; width: 250px;">
            </div>
            <div class="file-actions">
                <button type="button" class="btn-modern btn-primary btn-sm" onclick="window.open('/excel_viewer?file=' + encodeURIComponent('${filename}'), '_blank')">👁️ Prohlížeč</button>
                <button type="button" class="btn-modern btn-warning btn-sm" onclick="fmStartRename('${filename}')">✏️ Přejmenovat</button>
                <button type="button" class="btn-modern btn-success btn-sm" onclick="fmConfirmRename('${filename}')" id="fm-confirm-${filename}" style="display: none;">✅ Potvrdit</button>
                <button type="button" class="btn-modern btn-secondary btn-sm" onclick="fmCancelRename('${filename}')" id="fm-cancel-${filename}" style="display: none;">❌ Zrušit</button>
            </div>
        `;
        container.appendChild(item);
    });
}

function fmStartRename(filename) {
    const nameSpan = document.getElementById(`fm-filename-${filename}`);
    const renameInput = document.getElementById(`fm-rename-${filename}`);
    const confirmBtn = document.getElementById(`fm-confirm-${filename}`);
    const cancelBtn = document.getElementById(`fm-cancel-${filename}`);

    nameSpan.style.display = 'none';
    renameInput.style.display = 'block';
    confirmBtn.style.display = 'inline-block';
    cancelBtn.style.display = 'inline-block';

    renameInput.focus();
    renameInput.select();
}

function fmCancelRename(filename) {
    const nameSpan = document.getElementById(`fm-filename-${filename}`);
    const renameInput = document.getElementById(`fm-rename-${filename}`);
    const confirmBtn = document.getElementById(`fm-confirm-${filename}`);
    const cancelBtn = document.getElementById(`fm-cancel-${filename}`);

    nameSpan.style.display = 'block';
    renameInput.style.display = 'none';
    confirmBtn.style.display = 'none';
    cancelBtn.style.display = 'none';

    renameInput.value = filename.replace('.xlsx', '');
}

async function fmConfirmRename(oldFilename) {
    const renameInput = document.getElementById(`fm-rename-${oldFilename}`);
    const newBaseName = renameInput.value.trim();

    if (!newBaseName) {
        alert('Název souboru nemůže být prázdný');
        return;
    }

    const newFilename = newBaseName.endsWith('.xlsx') ? newBaseName : newBaseName + '.xlsx';

    if (newFilename === oldFilename) {
        fmCancelRename(oldFilename);
        return;
    }

    if (fmAvailableFiles.includes(newFilename)) {
        alert(`Soubor s názvem "${newFilename}" již existuje`);
        return;
    }

    try {
        const response = await fetch('/api/files/rename', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                old_filename: oldFilename,
                new_filename: newFilename
            })
        });

        const result = await response.json();
        if (result.success) {
            if (typeof showToast === 'function') {
                showToast(`Soubor "${oldFilename}" byl přejmenován na "${newFilename}"`, 'success');
            } else {
                alert(`Soubor byl přejmenován na "${newFilename}"`);
            }

            const index = fmAvailableFiles.indexOf(oldFilename);
            if (index !== -1) {
                fmAvailableFiles[index] = newFilename;
            }
            renderFileManagementList();

            // Pokud existuje ExcelMapper, aktualizovat i jeho seznam
            if (window.ExcelMapper && typeof ExcelMapper.loadFiles === 'function') {
                await ExcelMapper.loadFiles();
                ExcelMapper.renderToolbar();
                if (ExcelMapper.currentFile === oldFilename) {
                    await ExcelMapper.selectFile(newFilename);
                }
            }
        } else {
            alert(result.error || 'Neznámá chyba při přejmenování');
        }
    } catch (err) {
        console.error('Chyba při přejmenování souboru:', err);
        alert('Chyba při přejmenování: ' + err.message);
    }
}
