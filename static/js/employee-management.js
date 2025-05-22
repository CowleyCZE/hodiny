// Funkce pro zobrazení/skrytí formuláře pro úpravu jména zaměstnance.
// employeeName se používá k sestavení unikátního ID formuláře.
// Mezery ve jméně jsou nahrazeny podtržítky, aby se vytvořilo platné ID.
// Např. pro "Jan Novák" bude ID formuláře "edit-form-Jan_Novák".
function showEditForm(employeeName) {
    const formId = `edit-form-${employeeName.replace(/ /g, '_')}`; // Nahrazení mezer podtržítky pro ID
    const form = document.getElementById(formId);
    const button = document.querySelector(`button[aria-controls="${formId}"]`);
    const isExpanded = form.classList.contains('active');
    
    form.classList.toggle('active'); // Přepne viditelnost formuláře
    button.setAttribute('aria-expanded', !isExpanded); // Aktualizuje ARIA atribut pro přístupnost
}

// Globální event listener pro kliknutí, který odchytává kliknutí na tlačítka pro smazání.
// Používá se delegování událostí, aby nebylo nutné připojovat listener ke každému tlačítku zvlášť.
document.addEventListener('click', function(e) {
    // Kontroluje, zda kliknutý prvek je tlačítko s atributem data-confirm="delete"
    if (e.target.matches('button[data-confirm="delete"]')) {
        // Zobrazí nativní potvrzovací dialog prohlížeče.
        if (!confirm('Opravdu chcete smazat tohoto zaměstnance?')) {
            e.preventDefault(); // Pokud uživatel nepotvrdí, zabrání se výchozí akci (např. odeslání formuláře).
        }
    }
}); 