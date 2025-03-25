// Funkce pro zobrazení/skrytí formuláře pro úpravu
function showEditForm(employeeName) {
    const formId = `edit-form-${employeeName.replace(/ /g, '_')}`;
    const form = document.getElementById(formId);
    const button = document.querySelector(`button[aria-controls="${formId}"]`);
    const isExpanded = form.classList.contains('active');
    
    form.classList.toggle('active');
    button.setAttribute('aria-expanded', !isExpanded);
}

// Potvrzení smazání zaměstnance
document.addEventListener('click', function(e) {
    if (e.target.matches('button[data-confirm="delete"]')) {
        if (!confirm('Opravdu chcete smazat tohoto zaměstnance?')) {
            e.preventDefault();
        }
    }
}); 