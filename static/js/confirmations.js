// Zobrazí potvrzovací dialog pro smazání zaměstnance.
function confirmDeleteEmployee() {
    return confirm("Opravdu chcete smazat tohoto zaměstnance?");
}

// Zobrazí potvrzovací dialog pro odeslání emailu.
function confirmSendEmail() {
    return confirm("Opravdu chcete odeslat email?");
}

// Zobrazí potvrzovací dialog pro smazání záznamu (obecná funkce).
// Tato funkce se aktuálně v HTML šablonách explicitně nevolá, 
// ale je zde pro případné budoucí použití.
function confirmDeleteRecord() {
    return confirm("Opravdu chcete smazat tento záznam?");
}
