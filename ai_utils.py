import ollama

def classify_work_description(description: str) -> str:
    """
    Používá Ollama model pro kategorizaci popisu práce.

    Args:
        description: Textový popis pracovní činnosti.

    Returns:
        Název kategorie (např. "Administrativa", "Vývoj", "Schůzka, "Support", "Ostatní").
    """
    try:
        # Použijeme model gemma, který máte stažený
        # Můžete zde specifikovat i jiný model, pokud ho máte nebo stáhnete
        response = ollama.chat(model='gemma', messages=[
            {
                'role': 'system',
                'content': 'Jsi asistent pro kategorizaci pracovních úkonů. Kategorizuj následující popis práce do jedné z těchto kategorií: Administrativa, Vývoj, Schůzka, Support, Ostatní. Odpověz pouze názvem kategorie.'
            },
            {
                'role': 'user',
                'content': f'Popis práce: "{description}"'
            },
        ])
        category = response['message']['content'].strip()
        return category if category in ["Administrativa", "Vývoj", "Schůzka", "Support", "Ostatní"] else "Ostatní"
    except Exception as e:
        print(f"Chyba při komunikaci s Ollamou: {e}")
        return "Ostatní"

if __name__ == "__main__":
    # Příklad použití
    print(f"Příklad 1: {classify_work_description('Napsat kód pro novou funkci')}")
    print(f"Příklad 2: {classify_work_description('Vyplnit formuláře pro dovolenou')}")
    print(f"Příklad 3: {classify_work_description('Denní stand-up meeting')}")
    print(f"Příklad 4: {classify_work_description('Opravit chybu na serveru')}")
    print(f"Příklad 5: {classify_work_description('Procházka se psem')}")