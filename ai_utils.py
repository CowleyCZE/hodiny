"""Optional Ollama-based work-description classification."""

try:
    import ollama
except ImportError:  # Ollama is optional; the application has a safe fallback.
    ollama = None


VALID_CATEGORIES = {"Administrativa", "Vývoj", "Schůzka", "Support", "Ostatní"}


def classify_work_description(description: str) -> str:
    """
    Používá Ollama model pro kategorizaci popisu práce.

    Args:
        description: Textový popis pracovní činnosti.

    Returns:
        Název kategorie nebo ``Ostatní``, pokud Ollama není dostupná či selže.
    """
    if ollama is None:
        return "Ostatní"

    try:
        response = ollama.chat(
            model="gemma",
            messages=[
                {
                    "role": "system",
                    "content": (
                        "Jsi asistent pro kategorizaci pracovních úkonů. Kategorizuj následující "
                        "popis práce do jedné z těchto kategorií: Administrativa, Vývoj, Schůzka, "
                        "Support, Ostatní. Odpověz pouze názvem kategorie."
                    ),
                },
                {"role": "user", "content": f'Popis práce: "{description}"'},
            ],
        )
        category = response["message"]["content"].strip()
        return category if category in VALID_CATEGORIES else "Ostatní"
    except Exception as error:
        print(f"Chyba při komunikaci s Ollamou: {error}")
        return "Ostatní"


if __name__ == "__main__":
    print(f"Příklad 1: {classify_work_description('Napsat kód pro novou funkci')}")
    print(f"Příklad 2: {classify_work_description('Vyplnit formuláře pro dovolenou')}")
    print(f"Příklad 3: {classify_work_description('Denní stand-up meeting')}")
    print(f"Příklad 4: {classify_work_description('Opravit chybu na serveru')}")
    print(f"Příklad 5: {classify_work_description('Procházka se psem')}")
