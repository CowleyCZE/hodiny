// static/js/voice-handler.js
const voiceButton = document.getElementById('voice-button');
const voiceResult = document.getElementById('voice-result');
const loadingIndicator = document.getElementById('loading-indicator');

let recognition;
let isListening = false;

// Inicializace Web Speech API
try {
    // Pro Chrome a podobné prohlížeče
    window.SpeechRecognition = window.SpeechRecognition || webkitSpeechRecognition;
    recognition = new window.SpeechRecognition();
    
    // Nastavení parametrů
    recognition.continuous = false; // Rozpoznávání se zastaví po první promluvě.
    recognition.interimResults = false; // Chceme pouze finální výsledky.
    recognition.lang = 'cs-CZ'; // Nastavení jazyka na češtinu.
    recognition.maxAlternatives = 1; // Počet alternativních transkriptů (chceme jen nejlepší).
    
    // Event handler: Spustí se, když Web Speech API začne naslouchat.
    recognition.onstart = () => {
        isListening = true;
        voiceButton.classList.add('active'); // Vizuální zpětná vazba tlačítka
        voiceResult.textContent = 'Naslouchám...'; // Informace pro uživatele
        loadingIndicator.style.display = 'inline-block'; // Zobrazení indikátoru načítání
    };
    
    // Event handler: Spustí se, když Web Speech API přestane naslouchat.
    recognition.onend = () => {
        isListening = false;
        voiceButton.classList.remove('active');
        loadingIndicator.style.display = 'none';
    };
    
    // Event handler: Spustí se, když jsou k dispozici výsledky rozpoznávání.
    recognition.onresult = (event) => {
        const transcript = event.results[0][0].transcript.trim(); // Získání rozpoznaného textu
        voiceResult.textContent = `Rozpoznáno: ${transcript}`; // Zobrazení textu uživateli
        
        // Odeslání rozpoznaného textu na server k dalšímu zpracování
        sendVoiceCommandToServer(transcript);
    };
    
    // Event handler: Spustí se při chybě v procesu rozpoznávání.
    recognition.onerror = (event) => {
        console.error('Chyba při hlasovém rozpoznávání:', event.error); // Logování chyby do konzole
        voiceResult.textContent = `Chyba: ${event.error}`; // Zobrazení chyby uživateli
        isListening = false;
        voiceButton.classList.remove('active');
        loadingIndicator.style.display = 'none';
    };
    
} catch (error) {
    console.error('Web Speech API není podporováno v tomto prohlížeči:', error);
    voiceButton.disabled = true;
    voiceResult.textContent = 'Hlasové rozpoznávání není podporováno v tomto prohlížeči.';
}

// Funkce pro spuštění/nastavení hlasového vstupu
voiceButton.addEventListener('click', () => {
    if (!isListening) {
        recognition.start();
    } else {
        recognition.stop();
    }
});

// Funkce pro odeslání rozpoznaného textu (hlasového příkazu) na server
function sendVoiceCommandToServer(text) {
    // Zobrazení indikátoru načítání, protože odesíláme data a čekáme na odpověď
    loadingIndicator.style.display = 'inline-block'; 
    voiceResult.textContent = `Zpracovávám: "${text}"`; // Indikace zpracování

    fetch('/voice-command', { // Cílová URL na serveru
        method: 'POST', // Metoda HTTP požadavku
        headers: {
            'Content-Type': 'application/json',
            'X-Requested-With': 'XMLHttpRequest'
        },
        body: JSON.stringify({ command: text })
    })
    .then(response => {
        if (!response.ok) {
            throw new Error('Síťová chyba');
        }
        return response.json();
    })
    .then(data => {
        if (data.success) {
            // Pokud server odpověděl úspěšně, zobrazíme výsledek a provedeme akci
            displaySuccessResult(data); // Zobrazí formátovaný úspěšný výsledek
            handleVoiceAction(data);    // Provede akci na základě rozpoznaných entit
        } else {
            // Pokud server odpověděl s chybou, zobrazíme ji
            displayErrorResult(data);
        }
    })
    .catch(error => {
        // Zpracování chyb sítě nebo jiných problémů s fetch požadavkem
        console.error('Chyba při odesílání hlasového příkazu:', error);
        voiceResult.textContent = 'Nepodařilo se odeslat hlasový příkaz na server.';
    })
    .finally(() => {
        // Skryjeme indikátor načítání po dokončení (ať už úspěšně nebo neúspěšně)
        loadingIndicator.style.display = 'none';
    });
}

// Zobrazení úspěšného výsledku zpracování serverem
function displaySuccessResult(data) {
    const confidence = (data.confidence * 100).toFixed(1); // Spolehlivost rozpoznání
    // Sestavení HTML pro zobrazení výsledků a entit
    voiceResult.innerHTML = `
        <div class="voice-result-success">
            <p>Úspěšně zpracováno (spolehlivost: ${confidence}%)</p>
            <ul>
                ${Object.entries(data.entities).map(([key, value]) => 
                    `<li><strong>${key}:</strong> ${value === null || value === undefined ? '<em>N/A</em>' : value}</li>`).join('')}
            </ul>
            ${data.operation_result ? `<p>Výsledek operace: ${data.operation_result.message || 'Neznámá odpověď'}</p>` : ''}
            ${data.stats ? `<p>Statistiky: ${JSON.stringify(data.stats, null, 2)}</p>` : ''}
        </div>
    `;
}

// Zobrazení chybového výsledku zpracování serverem
function displayErrorResult(data) {
    // Sestavení HTML pro zobrazení chyby
    voiceResult.innerHTML = `
        <div class="voice-result-error">
            <p>Chyba: ${data.error || 'Neznámá chyba serveru'}</p>
            ${data.errors ? `
                <ul>
                    ${data.errors.map(error => `<li>${error}</li>`).join('')}
                </ul>
            ` : ''}
            ${data.original_text ? `<p>Původní text: "${data.original_text}"</p>` : ''}
        </div>
    `;
}

// Zpracování akce podle typu příkazu z rozpoznaného textu
function handleVoiceAction(data) {
    switch (data.entities.action) {
        case 'record_time': // Akce pro záznam pracovní doby nebo volného dne
            prefillWorkTimeForm(data.entities);
            break;
            
        case 'add_advance': // Akce pro přidání zálohy
            // TODO: Implementovat detekci akce 'add_advance' v VoiceProcessor._extract_entities a otestovat.
            prefillAdvanceForm(data.entities);
            break;
            
        case 'get_stats': // Akce pro zobrazení statistik
            redirectToStatistics(data.entities);
            break;
            
        default: // Neznámá akce
            voiceResult.textContent = 'Neznámá akce byla rozpoznána.';
    }
}

// Předvyplnění formuláře pracovní doby na základě rozpoznaných entit
function prefillWorkTimeForm(entities) {
    const form = document.getElementById('record-time-form'); // Získání formuláře
    if (!form) return; // Pokud formulář neexistuje, nic neděláme
    
    if (entities.date) {
        const dateInput = form.querySelector('input[name="date"]');
        if (dateInput) {
            dateInput.value = entities.date;
        }
    }

    const freeDayCheckbox = form.querySelector('input[name="is_free_day"]');
    if (freeDayCheckbox) {
        freeDayCheckbox.checked = entities.is_free_day || false;
        // Zavolat funkci pro skrytí/zobrazení časových polí
        if (typeof toggleTimeFields === 'function') {
            toggleTimeFields();
        }
    }
    
    if (!entities.is_free_day) {
        // Pouze pokud to není volný den, nastavíme časy
        if (entities.start_time) {
            const startTimeInput = form.querySelector('input[name="start_time"]');
            if (startTimeInput) {
                startTimeInput.value = entities.start_time;
            }
        }
        
        if (entities.end_time) {
            const endTimeInput = form.querySelector('input[name="end_time"]');
            if (endTimeInput) {
                endTimeInput.value = entities.end_time;
            }
        }

        const lunchInput = form.querySelector('input[name="lunch_duration"]');
        if (lunchInput) {
            lunchInput.value = entities.lunch_duration || "1.0";
        }
    }
    
    // Automatické přepnutí na sekci formuláře
    document.getElementById('record-time-section').scrollIntoView({ behavior: 'smooth' });
}

// Předvyplnění formuláře pro zadání zálohy na základě rozpoznaných entit
function prefillAdvanceForm(entities) {
    const form = document.getElementById('advance-form'); // Získání formuláře
    if (!form) return; // Pokud formulář neexistuje, nic neděláme
    
    // Předvyplnění jména zaměstnance, pokud bylo rozpoznáno
    if (entities.employee) {
        const employeeSelect = form.querySelector('select[name="employee_name"]'); // Opraven název selektoru
        if (employeeSelect) {
            // Pokusíme se najít option, jehož hodnota (jméno zaměstnance) se shoduje
            let found = false;
            for (let i = 0; i < employeeSelect.options.length; i++) {
                if (employeeSelect.options[i].value === entities.employee) {
                    employeeSelect.selectedIndex = i;
                    found = true;
                    break;
                }
            }
            if (!found) {
                console.warn(`Zaměstnanec '${entities.employee}' nenalezen v selectu pro zálohy.`);
            }
        }
    }
    
    if (entities.date) {
        const dateInput = form.querySelector('input[name="date"]');
        if (dateInput) {
            dateInput.value = entities.date;
        }
    }
    
    if (entities.amount) {
        const amountInput = form.querySelector('input[name="amount"]');
        if (amountInput) {
            amountInput.value = entities.amount;
        }
    }
    
    if (entities.currency) {
        const currencySelect = form.querySelector('select[name="currency"]');
        if (currencySelect) {
            currencySelect.value = entities.currency;
        }
    }
    
    // Automatické přepnutí na sekci formuláře
    document.getElementById('advance-section').scrollIntoView({ behavior: 'smooth' });
}

// Přesměrování na stránku statistik s případnými parametry
function redirectToStatistics(entities) {
    let url = '/monthly_report'; // Cílová URL pro statistiky (měsíční report)
    
    const params = new URLSearchParams(); // Použití URLSearchParams pro snadné přidávání parametrů
    
    // Přidání parametrů, pokud byly rozpoznány
    if (entities.employee) {
        params.append('employees', entities.employee); // Název parametru je 'employees' pro monthly_report
    }
    if (entities.time_period) {
        // Mapování rozpoznaného období na hodnoty očekávané backendem, pokud je to nutné
        // Prozatím předpokládáme, že backend akceptuje 'week', 'month', 'year' přímo,
        // nebo že se to zpracuje na backendu. Zde jen předáváme.
        // params.append('period_type', entities.time_period); // Příklad, pokud by backend očekával jiný název
        
        // Měsíční report aktuálně filtruje podle měsíce a roku, nikoli obecného 'time_period'
        // Pokud je 'time_period' např. 'month', můžeme zkusit nastavit aktuální měsíc/rok
        // Nebo to ponechat na backendu, aby zobrazil výchozí pohled pro daný typ statistiky.
        // Pro jednoduchost zde pouze logujeme a nepřidáváme jako parametr,
        // protože /monthly_report očekává konkrétní měsíc/rok.
        console.log("Rozpoznáno časové období pro statistiky:", entities.time_period, "ale /monthly_report vyžaduje měsíc/rok.");
    }
    
    // Datum pro statistiky - /monthly_report typicky vyžaduje měsíc a rok, ne konkrétní datum.
    // Pokud by bylo potřeba předat datum pro jiný typ statistik, zde by se to přidalo.
    if (entities.date) {
         console.log("Rozpoznáno datum pro statistiky:", entities.date, "ale /monthly_report vyžaduje měsíc/rok.");
        // params.append('specific_date', entities.date);
    }

    const queryString = params.toString();
    if (queryString) {
        url += `?${queryString}`;
    }
    
    console.log("Přesměrování na URL pro statistiky:", url);
    window.location.href = url; // Přesměrování prohlížeče
}

// Event listener pro klávesové zkratky
document.addEventListener('keydown', (event) => {
    // Ctrl + Shift + V pro spuštění hlasového vstupu
    if (event.ctrlKey && event.shiftKey && event.key === 'V') {
        event.preventDefault(); // Zabráníme výchozí akci prohlížeče pro tuto kombinaci
        if (!isListening && recognition) { // Zkontrolujeme, zda recognition existuje
            recognition.start();
        }
    }
});
