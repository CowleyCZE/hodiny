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
    recognition.continuous = false;
    recognition.interimResults = false;
    recognition.lang = 'cs-CZ'; // Čeština
    recognition.maxAlternatives = 1;
    
    // Event: Začátek naslouchání
    recognition.onstart = () => {
        isListening = true;
        voiceButton.classList.add('active');
        voiceResult.textContent = 'Naslouchám...';
        loadingIndicator.style.display = 'inline-block';
    };
    
    // Event: Konec naslouchání
    recognition.onend = () => {
        isListening = false;
        voiceButton.classList.remove('active');
        loadingIndicator.style.display = 'none';
    };
    
    // Event: Výsledky rozpoznání
    recognition.onresult = (event) => {
        const transcript = event.results[0][0].transcript.trim();
        voiceResult.textContent = `Rozpoznáno: ${transcript}`;
        
        // Odeslání výsledku na server
        sendVoiceCommandToServer(transcript);
    };
    
    // Event: Chyba
    recognition.onerror = (event) => {
        console.error('Chyba při hlasovém rozpoznávání:', event.error);
        voiceResult.textContent = `Chyba: ${event.error}`;
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

// Funkce pro odeslání hlasového příkazu na server
function sendVoiceCommandToServer(text) {
    fetch('/voice-command', {
        method: 'POST',
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
            displaySuccessResult(data);
            handleVoiceAction(data);
        } else {
            displayErrorResult(data);
        }
    })
    .catch(error => {
        console.error('Chyba při odesílání hlasového příkazu:', error);
        voiceResult.textContent = 'Nepodařilo se odeslat hlasový příkaz na server.';
    });
}

// Zobrazení úspěšného výsledku
function displaySuccessResult(data) {
    const confidence = (data.confidence * 100).toFixed(1);
    voiceResult.innerHTML = `
        <div class="voice-result-success">
            <p>Úspěšně rozpoznáno (spolehlivost: ${confidence}%)</p>
            <ul>
                ${Object.entries(data.entities).map(([key, value]) => 
                    `<li><strong>${key}:</strong> ${value}</li>`).join('')}
            </ul>
        </div>
    `;
}

// Zobrazení chybového výsledku
function displayErrorResult(data) {
    voiceResult.innerHTML = `
        <div class="voice-result-error">
            <p>${data.error || 'Neznámá chyba'}</p>
            ${data.errors ? `
                <ul>
                    ${data.errors.map(error => `<li>${error}</li>`).join('')}
                </ul>
            ` : ''}
        </div>
    `;
}

// Zpracování akce podle typu příkazu
function handleVoiceAction(data) {
    switch (data.entities.action) {
        case 'record_time':
            prefillWorkTimeForm(data.entities);
            break;
            
        case 'add_advance':
            prefillAdvanceForm(data.entities);
            break;
            
        case 'get_stats':
            redirectToStatistics(data.entities);
            break;
            
        default:
            voiceResult.textContent = 'Neznámá akce byla rozpoznána.';
    }
}

// Předvyplnění formuláře pracovní doby
function prefillWorkTimeForm(entities) {
    const form = document.getElementById('record-time-form');
    if (!form) return;
    
    if (entities.date) {
        const dateInput = form.querySelector('input[name="date"]');
        if (dateInput) {
            dateInput.value = entities.date;
        }
    }
    
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

    // Nastavení délky oběda s přednastavenou hodnotou 1 hodina
    const lunchInput = form.querySelector('input[name="lunch_duration"]');
    if (lunchInput) {
        lunchInput.value = entities.lunch_duration || "1.0";
    }
    
    // Automatické přepnutí na sekci formuláře
    document.getElementById('record-time-section').scrollIntoView({ behavior: 'smooth' });
}

// Předvyplnění formuláře zálohy
function prefillAdvanceForm(entities) {
    const form = document.getElementById('advance-form');
    if (!form) return;
    
    if (entities.employee) {
        const employeeSelect = form.querySelector('select[name="employee"]');
        if (employeeSelect) {
            employeeSelect.value = entities.employee;
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

// Přesměrování na statistiky
function redirectToStatistics(entities) {
    let url = '/statistics';
    
    if (entities.employee) {
        url += `?employee=${encodeURIComponent(entities.employee)}`;
    }
    
    if (entities.time_period) {
        url += `${url.includes('?') ? '&' : '?'}period=${encodeURIComponent(entities.time_period)}`;
    }
    
    if (entities.date) {
        url += `${url.includes('?') ? '&' : '?'}date=${encodeURIComponent(entities.date)}`;
    }
    
    window.location.href = url;
}

// Funkce pro získání aktuálního času
function getCurrentTime() {
    const now = new Date();
    return now.toTimeString().split(' ')[0];
}

// Funkce pro získání dnešního data
function getTodayDate() {
    const today = new Date();
    return today.toISOString().split('T')[0];
}

// Funkce pro získání včerejšího data
function getYesterdayDate() {
    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    return yesterday.toISOString().split('T')[0];
}

// Event listener pro klávesové zkratky
document.addEventListener('keydown', (event) => {
    // Ctrl + Shift + V pro spuštění hlasového vstupu
    if (event.ctrlKey && event.shiftKey && event.key === 'V') {
        event.preventDefault();
        if (!isListening) {
            recognition.start();
        }
    }
});
