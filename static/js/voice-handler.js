// static/js/voice-handler.js
const voiceButton = document.getElementById('voice-button');
const voiceResult = document.getElementById('voice-result');
const loadingIndicator = document.getElementById('loading-indicator');

let recognition;
let isListening = false;

// Inicializace Web Speech API
try {
    window.SpeechRecognition = window.SpeechRecognition || webkitSpeechRecognition;
    recognition = new window.SpeechRecognition();
    
    // Nastavení parametrů
    recognition.continuous = false;
    recognition.interimResults = false;
    recognition.lang = 'cs-CZ';
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

// Funkce pro odeslání hlasového příkazu na server (textová verze)
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

// Funkce pro odeslání hlasového souboru na server (audio verze)
function sendVoiceAudioToServer(blob) {
    const formData = new FormData();
    formData.append('audio', blob, 'voice.webm');
    
    fetch('/process-audio', {
        method: 'POST',
        body: formData
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
        console.error('Chyba při odesílání hlasového souboru:', error);
        voiceResult.textContent = 'Nepodařilo se odeslat hlasový soubor na server.';
    });
}

// Získání hlasového záznamu (WebM)
function startVoiceRecording() {
    navigator.mediaDevices.getUserMedia({ audio: true })
        .then(stream => {
            const mediaRecorder = new MediaRecorder(stream);
            const audioChunks = [];
            
            mediaRecorder.ondataavailable = (event) => {
                audioChunks.push(event.data);
            };
            
            mediaRecorder.onstop = () => {
                const blob = new Blob(audioChunks, { type: 'audio/webm' });
                sendVoiceAudioToServer(blob);
            };
            
            mediaRecorder.start();
            setTimeout(() => mediaRecorder.stop(), 10000); // Max 10 sekund nahrávání
        })
        .catch(err => {
            console.error('Chyba při záznamu hlasu:', err);
            voiceResult.textContent = 'Nepodařilo se spustit záznam hlasu.';
        });
}

// Přepsání původního tlačítka pro novou funkci
voiceButton.addEventListener('click', () => {
    if (!isListening) {
        startVoiceRecording();  // Začneme záznam hlasu
    } else {
        recognition.stop();
    }
});
