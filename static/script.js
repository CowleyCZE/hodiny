document.addEventListener('DOMContentLoaded', (event) => {
    // Add any client-side JavaScript functionality here
    console.log('DOM fully loaded and parsed');
});

// Example: Function to calculate work hours
function calculateWorkHours() {
    const startTime = document.getElementById('start_time').value;
    const endTime = document.getElementById('end_time').value;
    const lunchDuration = parseFloat(document.getElementById('lunch_duration').value);

    if (startTime && endTime && !isNaN(lunchDuration)) {
        const start = new Date(`2000-01-01T${startTime}`);
        const end = new Date(`2000-01-01T${endTime}`);
        
        let diff = (end - start) / 1000 / 60 / 60; // Convert to hours
        diff -= lunchDuration;
        
        alert(`Odpracované hodiny: ${diff.toFixed(2)}`);
    }
}

// Add event listeners to relevant forms or buttons
const timeRecordForm = document.querySelector('form');
if (timeRecordForm) {
    timeRecordForm.addEventListener('submit', (e) => {
        e.preventDefault();
        calculateWorkHours();
        timeRecordForm.submit();
    });
}