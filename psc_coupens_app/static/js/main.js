// main.js - Base scripts for Praktutik SparshCare
document.addEventListener('DOMContentLoaded', () => {
    console.log('Prakrutik SparshCare UI Initialized');
    
    // Auto-dismiss alerts after 5 seconds
    const alerts = document.querySelectorAll('.alert');
    alerts.forEach(alert => {
        setTimeout(() => {
            const bsAlert = new bootstrap.Alert(alert);
            bsAlert.close();
        }, 5000);
    });
});
