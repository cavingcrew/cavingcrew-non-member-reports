function refreshCredentials() {
    const spreadsheetId = '1MVTO45ZusIw2BRgtRh4BffPRbbuIZHViZKPZE3OntEc'; // Caving Crew credentials store
    const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName('Credentials');
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const labelIndex = headers.indexOf('label');
    const valueIndex = headers.indexOf('value');

    if (labelIndex === -1 || valueIndex === -1) {
        throw new Error('Required columns "label" and "value" not found in the Credentials sheet.');
    }

    data.forEach(row => {
        const label = row[labelIndex];
        const value = row[valueIndex];
        if (label && value) {
            PropertiesService.getScriptProperties().setProperty(`cred_${label}`, value);
        }
    });
}

function createNightlyTrigger() {
    // Delete existing triggers
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
        if (trigger.getHandlerFunction() === 'refreshCredentials') {
            ScriptApp.deleteTrigger(trigger);
        }
    });

    // Create a new trigger to run at 1:00 AM every day
    ScriptApp.newTrigger('refreshCredentials')
        .timeBased()
        .everyDays(1)
        .atHour(1)
        .create();
}

// Run this function once to set up the nightly trigger
function setupCredentialRefresh() {
    refreshCredentials();
    createNightlyTrigger();
}
