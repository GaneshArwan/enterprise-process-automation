// Helper function to create a styled paragraph
function createStyledParagraph(text) {
    return `<p style="margin: 5px 0;">${text}</p>`;
}

// Helper function to create table rows
function createTableRow(cells, isHeader = false) {
    const cellTag = isHeader ? 'th' : 'td';
    return `<tr>${cells.map(cell =>
        `<${cellTag}>${cell}</${cellTag}>`
    ).join('')}</tr>`;
}

function createTableHTML(values) {
    if (!values || values.length === 0) return '<table></table>';

    // Get all unique headers from the keys of the first object
    const headers = Object.keys(values[0]);

    // Start table HTML
    let tableHTML = `<table border="1" cellpadding="5" cellspacing="0">`;

    // Add header row
    tableHTML += createTableRow(headers, true);

    // Add data rows
    for (const row of values) {
        const cells = headers.map(header => row[header] ?? ''); // Fill missing values with empty string
        tableHTML += createTableRow(cells);
    }

    tableHTML += '</table>';
    return tableHTML;
}

function validateEmails(emailStr) {
    if (!emailStr || typeof emailStr !== 'string') {
        return { valid: [], invalid: [] };
    }

    const emails = emailStr.split(',')
        .map(email => email.trim())
        .filter(email => email !== '');

    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

    const validEmails = [];
    const invalidEmails = [];

    emails.forEach(email => {
        if (email) {
            emailRegex.test(email) 
                ? validEmails.push(email)
                : invalidEmails.push(email)
        }
    });

    return {
        valid: validEmails,
        invalid: invalidEmails
    };
}
