// Clean content (similar to clean_content in Python)
function cleanContent(content) {
    // Remove date/sender lines (e.g., [12/31, 12:00] Name:)
    content = content.replace(/\[(\d{1,2}\/\d{1,2},\s*\d{1,2}:\d{2})\]\s*[^:]+:/g, '');
    // Remove WhatsApp forwarded indicators
    content = content.replace(/-\s*Forwarded\s*message\s*-/gi, '');
    // Remove section headers (PTFE, PFA, SR#3, etc.)
    content = content.replace(/^\s*(PTFE|PFA|SR#\d+)\s*$/gim, '');
    // Remove bullet points and special characters at start of lines
    content = content.replace(/^\s*[°•\-]\s*/gm, '');
    return content;
}

// Extract info (similar to extract_info in Python)
function extractInfo(entry) {
    entry = entry.trim();
    if (!entry) return { equipment: '', description: '', technician: '' };

    // Handle M/s Arti pattern
    if (entry.includes('M/s Arti')) {
        let parts = entry.split('M/s Arti');
        let description = parts[0].trim();
        return { equipment: '', description: description, technician: 'M/s Arti' };
    }

    // Extract equipment (starts with capital letter, may contain digits, hyphens, spaces, etc.)
    let equipment = '';
    let equipmentMatch = entry.match(/^([A-Z][\w\s\-#\/]+?)[,.]/);
    if (equipmentMatch) {
        equipment = equipmentMatch[1].trim();
        entry = entry.replace(equipment, '').replace(/^[ ,.-]+/, '');
    }

    // Handle names with prefixes
    let description = entry;
    let technician = '';
    if (entry.includes(',')) {
        let lastComma = entry.lastIndexOf(',');
        let nameCandidate = entry.slice(lastComma + 1).trim();
        if (nameCandidate.match(/^[A-Z][a-z]+(\s+[A-Z][a-z]+)*$/)) {
            description = entry.slice(0, lastComma).trim();
            technician = nameCandidate;
            return { equipment, description, technician };
        }
    }

    // Ends with parentheses
    if (entry.includes('(') && entry.includes(')')) {
        let lastOpen = entry.lastIndexOf('(');
        let lastClose = entry.lastIndexOf(')');
        if (lastOpen < lastClose) {
            technician = entry.slice(lastOpen + 1, lastClose).trim();
            description = entry.slice(0, lastOpen).trim();
            return { equipment, description, technician };
        }
    }

    // Ends with a name
    let nameMatch = entry.match(/(\b[A-Z][a-z]+)\s*$/);
    if (nameMatch) {
        technician = nameMatch[1].trim();
        description = entry.slice(0, nameMatch.index).trim();
        return { equipment, description, technician };
    }

    return { equipment, description, technician: '' };
}

// Generate spreadsheet
function generateSpreadsheet() {
    const input = document.getElementById('inputMessages').value;
    const cleanedContent = cleanContent(input);

    // Split entries (similar to Python logic)
    const entries = [];
    let currentEntry = [];
    const lines = cleanedContent.split('\n').map(line => line.trim()).filter(line => line);

    for (let line of lines) {
        currentEntry.push(line);
        // Check if line looks like an entry end
        const isEnd = (
            (line.includes('(') && line.includes(')') && line.lastIndexOf(')') === line.length - 1) ||
            line.match(/[A-Z][a-z]+\s*$/) ||
            line.endsWith('.') ||
            line.endsWith(')')
        );
        if (isEnd) {
            entries.push(currentEntry.join(' '));
            currentEntry = [];
        }
    }
    if (currentEntry.length) entries.push(currentEntry.join(' '));

    // Process entries and create data for spreadsheet
    const data = [['Equipment', 'Description', 'Technician']]; // Headers
    for (let entry of entries) {
        const { equipment, description, technician } = extractInfo(entry);
        if (description) data.push([equipment, description, technician]);
    }

    // Create workbook with SheetJS
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(data);
    
    // Auto-size columns (approximation)
    const colWidths = data[0].map((_, colIndex) => {
        return Math.max(...data.map(row => (row[colIndex] || '').length)) * 1.2;
    });
    ws['!cols'] = colWidths.map(w => ({ wch: w }));

    XLSX.utils.book_append_sheet(wb, ws, 'Work Reports');

    // Download the file
    XLSX.write(wb, 'work_reports.xlsx');
    document.getElementById('output').innerText = `Processed ${entries.length} entries. Spreadsheet downloaded.`;
}