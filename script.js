// Initialize Lucide Icons
lucide.createIcons();

const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('fileInput');
const dashboard = document.getElementById('dashboard');
const uploadSection = document.querySelector('.upload-section');

const activeCountEl = document.getElementById('activeCount');
const inactiveCountEl = document.getElementById('inactiveCount');
const duplicateCountEl = document.getElementById('duplicateCount');
const studentsTableBody = document.querySelector('#studentsTable tbody');
const duplicatesTableBody = document.querySelector('#duplicatesTable tbody');

// Drag and drop events
dropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropzone.classList.add('dragover');
});

dropzone.addEventListener('dragleave', () => {
    dropzone.classList.remove('dragover');
});

dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropzone.classList.remove('dragover');
    if (e.dataTransfer.files.length) {
        fileInput.files = e.dataTransfer.files;
        handleFile(e.dataTransfer.files[0]);
    }
});

// File click event
fileInput.addEventListener('change', (e) => {
    if (e.target.files.length) {
        handleFile(e.target.files[0]);
    }
});

function handleFile(file) {
    if (!file) return;

    // Show a small loading state
    dropzone.querySelector('h2').innerText = "Processing File...";

    const reader = new FileReader();

    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            // Assume first sheet
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            // Convert to JSON
            const json = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
            processData(json);
        } catch (error) {
            console.error("Error parsing file:", error);
            alert("Error reading the Excel file. Is it valid?");
            dropzone.querySelector('h2').innerText = "Upload Student List (Excel)";
        }
    };

    reader.readAsArrayBuffer(file);
}

function processData(data) {
    if (!data || data.length === 0) {
        alert("The Excel file is empty or formatted incorrectly.");
        dropzone.querySelector('h2').innerText = "Upload Student List (Excel)";
        return;
    }

    let activeCount = 0;
    let deactiveCount = 0;
    let students = [];
    
    // Guess columns based on common variants
    const firstRow = data[0];
    const keys = Object.keys(firstRow);
    const nameKey = keys.find(k => k.toLowerCase().includes('name')) || keys[0] || 'Name';
    const statusKey = keys.find(k => k.toLowerCase().includes('status') || k.toLowerCase().includes('active')) || keys[1] || 'Status';
    const idKey = keys.find(k => k.toLowerCase().includes('id') || k.toLowerCase().includes('roll')) || keys[2] || 'ID';

    data.forEach(row => {
        const name = (row[nameKey] || 'Unknown').toString().trim();
        const rawStatus = (row[statusKey] || '').toString().toLowerCase().trim();
        
        // Skip empty rows
        if (!name && !rawStatus) return;

        let isActive = true;
        
        // Define deactive flags
        const deactiveFlags = ['deactive', 'inactive', 'no', '0', 'false', 'left', 'suspended'];
        if (deactiveFlags.some(flag => rawStatus.includes(flag))) {
            isActive = false;
        }

        if (isActive) {
            activeCount++;
        } else {
            deactiveCount++;
        }

        students.push({
            name: name,
            status: isActive ? 'Active' : 'Deactive',
            id: row[idKey] || '-'
        });
    });

    // Bifurcate Duplicates Logic
    const nameCounts = {};
    students.forEach(s => {
        const lowerName = s.name.toLowerCase();
        nameCounts[lowerName] = (nameCounts[lowerName] || 0) + 1;
    });

    const duplicateNames = [];
    for (const [name, count] of Object.entries(nameCounts)) {
        if (count > 1) {
            duplicateNames.push({ name: capitalize(name), count });
        }
    }

    // Update UI Stats
    animateCounter(activeCountEl, activeCount);
    animateCounter(inactiveCountEl, deactiveCount);
    animateCounter(duplicateCountEl, duplicateNames.length);

    // Render Tables
    renderStudentsTable(students);
    renderDuplicatesTable(duplicateNames);
    
    // Switch View Component
    uploadSection.style.display = 'none';
    dashboard.classList.remove('hidden');
    lucide.createIcons(); // Re-initialize icons for new DOM elements
}

// Visual counter animation
function animateCounter(el, target) {
    let current = 0;
    const increment = Math.max(1, Math.floor(target / 40));
    const timer = setInterval(() => {
        if (current >= target) {
            el.innerText = target;
            clearInterval(timer);
        } else {
            current += increment;
            if (current > target) current = target;
            el.innerText = current;
        }
    }, 20);
}

function capitalize(str) {
    return str.split(' ').map(word => word.charAt(0).toUpperCase() + word.slice(1)).join(' ');
}

function renderStudentsTable(students) {
    studentsTableBody.innerHTML = '';
    // Limit to exactly the first 200 items for DOM performance in standard lists
    students.slice(0, 200).forEach(student => { 
        const tr = document.createElement('tr');
        const statusClass = student.status === 'Active' ? 'status-active' : 'status-deactive';
        
        tr.innerHTML = `
            <td><strong>${student.name}</strong></td>
            <td><span class="status-badge ${statusClass}">${student.status}</span></td>
            <td>${student.id}</td>
        `;
        studentsTableBody.appendChild(tr);
    });
    
    if (students.length > 200) {
        const tr = document.createElement('tr');
        tr.innerHTML = `<td colspan="3" style="text-align:center; color: var(--text-secondary)">...and ${students.length - 200} more students</td>`;
        studentsTableBody.appendChild(tr);
    }
}

function renderDuplicatesTable(duplicates) {
    duplicatesTableBody.innerHTML = '';
    
    if (duplicates.length === 0) {
        duplicatesTableBody.innerHTML = '<tr><td colspan="2" style="text-align: center; color: var(--text-secondary); padding: 3rem;">No duplicate names bifurcated.</td></tr>';
        return;
    }

    // Sort by most occurrences
    duplicates.sort((a,b) => b.count - a.count).forEach(dup => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${dup.name}</td>
            <td><span style="font-weight: 700; color: var(--duplicate-color); padding: 0.25rem 0.75rem; background: #fef3c7; border-radius: 999px;">${dup.count} Records found</span></td>
        `;
        duplicatesTableBody.appendChild(tr);
    });
}
