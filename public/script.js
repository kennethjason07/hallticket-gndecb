document.addEventListener('DOMContentLoaded', () => {
    const form = document.getElementById('generatorForm');
    const excelInput = document.getElementById('excelFile');
    const subjectRadios = document.getElementsByName('subjectMode');
    const manualSection = document.getElementById('manualSubjectsSection');
    const subjectsList = document.getElementById('subjectsList');
    const addSubjectBtn = document.getElementById('addSubjectBtn');
    const generateBtn = document.getElementById('generateBtn');
    const btnText = generateBtn.querySelector('.btn-text');
    const loader = generateBtn.querySelector('.loader');

    // File Input UI
    function updateFileUpload(input, containerId) {
        const container = document.getElementById(containerId);
        const textSpan = container.querySelector('.text');

        input.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                container.classList.add('has-file');
                textSpan.textContent = e.target.files[0].name;
            } else {
                container.classList.remove('has-file');
                textSpan.textContent = 'Drop Excel file here or click to browse';
            }
        });
    }

    updateFileUpload(excelInput, 'excelUpload');

    // Subject Mode Toggle
    subjectRadios.forEach(radio => {
        radio.addEventListener('change', (e) => {
            if (e.target.value === 'manual') {
                manualSection.classList.remove('hidden');
                if (subjectsList.children.length === 0) {
                    addSubject(); // Add one default row
                }
            } else {
                manualSection.classList.add('hidden');
            }
        });
    });

    // Dynamic Subjects
    function addSubject() {
        const div = document.createElement('div');
        div.className = 'subject-row';
        div.innerHTML = `
            <input type="text" placeholder="Subject Code - Name" class="subject-input">
            <button type="button" class="remove-btn" onclick="this.parentElement.remove()">×</button>
        `;
        subjectsList.appendChild(div);
    }

    addSubjectBtn.addEventListener('click', addSubject);

    // Form Submission
    form.addEventListener('submit', async (e) => {
        e.preventDefault();

        // Validation
        if (!excelInput.files[0]) {
            alert('Please select an Excel file');
            return;
        }

        // Loading State
        generateBtn.disabled = true;
        btnText.textContent = 'Generating...';
        loader.classList.remove('hidden');

        const formData = new FormData();
        formData.append('excelFile', excelInput.files[0]);
        formData.append('deptName', document.getElementById('deptName').value);
        formData.append('examName', document.getElementById('examName').value);
        formData.append('semester', document.getElementById('semester').value);

        const useManual = document.querySelector('input[name="subjectMode"]:checked').value === 'manual';
        formData.append('useManualSubjects', useManual);

        if (useManual) {
            const inputs = document.querySelectorAll('.subject-input');
            const subjects = Array.from(inputs)
                .map(input => input.value.trim())
                .filter(val => val.length > 0);
            formData.append('customSubjects', JSON.stringify(subjects));
        }

        try {
            const response = await fetch('/generate', {
                method: 'POST',
                body: formData
            });

            if (!response.ok) {
                throw new Error(await response.text());
            }

            // Handle Download
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'halltickets.zip';
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            a.remove();

        } catch (error) {
            alert('Error: ' + error.message);
        } finally {
            // Reset State
            generateBtn.disabled = false;
            btnText.textContent = 'Generate Hall Tickets';
            loader.classList.add('hidden');
        }
    });
    // Help Modal Logic
    const helpBtn = document.getElementById('helpBtn');
    const helpModal = document.getElementById('helpModal');
    const closeModal = document.querySelector('.close-modal');

    helpBtn.addEventListener('click', () => {
        helpModal.classList.remove('hidden');
    });

    closeModal.addEventListener('click', () => {
        helpModal.classList.add('hidden');
    });

    window.addEventListener('click', (e) => {
        if (e.target === helpModal) {
            helpModal.classList.add('hidden');
        }
    });
});
