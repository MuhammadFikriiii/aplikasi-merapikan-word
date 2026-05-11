document.addEventListener('DOMContentLoaded', () => {
    // DOM Elements
    const dropzone = document.getElementById('dropzone');
    const fileInput = document.getElementById('fileInput');
    const filePreview = document.getElementById('filePreview');
    const fileName = document.getElementById('fileName');
    const fileSize = document.getElementById('fileSize');
    const fileRemove = document.getElementById('fileRemove');
    const submitBtn = document.getElementById('submitBtn');
    
    const optionsToggle = document.getElementById('optionsToggle');
    const optionsSection = optionsToggle.closest('.options-section');
    
    // Sections
    const uploadSection = document.getElementById('uploadSection');
    const processingSection = document.getElementById('processingSection');
    const resultSection = document.getElementById('resultSection');
    const errorSection = document.getElementById('errorSection');
    
    let selectedFile = null;

    // --- Format Options Toggle ---
    optionsToggle.addEventListener('click', () => {
        optionsSection.classList.toggle('open');
    });

    // --- File Drag and Drop Logic ---
    dropzone.addEventListener('click', () => fileInput.click());

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
            handleFile(e.dataTransfer.files[0]);
        }
    });

    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length) {
            handleFile(e.target.files[0]);
        }
    });

    fileRemove.addEventListener('click', () => {
        selectedFile = null;
        fileInput.value = '';
        dropzone.style.display = 'block';
        filePreview.style.display = 'none';
        submitBtn.disabled = true;
    });

    function handleFile(file) {
        // Validate file
        if (!file.name.toLowerCase().endsWith('.docx')) {
            alert('Error: Hanya file .docx yang diperbolehkan!');
            return;
        }

        if (file.size > 50 * 1024 * 1024) {
            alert('Error: Ukuran file melebihi 50MB!');
            return;
        }

        selectedFile = file;
        
        // Update UI
        fileName.textContent = file.name;
        fileSize.textContent = formatBytes(file.size);
        
        dropzone.style.display = 'none';
        filePreview.style.display = 'flex';
        submitBtn.disabled = false;
        
        // Auto-close options to save space
        optionsSection.classList.remove('open');
    }

    function formatBytes(bytes, decimals = 2) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const dm = decimals < 0 ? 0 : decimals;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
    }

    // --- Form Submission ---
    submitBtn.addEventListener('click', async () => {
        if (!selectedFile) return;

        // Hide upload, show processing
        uploadSection.style.display = 'none';
        optionsSection.style.display = 'none';
        submitBtn.style.display = 'none';
        processingSection.style.display = 'block';
        
        // Start animation sequence
        simulateProgress();

        const formData = new FormData();
        formData.append('document', selectedFile);
        
        // Append options
        formData.append('fontFamily', document.getElementById('fontFamily').value);
        formData.append('fontSize', document.getElementById('fontSize').value);
        formData.append('lineSpacing', document.getElementById('lineSpacing').value);
        formData.append('marginTop', document.getElementById('marginTop').value);
        formData.append('marginBottom', document.getElementById('marginBottom').value);
        formData.append('marginLeft', document.getElementById('marginLeft').value);
        formData.append('marginRight', document.getElementById('marginRight').value);
        formData.append('removeImages', document.getElementById('removeImages').checked);
        formData.append('fixTables', document.getElementById('fixTables').checked);
        formData.append('fixHeadings', document.getElementById('fixHeadings').checked);

        try {
            const response = await fetch('/api/process', {
                method: 'POST',
                body: formData
            });

            const result = await response.json();

            if (!response.ok) {
                throw new Error(result.error || 'Terjadi kesalahan pada server');
            }

            // Success
            showResult(result);

        } catch (error) {
            console.error('Error:', error);
            showError(error.message);
        }
    });

    // --- Progress Animation ---
    let progressInterval;
    function simulateProgress() {
        const steps = [
            document.getElementById('step1'),
            document.getElementById('step2'),
            document.getElementById('step3'),
            document.getElementById('step4')
        ];
        
        let currentStep = 0;
        
        // Reset steps
        steps.forEach(step => {
            step.classList.remove('active', 'done');
        });
        
        if(steps[0]) steps[0].classList.add('active');

        progressInterval = setInterval(() => {
            if (currentStep < steps.length - 1) {
                steps[currentStep].classList.remove('active');
                steps[currentStep].classList.add('done');
                currentStep++;
                steps[currentStep].classList.add('active');
            }
        }, 1500); // Move to next step every 1.5s
    }

    // --- Show Results ---
    function showResult(data) {
        clearInterval(progressInterval);
        processingSection.style.display = 'none';
        resultSection.style.display = 'block';
        
        // Populate report
        if (data.report) {
            animateValue('reportHeadings', 0, data.report.headingsFixed || 0, 1000);
            animateValue('reportTables', 0, data.report.tablesFixed || 0, 1000);
            animateValue('reportImages', 0, data.report.imagesRemoved || 0, 1000);
            animateValue('reportFormatting', 0, data.report.formattingFixed || 0, 1000);
        }

        // Set download link
        document.getElementById('downloadBtn').href = data.downloadUrl;
        document.getElementById('downloadBtn').download = data.filename;
    }

    function showError(message) {
        clearInterval(progressInterval);
        processingSection.style.display = 'none';
        errorSection.style.display = 'block';
        document.getElementById('errorText').textContent = message;
    }

    // --- Counter Animation ---
    function animateValue(id, start, end, duration) {
        if (start === end) {
            document.getElementById(id).innerHTML = end;
            return;
        }
        let range = end - start;
        let current = start;
        let increment = end > start ? 1 : -1;
        // Make it faster for larger numbers
        let stepTime = Math.abs(Math.floor(duration / range));
        if (stepTime < 20) stepTime = 20; 
        
        let obj = document.getElementById(id);
        let timer = setInterval(function() {
            current += increment;
            obj.innerHTML = current;
            if (current == end) {
                clearInterval(timer);
            }
        }, stepTime);
    }

    // --- Reset ---
    document.getElementById('resetBtn').addEventListener('click', resetApp);
    document.getElementById('errorResetBtn').addEventListener('click', resetApp);

    function resetApp() {
        resultSection.style.display = 'none';
        errorSection.style.display = 'none';
        
        uploadSection.style.display = 'block';
        optionsSection.style.display = 'block';
        submitBtn.style.display = 'block';
        
        fileRemove.click(); // Clears file input and resets UI
    }

    // --- Add floating particles background ---
    function createParticles() {
        const particlesContainer = document.getElementById('particles');
        const particleCount = 20;
        
        for (let i = 0; i < particleCount; i++) {
            const particle = document.createElement('div');
            
            // Random properties
            const size = Math.random() * 4 + 2;
            const x = Math.random() * 100;
            const y = Math.random() * 100;
            const duration = Math.random() * 20 + 10;
            const delay = Math.random() * 5;
            const opacity = Math.random() * 0.5 + 0.1;
            
            particle.style.cssText = `
                position: absolute;
                width: ${size}px;
                height: ${size}px;
                background: white;
                border-radius: 50%;
                top: ${y}%;
                left: ${x}%;
                opacity: ${opacity};
                pointer-events: none;
                animation: float-particle ${duration}s linear infinite;
                animation-delay: ${delay}s;
                z-index: -1;
            `;
            
            particlesContainer.appendChild(particle);
        }
        
        // Add animation style dynamically
        const style = document.createElement('style');
        style.innerHTML = `
            @keyframes float-particle {
                0% { transform: translateY(0) translateX(0); }
                33% { transform: translateY(-30px) translateX(20px); opacity: 0; }
                66% { transform: translateY(30px) translateX(-20px); }
                100% { transform: translateY(0) translateX(0); }
            }
        `;
        document.head.appendChild(style);
    }
    
    createParticles();
});
