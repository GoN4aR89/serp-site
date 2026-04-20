/**
 * Landing Page JavaScript
 * Demo comparison without registration
 */

// File upload handling
document.addEventListener('DOMContentLoaded', function() {
    const file1Input = document.getElementById('file1');
    const file2Input = document.getElementById('file2');
    const file1Name = document.getElementById('file1-name');
    const file2Name = document.getElementById('file2-name');
    const demoForm = document.getElementById('demo-form');
    const progressWrapper = document.getElementById('demo-progress');
    const progressBar = progressWrapper.querySelector('.progress-bar');
    const progressText = progressWrapper.querySelector('.progress-text');

    // Handle file selection
    function handleFileSelect(input, nameElement, dropZone) {
        input.addEventListener('change', function() {
            if (this.files && this.files[0]) {
                nameElement.textContent = this.files[0].name;
                dropZone.classList.add('has-file');
                dropZone.style.borderColor = 'var(--primary)';
            }
        });
    }

    handleFileSelect(file1Input, file1Name, document.getElementById('drop-zone-1'));
    handleFileSelect(file2Input, file2Name, document.getElementById('drop-zone-2'));

    // Drag and drop
    function handleDragDrop(dropZone, input) {
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            dropZone.addEventListener(eventName, preventDefaults, false);
        });

        function preventDefaults(e) {
            e.preventDefault();
            e.stopPropagation();
        }

        ['dragenter', 'dragover'].forEach(eventName => {
            dropZone.addEventListener(eventName, () => {
                dropZone.classList.add('dragover');
            }, false);
        });

        ['dragleave', 'drop'].forEach(eventName => {
            dropZone.addEventListener(eventName, () => {
                dropZone.classList.remove('dragover');
            }, false);
        });

        dropZone.addEventListener('drop', (e) => {
            const files = e.dataTransfer.files;
            if (files.length > 0) {
                input.files = files;
                input.dispatchEvent(new Event('change'));
            }
        }, false);
    }

    handleDragDrop(document.getElementById('drop-zone-1'), file1Input);
    handleDragDrop(document.getElementById('drop-zone-2'), file2Input);

    // Form submission
    demoForm.addEventListener('submit', function(e) {
        e.preventDefault();

        if (!file1Input.files[0] || !file2Input.files[0]) {
            alert('Пожалуйста, загрузите оба файла');
            return;
        }

        const formData = new FormData();
        formData.append('file1', file1Input.files[0]);
        formData.append('file2', file2Input.files[0]);

        // Show progress
        progressWrapper.style.display = 'block';
        progressBar.style.width = '0%';
        progressText.textContent = 'Загрузка файлов...';

        // Simulate progress
        let progress = 0;
        const progressInterval = setInterval(() => {
            progress += 10;
            if (progress <= 90) {
                progressBar.style.width = progress + '%';
                if (progress === 30) {
                    progressText.textContent = 'Анализ данных...';
                } else if (progress === 60) {
                    progressText.textContent = 'Построение диаграмм...';
                }
            }
        }, 200);

        fetch('/demo-compare', {
            method: 'POST',
            body: formData
        })
        .then(response => response.json())
        .then(data => {
            clearInterval(progressInterval);
            progressBar.style.width = '100%';
            progressText.textContent = 'Готово!';

            if (data.success) {
                window.location.href = data.redirect;
            } else {
                alert(data.error || 'Ошибка при сравнении');
                progressWrapper.style.display = 'none';
            }
        })
        .catch(error => {
            clearInterval(progressInterval);
            console.error('Error:', error);
            alert('Произошла ошибка при загрузке');
            progressWrapper.style.display = 'none';
        });
    });
});

// Smooth scroll for anchor links
document.querySelectorAll('a[href^="#"]').forEach(anchor => {
    anchor.addEventListener('click', function (e) {
        e.preventDefault();
        const target = document.querySelector(this.getAttribute('href'));
        if (target) {
            target.scrollIntoView({
                behavior: 'smooth',
                block: 'start'
            });
        }
    });
});
