/**
 * Общие JavaScript функции для SERP Comparator
 */

// === МОДАЛЬНЫЕ ОКНА ===
function openModal(img) {
    document.getElementById('modalChartImage').src = img.src;
    new bootstrap.Modal(document.getElementById('chartModal')).show();
}

function closeModal() {
    bootstrap.Modal.getInstance(document.getElementById('chartModal')).hide();
}

// === ТЕМА ===
function setTheme(isDark) {
    const body = document.body;
    const themeIconSun = document.getElementById('themeIconSun');
    const themeIconMoon = document.getElementById('themeIconMoon');
    const themeTextFull = document.querySelector('.full-text');
    const themeTextShort = document.querySelector('.short-text');
    const themeText = document.getElementById('themeText');

    if (isDark) {
        body.classList.add('dark-theme');
        if (themeIconSun) themeIconSun.style.display = 'inline-block';
        if (themeIconMoon) themeIconMoon.style.display = 'none';
        if (themeTextFull) themeTextFull.textContent = 'Светлая тема';
        if (themeTextShort) themeTextShort.textContent = 'Светлая';
        if (themeText) themeText.textContent = 'Светлая тема';
        localStorage.setItem('theme', 'dark');
    } else {
        body.classList.remove('dark-theme');
        if (themeIconSun) themeIconSun.style.display = 'none';
        if (themeIconMoon) themeIconMoon.style.display = 'inline-block';
        if (themeTextFull) themeTextFull.textContent = 'Тёмная тема';
        if (themeTextShort) themeTextShort.textContent = 'Тема';
        if (themeText) themeText.textContent = 'Тёмная тема';
        localStorage.setItem('theme', 'light');
    }
}

function initTheme() {
    const sidebarToggle = document.getElementById('sidebarToggle');
    const sidebar = document.getElementById('sidebar');
    const mainContent = document.querySelector('.main-content');

    if (sidebarToggle && sidebar) {
        sidebarToggle.addEventListener('click', function() {
            sidebar.classList.toggle('collapsed');
            
            // Adjust main content margin based on sidebar state
            if (sidebar.classList.contains('collapsed')) {
                mainContent.style.marginLeft = '60px';
            } else {
                mainContent.style.marginLeft = '250px';
            }
        });
    }

    const themeToggle = document.getElementById('themeToggle');
    if (!themeToggle) return;

    const savedTheme = localStorage.getItem('theme');
    if (savedTheme === 'dark') {
        setTheme(true);
    } else {
        setTheme(false);
    }

    themeToggle.addEventListener('click', () => {
        setTheme(!document.body.classList.contains('dark-theme'));
    });
}

// === TOAST УВЕДОМЛЕНИЯ ===
function showToast(message, type = 'info') {
    const container = document.getElementById('toast-container');
    if (!container) return;

    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    toast.innerHTML = `
        <div class="toast-message">${message}</div>
        <button class="toast-close" onclick="this.parentElement.remove()">&times;</button>
    `;

    container.appendChild(toast);

    setTimeout(() => {
        toast.classList.add('show');
    }, 10);

    setTimeout(() => {
        toast.classList.remove('show');
        setTimeout(() => toast.remove(), 300);
    }, 3000);
}

// === BASELINE МЕТРИКИ ===
function toggleTotalUrls(prefix) {
    const inputType = document.getElementById(prefix + '_input_type');
    const container = document.getElementById(prefix + '_total_urls_container');

    if (!inputType || !container) return;

    if (inputType.value === 'count') {
        container.style.display = 'block';
    } else {
        container.style.display = 'none';
    }
}

// === ИНИЦИАЛИЗАЦИЯ ===
document.addEventListener('DOMContentLoaded', () => {
    initTheme();
});
