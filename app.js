const GROUP_KEY = 'selected_group';
let scheduleData = null;

// Расписание звонков
const bellSchedule = {
    monday: [
        { name: 'Классный час', time: '08:45 – 09:15' },
        { name: '1 пара', time: '09:20 – 10:40' },
        { name: '2 пара', time: '10:50 – 12:10' },
        { name: '3 пара', time: '12:25 – 13:45' },
        { name: '4 пара', time: '13:50 – 15:10' },
        { name: '5 пара', time: '15:15 – 16:35' },
        { name: '6 пара', time: '16:40 – 18:00' }
    ],
    other: [
        { name: '1 пара', time: '08:45 – 10:05' },
        { name: '2 пара', time: '10:15 – 11:35' },
        { name: '3 пара', time: '11:45 – 13:05' },
        { name: '4 пара', time: '13:20 – 14:40' },
        { name: '5 пара', time: '14:45 – 16:05' },
        { name: '6 пара', time: '16:10 – 17:30' }
    ]
};

// Загрузка данных
async function init() {
    const loader = document.getElementById('loader');
    const container = document.getElementById('schedule-container');
    
    try {
        loader.classList.remove('hidden');
        const resp = await fetch('schedule.json?t=' + Date.now()); // cache busting
        if (!resp.ok) throw new Error('Файл не найден');
        scheduleData = await resp.json();
        
        document.getElementById('date-display').textContent = scheduleData.date || 'Расписание';
        populateGroups();
        restoreSelection();
        initBellSchedule();
    } catch (e) {
        container.innerHTML = '<div class="placeholder">❌ Не удалось загрузить расписание.<br>Проверьте файл schedule.json</div>';
        console.error(e);
    } finally {
        loader.classList.add('hidden');
    }
}

function populateGroups() {
    const select = document.getElementById('group-select');
    const groups = scheduleData.groups || [];
    groups.forEach(g => {
        const opt = document.createElement('option');
        opt.value = g;
        opt.textContent = g;
        select.appendChild(opt);
    });
    
    select.addEventListener('change', (e) => {
        const group = e.target.value;
        localStorage.setItem(GROUP_KEY, group);
        renderSchedule(group);
    });
}

function restoreSelection() {
    const saved = localStorage.getItem(GROUP_KEY);
    const select = document.getElementById('group-select');
    if (saved && scheduleData.groups.includes(saved)) {
        select.value = saved;
        renderSchedule(saved);
    }
}

function renderSchedule(group) {
    const container = document.getElementById('schedule-container');
    
    if (!group) {
        container.innerHTML = '<div class="placeholder">👆 Выберите группу выше</div>';
        return;
    }
    
    const lessons = scheduleData.schedule?.[group];
    
    if (!lessons || lessons.length === 0) {
        container.innerHTML = '<div class="placeholder">На сегодня занятий нет 🎉</div>';
        return;
    }
    
    container.innerHTML = lessons.map(l => `
        <div class="lesson-card ${l.subject ? '' : 'empty'}">
            <div class="lesson-header">
                <span class="lesson-num">${l.num}-я пара</span>
                <span class="lesson-subject">${l.subject || 'Окно'}</span>
            </div>
            ${l.teacher || l.room ? `<div class="lesson-meta">
                ${l.teacher ? `<span>👨‍🏫 ${l.teacher}</span>` : ''}
                ${l.room ? `<span>🚪 Каб. ${l.room}</span>` : ''}
            </div>` : ''}
        </div>
    `).join('');
}

document.addEventListener('DOMContentLoaded', init);

// Инициализация расписания звонков
function initBellSchedule() {
    const modal = document.getElementById('bell-modal');
    const openBtn = document.getElementById('bell-schedule-btn');
    const closeBtn = document.getElementById('close-modal');
    const tabBtns = document.querySelectorAll('.tab-btn');
    
    // Открытие модального окна
    openBtn.addEventListener('click', () => {
        modal.classList.remove('hidden');
        renderBellSchedule('monday');
    });
    
    // Закрытие модального окна
    closeBtn.addEventListener('click', () => {
        modal.classList.add('hidden');
    });
    
    // Закрытие по клику вне контента
    modal.addEventListener('click', (e) => {
        if (e.target === modal) {
            modal.classList.add('hidden');
        }
    });
    
    // Переключение табов
    tabBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            tabBtns.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            renderBellSchedule(btn.dataset.day);
        });
    });
}

// Рендеринг расписания звонков
function renderBellSchedule(day) {
    const container = document.getElementById('bell-content');
    const schedule = bellSchedule[day];
    
    let html = '<table class="bell-table"><thead><tr><th>Занятие</th><th>Время</th></tr></thead><tbody>';
    schedule.forEach(item => {
        html += `<tr><td>${item.name}</td><td class="bell-time">${item.time}</td></tr>`;
    });
    html += '</tbody></table>';
    
    container.innerHTML = html;
}