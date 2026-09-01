const GROUP_KEY = 'selected_group';
let scheduleData = null;

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