const GROUP_KEY = 'selected_group';
const VERSION_KEY = 'app_version';
let scheduleData = null;
let currentVersion = 'student'; // 'student' или 'teacher'

// Расписание звонков
const bellSchedule = {
    monday: [
        { name: 'Классный час', time: '08:45 – 09:15', start: '08:45', end: '09:15' },
        { name: '1 пара', time: '09:20 – 10:40', start: '09:20', end: '10:40' },
        { name: '2 пара', time: '10:50 – 12:10', start: '10:50', end: '12:10' },
        { name: '3 пара', time: '12:25 – 13:45', start: '12:25', end: '13:45' },
        { name: '4 пара', time: '13:50 – 15:10', start: '13:50', end: '15:10' },
        { name: '5 пара', time: '15:15 – 16:35', start: '15:15', end: '16:35' },
        { name: '6 пара', time: '16:40 – 18:00', start: '16:40', end: '18:00' }
    ],
    other: [
        { name: '1 пара', time: '08:45 – 10:05', start: '08:45', end: '10:05' },
        { name: '2 пара', time: '10:15 – 11:35', start: '10:15', end: '11:35' },
        { name: '3 пара', time: '11:45 – 13:05', start: '11:45', end: '13:05' },
        { name: '4 пара', time: '13:20 – 14:40', start: '13:20', end: '14:40' },
        { name: '5 пара', time: '14:45 – 16:05', start: '14:45', end: '16:05' },
        { name: '6 пара', time: '16:10 – 17:30', start: '16:10', end: '17:30' }
    ]
};

// Функция для получения времени пары по номеру
function getLessonTime(lessonNum, isMonday) {
    const schedule = isMonday ? bellSchedule.monday : bellSchedule.other;
    // Для понедельника классный час - это 0 пара, поэтому сдвигаем индекс
    const index = isMonday ? lessonNum : lessonNum - 1;
    if (index >= 0 && index < schedule.length) {
        return schedule[index];
    }
    return null;
}

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
        initVersionSwitcher();
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
    
    // Определяем день недели (для простоты считаем понедельник, если дата содержит "понедельник")
    const dateStr = scheduleData.date || '';
    const isMonday = dateStr.toLowerCase().includes('понедельник');
    
    container.innerHTML = lessons.map(l => {
        const lessonNum = parseInt(l.num);
        const timeInfo = getLessonTime(lessonNum, isMonday);
        const timeDisplay = timeInfo ? `<span class="lesson-time">${timeInfo.time}</span>` : '';
        
        return `
        <div class="lesson-card ${l.subject ? '' : 'empty'}">
            <div class="lesson-header">
                <span class="lesson-num">${l.num}-я пара</span>
                ${timeDisplay}
                <span class="lesson-subject">${l.subject || 'Окно'}</span>
            </div>
            ${l.teacher || l.room ? `<div class="lesson-meta">
                ${l.teacher ? `<span>👨‍🏫 ${l.teacher}</span>` : ''}
                ${l.room ? `<span>🚪 Каб. ${l.room}</span>` : ''}
            </div>` : ''}
        </div>
    `;
    }).join('');
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

// Инициализация переключателя версий
function initVersionSwitcher() {
    const studentBtn = document.getElementById('student-version-btn');
    const teacherBtn = document.getElementById('teacher-version-btn');
    const groupSelect = document.getElementById('group-select');
    const headerTitle = document.getElementById('date-display');
    
    // Восстанавливаем сохраненную версию
    const savedVersion = localStorage.getItem(VERSION_KEY);
    if (savedVersion && (savedVersion === 'student' || savedVersion === 'teacher')) {
        currentVersion = savedVersion;
        updateVersionUI();
    }
    
    // Переключение на студенческую версию
    studentBtn.addEventListener('click', () => {
        if (currentVersion !== 'student') {
            currentVersion = 'student';
            localStorage.setItem(VERSION_KEY, 'student');
            updateVersionUI();
            restoreSelection();
        }
    });
    
    // Переключение на версию преподавателя
    teacherBtn.addEventListener('click', () => {
        if (currentVersion !== 'teacher') {
            currentVersion = 'teacher';
            localStorage.setItem(VERSION_KEY, 'teacher');
            updateVersionUI();
            renderTeacherView();
        }
    });
    
    function updateVersionUI() {
        if (currentVersion === 'student') {
            studentBtn.classList.add('active');
            teacherBtn.classList.remove('active');
            groupSelect.style.display = 'block';
            headerTitle.textContent = scheduleData.date || 'Расписание';
        } else {
            studentBtn.classList.remove('active');
            teacherBtn.classList.add('active');
            groupSelect.style.display = 'none';
            headerTitle.textContent = 'Преподаватели';
        }
    }
}

// Рендеринг версии для преподавателей
function renderTeacherView() {
    const container = document.getElementById('schedule-container');
    const allGroups = scheduleData.groups || [];
    
    if (allGroups.length === 0) {
        container.innerHTML = '<div class="placeholder">Нет данных о группах</div>';
        return;
    }
    
    // Собираем всех уникальных преподавателей
    const teachersMap = new Map();
    
    allGroups.forEach(group => {
        const lessons = scheduleData.schedule?.[group] || [];
        lessons.forEach(lesson => {
            if (lesson.teacher && lesson.subject) {
                // Очищаем имя преподавателя от лишних данных
                let teacherName = lesson.teacher.trim();
                
                if (!teachersMap.has(teacherName)) {
                    teachersMap.set(teacherName, []);
                }
                
                teachersMap.get(teacherName).push({
                    group: group,
                    subject: lesson.subject,
                    room: lesson.room || '',
                    lessonNum: lesson.num,
                    num: parseInt(lesson.num) || 99
                });
            }
        });
    });
    
    // Определяем день недели
    const dateStr = scheduleData.date || '';
    const isMonday = dateStr.toLowerCase().includes('понедельник');
    
    // Сортируем преподавателей по имени
    const sortedTeachers = Array.from(teachersMap.entries()).sort((a, b) => 
        a[0].localeCompare(b[0])
    );
    
    if (sortedTeachers.length === 0) {
        container.innerHTML = '<div class="placeholder">На сегодня занятий нет 🎉</div>';
        return;
    }
    
    let html = '';
    sortedTeachers.forEach(([teacherName, lessons]) => {
        // Сортируем пары по номеру
        lessons.sort((a, b) => a.num - b.num);
        
        html += `
        <div class="lesson-card">
            <div class="lesson-header">
                <span class="lesson-subject">👨‍🏫 ${teacherName}</span>
            </div>
        `;
        
        lessons.forEach(lesson => {
            const timeInfo = getLessonTime(lesson.num, isMonday);
            const timeDisplay = timeInfo ? `<span class="lesson-time">${timeInfo.time}</span>` : '';
            
            html += `
            <div class="teacher-lesson">
                <div class="teacher-lesson-header">
                    <span class="lesson-num">${lesson.num}-я пара</span>
                    ${timeDisplay}
                </div>
                <div class="teacher-lesson-subject">${lesson.subject}</div>
                <div class="teacher-lesson-meta">
                    <span>Группа: ${lesson.group}</span>
                    ${lesson.room ? `<span>Каб. ${lesson.room}</span>` : ''}
                </div>
            </div>
            `;
        });
        
        html += '</div>';
    });
    
    container.innerHTML = html;
}