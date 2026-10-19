const GROUP_KEY = 'selected_group';
const VERSION_KEY = 'app_version';

const MONTHS_RU = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня',
    'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'];
const MONTHS_SHORT = ['янв', 'фев', 'мар', 'апр', 'мая', 'июн',
    'июл', 'авг', 'сен', 'окт', 'ноя', 'дек'];
const WEEKDAYS_SHORT = ['Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб', 'Вс'];
const WEEKDAYS_FULL = ['понедельник', 'вторник', 'среда', 'четверг', 'пятница', 'суббота', 'воскресенье'];

// ---------- Состояние ----------
let days = {};            // { '2026-10-20': { groups, schedule, ... } } — вся накопленная история
let allGroups = [];       // объединение групп по всем дням
let selectedIso = null;   // выбранный день
let currentGroup = '';
let currentVersion = 'student'; // 'student' или 'teacher'

// ---------- Расписание звонков ----------
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

// Время пары по номеру. В понедельник первая строка — классный час,
// поэтому номер пары совпадает с индексом; в остальные дни сдвиг на 1.
function getLessonTime(lessonNum, isMonday) {
    const schedule = isMonday ? bellSchedule.monday : bellSchedule.other;
    const index = isMonday ? lessonNum : lessonNum - 1;
    if (index >= 0 && index < schedule.length) {
        return schedule[index];
    }
    return null;
}

// ---------- Работа с датами (без UTC, чтобы день не «съезжал») ----------
function toIso(d) {
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
}
function fromIso(iso) {
    const [y, m, d] = iso.split('-').map(Number);
    return new Date(y, m - 1, d);
}
function addDays(iso, n) {
    const d = fromIso(iso);
    d.setDate(d.getDate() + n);
    return toIso(d);
}
function mondayOf(iso) {
    const d = fromIso(iso);
    d.setDate(d.getDate() - ((d.getDay() + 6) % 7));
    return toIso(d);
}
function weekdayIndex(iso) { // 0 = понедельник ... 6 = воскресенье
    return (fromIso(iso).getDay() + 6) % 7;
}
// Понедельник определяем по реальной дате, а не по тексту
function isMondayIso(iso) {
    return weekdayIndex(iso) === 0;
}
function formatLongDate(iso) {
    const d = fromIso(iso);
    return `${d.getDate()} ${MONTHS_RU[d.getMonth()]} ${d.getFullYear()}, ${WEEKDAYS_FULL[weekdayIndex(iso)]}`;
}
function formatShortDate(iso) {
    const d = fromIso(iso);
    return `${d.getDate()} ${MONTHS_SHORT[d.getMonth()]}`;
}
function parseRuDate(text) {
    text = text || '';
    let m = text.match(/(\d{1,2})\s+([а-яё]+)\s+(\d{4})/i);
    if (m) {
        const idx = MONTHS_RU.indexOf(m[2].toLowerCase());
        if (idx >= 0) return toIso(new Date(+m[3], idx, +m[1]));
    }
    m = text.match(/(\d{1,2})\.(\d{1,2})\.(\d{4})/);
    if (m) return toIso(new Date(+m[3], +m[2] - 1, +m[1]));
    return null;
}

// ---------- Загрузка данных ----------
async function fetchJson(url) {
    try {
        const resp = await fetch(url + '?t=' + Date.now()); // cache busting
        if (!resp.ok) return null;
        return await resp.json();
    } catch (e) {
        return null;
    }
}

function buildDays(current, history) {
    const result = {};
    Object.entries(history?.days || {}).forEach(([iso, day]) => {
        result[iso] = day;
    });
    // schedule.json — самая свежая выгрузка, она главнее записи в истории
    if (current) {
        const iso = current.date_iso || parseRuDate(current.date);
        if (iso) result[iso] = { groups: current.groups || [], schedule: current.schedule || {} };
    }
    return result;
}

// Какой день показать при открытии: сегодняшний → ближайший будущий → последний
function pickInitialDay() {
    const keys = Object.keys(days).sort();
    const today = toIso(new Date());
    if (days[today]) return today;
    const upcoming = keys.find(k => k > today);
    return upcoming || keys[keys.length - 1];
}

async function init() {
    const loader = document.getElementById('loader');
    const container = document.getElementById('schedule-container');

    try {
        loader.classList.remove('hidden');
        const [current, history] = await Promise.all([
            fetchJson('schedule.json'),
            fetchJson('history.json')
        ]);
        days = buildDays(current, history);
        if (Object.keys(days).length === 0) throw new Error('Нет данных расписания');

        const groupSet = new Set();
        Object.values(days).forEach(d => (d.groups || []).forEach(g => groupSet.add(g)));
        allGroups = Array.from(groupSet).sort();

        selectedIso = pickInitialDay();
        currentGroup = localStorage.getItem(GROUP_KEY) || '';
        if (!allGroups.includes(currentGroup)) currentGroup = '';

        const savedVersion = localStorage.getItem(VERSION_KEY);
        if (savedVersion === 'student' || savedVersion === 'teacher') currentVersion = savedVersion;

        populateGroups();
        initWeekNav();
        initBellSchedule();
        initVersionSwitcher();
        renderAll();
    } catch (e) {
        container.innerHTML = '<div class="placeholder">❌ Не удалось загрузить расписание.<br>Проверьте файлы schedule.json и history.json</div>';
        console.error(e);
    } finally {
        loader.classList.add('hidden');
    }
}

function populateGroups() {
    const select = document.getElementById('group-select');
    allGroups.forEach(g => {
        const opt = document.createElement('option');
        opt.value = g;
        opt.textContent = g;
        select.appendChild(opt);
    });
    select.value = currentGroup;

    select.addEventListener('change', (e) => {
        currentGroup = e.target.value;
        localStorage.setItem(GROUP_KEY, currentGroup);
        renderAll();
    });
}

// ---------- Неделя ----------
function initWeekNav() {
    document.getElementById('week-nav').addEventListener('click', (e) => {
        const chip = e.target.closest('[data-iso]');
        if (chip && !chip.disabled) {
            selectedIso = chip.dataset.iso;
            renderAll();
            return;
        }
        const arrow = e.target.closest('[data-week]');
        if (arrow && !arrow.disabled) {
            const targetWeek = arrow.dataset.week;
            const inWeek = Object.keys(days).filter(k => mondayOf(k) === targetWeek).sort();
            // тот же день недели, если он есть; иначе первый доступный
            const sameDay = addDays(targetWeek, weekdayIndex(selectedIso));
            selectedIso = days[sameDay] ? sameDay : inWeek[0];
            renderAll();
        }
    });
}

function renderWeekNav() {
    const nav = document.getElementById('week-nav');
    const weekStart = mondayOf(selectedIso);
    const weeks = Array.from(new Set(Object.keys(days).map(mondayOf))).sort();
    const idx = weeks.indexOf(weekStart);
    const prevWeek = idx > 0 ? weeks[idx - 1] : null;
    const nextWeek = idx < weeks.length - 1 ? weeks[idx + 1] : null;

    const count = days[addDays(weekStart, 6)] ? 7 : 6; // воскресенье — только если есть данные
    const todayIso = toIso(new Date());

    let chips = '';
    for (let i = 0; i < count; i++) {
        const iso = addDays(weekStart, i);
        const has = !!days[iso];
        const cls = ['day-chip', iso === selectedIso ? 'active' : '', has ? '' : 'no-data', iso === todayIso ? 'today' : '']
            .filter(Boolean).join(' ');
        chips += `<button class="${cls}" data-iso="${iso}" ${has ? '' : 'disabled'}>
            <span class="day-chip-name">${WEEKDAYS_SHORT[i]}</span>
            <span class="day-chip-num">${fromIso(iso).getDate()}</span>
        </button>`;
    }

    nav.innerHTML = `
        <div class="week-head">
            <button class="week-arrow" data-week="${prevWeek || ''}" ${prevWeek ? '' : 'disabled'} aria-label="Предыдущая неделя">‹</button>
            <span class="week-label">${formatShortDate(weekStart)} – ${formatShortDate(addDays(weekStart, count - 1))}</span>
            <button class="week-arrow" data-week="${nextWeek || ''}" ${nextWeek ? '' : 'disabled'} aria-label="Следующая неделя">›</button>
        </div>
        <div class="week-days">${chips}</div>
    `;
}

// ---------- Отрисовка ----------
function renderAll() {
    const groupSelect = document.getElementById('group-select');
    const headerTitle = document.getElementById('date-display');
    const studentBtn = document.getElementById('student-version-btn');
    const teacherBtn = document.getElementById('teacher-version-btn');

    const isStudent = currentVersion === 'student';
    studentBtn.classList.toggle('active', isStudent);
    teacherBtn.classList.toggle('active', !isStudent);
    groupSelect.style.display = isStudent ? 'block' : 'none';
    headerTitle.textContent = isStudent
        ? formatLongDate(selectedIso)
        : `Преподаватели · ${formatShortDate(selectedIso)}`;

    renderWeekNav();
    if (isStudent) {
        renderSchedule(currentGroup);
    } else {
        renderTeacherView();
    }
}

function renderSchedule(group) {
    const container = document.getElementById('schedule-container');

    if (!group) {
        container.innerHTML = '<div class="placeholder">👆 Выберите группу выше</div>';
        return;
    }

    const lessons = days[selectedIso]?.schedule?.[group];

    if (!lessons || lessons.length === 0) {
        container.innerHTML = '<div class="placeholder">На этот день занятий нет 🎉</div>';
        return;
    }

    const isMonday = isMondayIso(selectedIso);

    container.innerHTML = lessons.map(l => {
        const lessonNum = parseInt(l.num);
        const timeInfo = getLessonTime(lessonNum, isMonday);
        const timeDisplay = timeInfo ? `<span class="lesson-time">${timeInfo.time}</span>` : '';

        // Показываем предмет, если он есть. Если предмета нет, но есть преподаватель или кабинет — показываем "Пара"
        // Если вообще ничего нет — показываем пустое значение (без надписи "Окно")
        const subjectDisplay = l.subject ? l.subject : (l.teacher || l.room ? 'Пара' : '');

        return `
        <div class="lesson-card ${l.subject || l.teacher || l.room ? '' : 'empty'}">
            <div class="lesson-header">
                <span class="lesson-num">${l.num}-я пара</span>
                ${timeDisplay}
                ${subjectDisplay ? `<span class="lesson-subject">${subjectDisplay}</span>` : ''}
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

// ---------- Расписание звонков ----------
function initBellSchedule() {
    const modal = document.getElementById('bell-modal');
    const openBtn = document.getElementById('bell-schedule-btn');
    const closeBtn = document.getElementById('close-modal');
    const tabBtns = document.querySelectorAll('.tab-btn');

    function showBellTab(day) {
        tabBtns.forEach(b => b.classList.toggle('active', b.dataset.day === day));
        renderBellSchedule(day);
    }

    // Открываем сразу ту вкладку, которая соответствует выбранному дню
    openBtn.addEventListener('click', () => {
        modal.classList.remove('hidden');
        showBellTab(isMondayIso(selectedIso) ? 'monday' : 'other');
    });

    closeBtn.addEventListener('click', () => {
        modal.classList.add('hidden');
    });

    // Закрытие по клику вне контента
    modal.addEventListener('click', (e) => {
        if (e.target === modal) {
            modal.classList.add('hidden');
        }
    });

    tabBtns.forEach(btn => {
        btn.addEventListener('click', () => showBellTab(btn.dataset.day));
    });
}

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

// ---------- Переключатель версий ----------
function initVersionSwitcher() {
    document.getElementById('student-version-btn').addEventListener('click', () => {
        if (currentVersion !== 'student') {
            currentVersion = 'student';
            localStorage.setItem(VERSION_KEY, 'student');
            renderAll();
        }
    });

    document.getElementById('teacher-version-btn').addEventListener('click', () => {
        if (currentVersion !== 'teacher') {
            currentVersion = 'teacher';
            localStorage.setItem(VERSION_KEY, 'teacher');
            renderAll();
        }
    });
}

// ---------- Версия для преподавателей ----------
function renderTeacherView() {
    const container = document.getElementById('schedule-container');
    const day = days[selectedIso];

    // Собираем всех уникальных преподавателей за выбранный день
    const teachersMap = new Map();

    (day?.groups || []).forEach(group => {
        const lessons = day.schedule?.[group] || [];
        lessons.forEach(lesson => {
            if (lesson.teacher && lesson.subject) {
                const teacherName = lesson.teacher.trim();

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

    const isMonday = isMondayIso(selectedIso);

    const sortedTeachers = Array.from(teachersMap.entries()).sort((a, b) =>
        a[0].localeCompare(b[0])
    );

    if (sortedTeachers.length === 0) {
        container.innerHTML = '<div class="placeholder">На этот день занятий нет 🎉</div>';
        return;
    }

    let html = '';
    sortedTeachers.forEach(([teacherName, lessons]) => {
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
