'use strict';

const STORAGE_KEY = 'khotwa-tasks';
// ضع هنا رابط Web App الذي تحصل عليه من Google Apps Script.
const SHEET_API_URL = '';
const priorityNames = { high: 'عالية', medium: 'متوسطة', low: 'منخفضة' };
const categoryNames = { personal: 'شخصي', work: 'عمل', achievement: 'إنجازات' };
const taskForm = document.querySelector('#taskForm');
const taskTitleInput = document.querySelector('#taskTitle');
const taskPriorityInput = document.querySelector('#taskPriority');
const taskCategoryInput = document.querySelector('#taskCategory');
const taskDueDateInput = document.querySelector('#taskDueDate');
const taskList = document.querySelector('#taskList');
const emptyState = document.querySelector('#emptyState');
const emptyTitle = document.querySelector('#emptyTitle');
const emptyMessage = document.querySelector('#emptyMessage');
const searchInput = document.querySelector('#searchInput');
const remainingCount = document.querySelector('#remainingCount');
const progressText = document.querySelector('#progressText');
const taskTotal = document.querySelector('#taskTotal');
const submitLabel = document.querySelector('#submitLabel');
const cancelEditButton = document.querySelector('#cancelEdit');
const statusMessage = document.querySelector('#statusMessage');
const syncStatus = document.querySelector('#syncStatus');

let tasks = loadTasks();
let activeFilter = 'all';
let editingTaskId = null;
let remoteSyncTimer = null;
let isFetchingRemote = false;
let localWritePending = false;
let localRevision = 0;
let hasLoadedSheet = false;

// نحمل المهام المحفوظة، ونتعامل مع البيانات التالفة بأمان.
function loadTasks() {
  try {
    const savedTasks = JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]');
    return Array.isArray(savedTasks) ? savedTasks : [];
  } catch {
    return [];
  }
}

// نحفظ القائمة كاملة بعد كل تغيير.
function saveTasks() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(tasks));
  } catch {
    announce('تعذر حفظ المهام في هذا المتصفح.');
  }

  if (SHEET_API_URL) {
    localRevision += 1;
    localWritePending = true;
    syncStatus.textContent = 'جارٍ إرسال التغييرات إلى Google Sheets...';
    clearTimeout(remoteSyncTimer);
    const revision = localRevision;
    remoteSyncTimer = setTimeout(() => pushTasksToSheet(revision), 500);
  }
}

function announce(message) {
  statusMessage.textContent = message;
}

// نقرأ الشيت عبر JSONP حتى يعمل الطلب من صفحة HTML محلية أيضًا.
function readTasksFromSheet() {
  return new Promise((resolve, reject) => {
    const callbackName = `khotwaSheet_${Date.now()}_${Math.random().toString(36).slice(2)}`;
    const script = document.createElement('script');
    let settled = false;
    const finish = (error, result) => {
      if (settled) return;
      settled = true;
      clearTimeout(timeoutId);
      script.remove();
      delete window[callbackName];
      if (error) reject(error);
      else resolve(result);
    };

    const timeoutId = setTimeout(() => finish(new Error('انتهت مهلة الاتصال')), 12000);
    window[callbackName] = (result) => {
      if (!result || !result.ok || !Array.isArray(result.tasks)) {
        finish(new Error('تعذر قراءة المهام من الشيت'));
        return;
      }
      finish(null, result.tasks);
    };
    script.onerror = () => finish(new Error('تعذر الاتصال بالشيت'));
    const separator = SHEET_API_URL.includes('?') ? '&' : '?';
    script.src = `${SHEET_API_URL}${separator}action=read&callback=${callbackName}&_=${Date.now()}`;
    document.head.append(script);
  });
}

// نرسل القائمة كاملة إلى الشيت بعد أن يهدأ المستخدم عن الكتابة قليلًا.
async function pushTasksToSheet(revision = localRevision) {
  if (!SHEET_API_URL) return;
  const savedTasks = tasks.map((task) => ({ ...task }));

  try {
    await fetch(SHEET_API_URL, {
      method: 'POST',
      mode: 'no-cors',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8' },
      body: new URLSearchParams({ payload: JSON.stringify({ action: 'save', tasks: savedTasks }) })
    });
    if (revision === localRevision) {
      localWritePending = false;
      syncStatus.textContent = 'تم إرسال التغييرات؛ يجري تحديث القائمة من الشيت تلقائيًا.';
    }
  } catch {
    syncStatus.textContent = 'تعذر الاتصال بالشيت؛ ستبقى مهامك محفوظة على هذا الجهاز.';
    remoteSyncTimer = setTimeout(() => pushTasksToSheet(revision), 10000);
  }
}

// نحدّث الموقع من الشيت دوريًا، إلا إذا كانت هناك تعديلات محلية لم تُرسل بعد.
async function syncTasksFromSheet() {
  if (!SHEET_API_URL || isFetchingRemote || localWritePending || editingTaskId) return;
  isFetchingRemote = true;

  try {
    const remoteTasks = await readTasksFromSheet();
    if (localWritePending) return;

    if (!hasLoadedSheet && remoteTasks.length === 0 && tasks.length > 0) {
      hasLoadedSheet = true;
      saveTasks();
      return;
    }

    tasks = remoteTasks;
    hasLoadedSheet = true;
    localStorage.setItem(STORAGE_KEY, JSON.stringify(tasks));
    syncStatus.textContent = 'متصل بـ Google Sheets؛ تتم المزامنة تلقائيًا.';
    renderTasks();
  } catch {
    syncStatus.textContent = 'تعذر تحديث المهام من الشيت؛ تحقق من رابط النشر والاتصال.';
  } finally {
    isFetchingRemote = false;
  }
}

function formatDate(dateString) {
  if (!dateString) return '';
  const [year, month, day] = dateString.split('-').map(Number);
  return new Intl.DateTimeFormat('ar', { day: 'numeric', month: 'short' }).format(new Date(year, month - 1, day));
}

function getTodayString() {
  const today = new Date();
  const localDate = new Date(today.getTime() - today.getTimezoneOffset() * 60000);
  return localDate.toISOString().slice(0, 10);
}

function isOverdue(task) {
  return task.dueDate && task.dueDate < getTodayString() && !task.completed;
}

// ننشئ عناصر المهمة باستخدام textContent حتى يبقى العنوان نصًا عاديًا.
function createTaskCard(task) {
  const card = document.createElement('article');
  card.className = `task-card${task.completed ? ' is-completed' : ''}`;
  card.dataset.taskId = task.id;

  const checkbox = document.createElement('input');
  checkbox.className = 'task-check';
  checkbox.type = 'checkbox';
  checkbox.checked = Boolean(task.completed);
  checkbox.setAttribute('aria-label', task.completed ? `إلغاء إكمال ${task.title}` : `إكمال ${task.title}`);

  const main = document.createElement('div');
  main.className = 'task-main';
  const description = document.createElement('div');
  description.className = 'task-description';
  const title = document.createElement('p');
  title.className = 'task-title';
  title.textContent = task.title;
  const meta = document.createElement('div');
  meta.className = 'task-meta';

  const priority = document.createElement('span');
  priority.className = `meta-tag priority-${task.priority}`;
  priority.textContent = `أولوية ${priorityNames[task.priority] || priorityNames.medium}`;
  const category = document.createElement('span');
  category.className = 'meta-tag';
  category.textContent = categoryNames[task.category] || categoryNames.personal;
  meta.append(priority, category);

  if (task.dueDate) {
    const dueDate = document.createElement('span');
    dueDate.className = `due-date${isOverdue(task) ? ' overdue' : ''}`;
    dueDate.textContent = `${isOverdue(task) ? 'متأخرة · ' : 'تستحق · '}${formatDate(task.dueDate)}`;
    meta.append(dueDate);
  }

  description.append(title, meta);
  main.append(description);

  const actions = document.createElement('div');
  actions.className = 'task-actions';
  const editButton = document.createElement('button');
  editButton.className = 'icon-button';
  editButton.type = 'button';
  editButton.dataset.action = 'edit';
  editButton.setAttribute('aria-label', `تعديل ${task.title}`);
  editButton.title = 'تعديل المهمة';
  editButton.textContent = '✎';
  const deleteButton = document.createElement('button');
  deleteButton.className = 'icon-button delete-button';
  deleteButton.type = 'button';
  deleteButton.dataset.action = 'delete';
  deleteButton.setAttribute('aria-label', `حذف ${task.title}`);
  deleteButton.title = 'حذف المهمة';
  deleteButton.textContent = '⌫';
  actions.append(editButton, deleteButton);

  card.append(checkbox, main, actions);
  return card;
}

// نطبق البحث والتصفية ثم نحدّث القائمة والعدادات.
function renderTasks() {
  const searchTerm = searchInput.value.trim().toLocaleLowerCase('ar');
  const visibleTasks = tasks.filter((task) => {
    const matchesFilter = activeFilter === 'all'
      || (activeFilter === 'active' && !task.completed)
      || (activeFilter === 'completed' && task.completed);
    return matchesFilter && task.title.toLocaleLowerCase('ar').includes(searchTerm);
  });

  taskList.replaceChildren(...visibleTasks.map(createTaskCard));
  emptyState.hidden = visibleTasks.length > 0;

  if (tasks.length === 0) {
    emptyTitle.textContent = 'مساحتك جاهزة';
    emptyMessage.textContent = 'أضف أول مهمة، ودع يومك يبدأ بخطوة جميلة.';
  } else if (visibleTasks.length === 0) {
    emptyTitle.textContent = 'لا توجد نتائج';
    emptyMessage.textContent = 'جرّب كلمة بحث أخرى أو غيّر طريقة التصفية.';
  }

  const remaining = tasks.filter((task) => !task.completed).length;
  remainingCount.textContent = remaining;
  progressText.textContent = tasks.length === 0
    ? 'ابدأ بإضافة أول مهمة لك'
    : remaining === 0 ? 'أحسنت! أنجزت جميع مهامك' : `لديك ${remaining} ${remaining === 1 ? 'مهمة' : 'مهام'} بانتظارك`;
  taskTotal.textContent = `${tasks.length} ${tasks.length === 1 ? 'مهمة' : 'مهام'}`;
}

function resetForm() {
  taskForm.reset();
  editingTaskId = null;
  submitLabel.textContent = 'أضف المهمة';
  cancelEditButton.hidden = true;
}

// يضيف النموذج مهمة جديدة أو يحفظ تعديلات المهمة المحددة.
taskForm.addEventListener('submit', (event) => {
  event.preventDefault();
  const title = taskTitleInput.value.trim();
  if (!title) {
    taskTitleInput.focus();
    return;
  }

  if (editingTaskId) {
    tasks = tasks.map((task) => task.id === editingTaskId
      ? { ...task, title, priority: taskPriorityInput.value, category: taskCategoryInput.value, dueDate: taskDueDateInput.value }
      : task);
    announce('تم تحديث المهمة.');
  } else {
    tasks.unshift({
      id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
      title,
      priority: taskPriorityInput.value,
      category: taskCategoryInput.value,
      dueDate: taskDueDateInput.value,
      completed: false
    });
    announce('تمت إضافة المهمة.');
  }

  saveTasks();
  resetForm();
  renderTasks();
  taskTitleInput.focus();
});

// تعالج أزرار الإكمال والتعديل والحذف داخل القائمة.
taskList.addEventListener('click', (event) => {
  const button = event.target.closest('button[data-action]');
  if (!button) return;

  const card = button.closest('[data-task-id]');
  const task = tasks.find((item) => item.id === card.dataset.taskId);
  if (!task) return;

  if (button.dataset.action === 'edit') {
    editingTaskId = task.id;
    taskTitleInput.value = task.title;
    taskPriorityInput.value = task.priority;
    taskCategoryInput.value = task.category;
    taskDueDateInput.value = task.dueDate || '';
    submitLabel.textContent = 'حفظ التعديل';
    cancelEditButton.hidden = false;
    taskTitleInput.focus();
    taskForm.scrollIntoView({ behavior: 'smooth', block: 'center' });
  }

  if (button.dataset.action === 'delete' && window.confirm(`هل تريد حذف المهمة «${task.title}»؟`)) {
    tasks = tasks.filter((item) => item.id !== task.id);
    if (editingTaskId === task.id) resetForm();
    saveTasks();
    renderTasks();
    announce('تم حذف المهمة.');
  }
});

taskList.addEventListener('change', (event) => {
  if (!event.target.matches('.task-check')) return;
  const card = event.target.closest('[data-task-id]');
  tasks = tasks.map((task) => task.id === card.dataset.taskId
    ? { ...task, completed: event.target.checked }
    : task);
  saveTasks();
  renderTasks();
  announce(event.target.checked ? 'تم إكمال المهمة.' : 'أعيدت المهمة إلى قيد التنفيذ.');
});

// البحث والتصفية يعملان مباشرة دون إعادة تحميل الصفحة.
searchInput.addEventListener('input', renderTasks);
document.querySelectorAll('.filter-button').forEach((button) => {
  button.addEventListener('click', () => {
    activeFilter = button.dataset.filter;
    document.querySelectorAll('.filter-button').forEach((filterButton) => {
      const isActive = filterButton === button;
      filterButton.classList.toggle('active', isActive);
      filterButton.setAttribute('aria-pressed', String(isActive));
    });
    renderTasks();
  });
});

cancelEditButton.addEventListener('click', resetForm);
document.querySelector('#todayLabel').textContent = new Intl.DateTimeFormat('ar', {
  weekday: 'long', day: 'numeric', month: 'long'
}).format(new Date());
renderTasks();

if (SHEET_API_URL) {
  syncStatus.textContent = 'جارٍ الاتصال بـ Google Sheets...';
  syncTasksFromSheet();
  window.setInterval(syncTasksFromSheet, 15000);
} else {
  syncStatus.textContent = 'المهام محفوظة على هذا الجهاز إلى أن يتم ربط Google Sheets.';
}
