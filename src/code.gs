/**
 * Quản lý Công việc - Google Apps Script Backend
 * Sử dụng SpreadsheetApp.getActiveSpreadsheet() để lấy sheet đang mở
 */

// ============ CONFIGURATION ============
const SHEET_NAMES = {
    SETTINGS: 'Cài Đặt',
    DEPARTMENTS: 'Phòng Ban',
    EMPLOYEES: 'Nhân Viên',
    CATEGORIES: 'Danh Mục',
    TASKS: 'Nhiệm Vụ',
    PROJECTS: 'Dự Án',
    COMMENTS: 'Bình Luận',
    NOTIFICATIONS: 'Thông Báo',
    ANNOUNCEMENTS: 'Thông Báo NB',
    SAVED_FILTERS: 'Bộ Lọc TB'
};

// ============ WEB APP ENTRY ============
function doGet() {
    return HtmlService.createHtmlOutputFromFile('index')
        .setTitle('Quản lý Công việc')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ============ HELPERS (OPTIMIZED) ============
var _ssCache = null;
function getSS() {
    if (!_ssCache) _ssCache = SpreadsheetApp.getActiveSpreadsheet();
    return _ssCache;
}

function getSheet(name) {
    return getSS().getSheetByName(name);
}

function formatPhone(phone) {
    if (phone === undefined || phone === null) return '';
    var p = String(phone).trim();
    if (p === '') return '';
    if (/^[1-9]\d{8}$/.test(p)) {
        return '0' + p;
    }
    return p;
}

function generateId(prefix) {
    return prefix + '_' + Date.now() + '_' + Math.random().toString(36).substr(2, 5);
}

function getLocalISOString(d) {
    var date = d || new Date();
    var tz = "GMT+7";
    try {
        tz = Session.getScriptTimeZone() || "GMT+7";
    } catch(e) {}
    return Utilities.formatDate(date, tz, "yyyy-MM-dd'T'HH:mm:ss.SSSXXX");
}

function getLocalDateString(d) {
    var date = d || new Date();
    var tz = "GMT+7";
    try {
        tz = Session.getScriptTimeZone() || "GMT+7";
    } catch(e) {}
    return Utilities.formatDate(date, tz, "yyyy-MM-dd");
}


// Caching Helpers forlarge data (CacheService limit 100KB)
const CACHE_TTL = 300; // 5 minutes
const CHUNK_SIZE = 90000; 

function setInCache(key, data) {
    const cache = CacheService.getScriptCache();
    const json = JSON.stringify(data);
    const numChunks = Math.ceil(json.length / CHUNK_SIZE);
    
    const meta = { numChunks: numChunks, timestamp: Date.now() };
    cache.put(key, JSON.stringify(meta), CACHE_TTL);
    
    for (let i = 0; i < numChunks; i++) {
        const chunk = json.substring(i * CHUNK_SIZE, (i + 1) * CHUNK_SIZE);
        cache.put(key + '_chunk_' + i, chunk, CACHE_TTL);
    }
}

function getFromCache(key) {
    const cache = CacheService.getScriptCache();
    const metaStr = cache.get(key);
    if (!metaStr) return null;
    
    try {
        const meta = JSON.parse(metaStr);
        let json = '';
        for (let i = 0; i < meta.numChunks; i++) {
            const chunk = cache.get(key + '_chunk_' + i);
            if (chunk === null) return null; // Corrupted cache
            json += chunk;
        }
        return JSON.parse(json);
    } catch (e) {
        return null;
    }
}

function clearCache(key) {
    const cache = CacheService.getScriptCache();
    const metaStr = cache.get(key);
    if (!metaStr) return;
    try {
        const meta = JSON.parse(metaStr);
        cache.remove(key);
        for (let i = 0; i < meta.numChunks; i++) {
            cache.remove(key + '_chunk_' + i);
        }
    } catch (e) {
        cache.remove(key);
    }
}

function clearAllCache() {
    Object.keys(SHEET_NAMES).forEach(function(key) {
        clearCache('data_v2_' + SHEET_NAMES[key]);
    });
    return { success: true };
}

function getSheetData(sheetName) {
    const cacheKey = 'data_v2_' + sheetName;
    const cached = getFromCache(cacheKey);
    if (cached) {
        console.log('Cache hit for: ' + sheetName);
        return cached;
    }

    console.log('Cache miss for: ' + sheetName + '. Reading from sheet...');
    const sheet = getSheet(sheetName);
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    // TRỊT ĐỂ LỖI DẤU CÁCH Ở TÊN CỘT BẰNG CÁCH TRIM
    const headers = data[0].map(h => String(h).trim());
    const result = data.slice(1).map((row, idx) => {
        const obj = { _rowIndex: idx + 2 };
        headers.forEach((h, i) => { obj[h] = row[i]; });
        return obj;
    });

    setInCache(cacheKey, result);
    return result;
}

// Optimized: search only the target column, avoid parsing all columns
function findRowByColumn(sheetName, colName, value) {
    const data = getSheetData(sheetName);
    return data.find(r => r[colName] === value) || null;
}

function appendRow(sheetName, rowData) {
    const sheet = getSheet(sheetName);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const row = headers.map(h => rowData[h] === undefined ? '' : rowData[h]);
    sheet.appendRow(row);
    clearCache('data_v2_' + sheetName);
}

// Optimized: batch write entire row in a single setValues() call
function updateRow(sheetName, rowIndex, rowData) {
    const sheet = getSheet(sheetName);
    const numCols = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, numCols).getValues()[0].map(h => String(h).trim());
    var currentRow = sheet.getRange(rowIndex, 1, 1, numCols).getValues()[0];
    var changed = false;
    headers.forEach(function(h, i) {
        if (rowData[h] !== undefined) {
            currentRow[i] = rowData[h];
            changed = true;
        }
    });
    if (changed) {
        sheet.getRange(rowIndex, 1, 1, numCols).setValues([currentRow]);
        clearCache('data_v2_' + sheetName);
    }
}

function deleteRow(sheetName, rowIndex) {
    getSheet(sheetName).deleteRow(rowIndex);
    clearCache('data_v2_' + sheetName);
}

function parseIds(str) {
    return (str || '').toString().split(',').map(function(x) { return x.trim(); }).filter(function(x) { return x; });
}

function ensureColumnExists(sheetName, colName) {
    var sheet = getSheet(sheetName);
    if (!sheet) return;
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h) { return String(h).trim(); });
    if (headers.indexOf(colName) === -1) {
        sheet.getRange(1, sheet.getLastColumn() + 1).setValue(colName);
        clearCache('data_v2_' + sheetName);
    }
}

// ============ BULK DATA FETCH (Single call from frontend) ============
function getAllData(userRole, userId, userDeptId) {
    var ss = getSS();
    
    // Read all sheets in bulk
    var deptsRaw = getSheetData(SHEET_NAMES.DEPARTMENTS);
    var empsRaw = getSheetData(SHEET_NAMES.EMPLOYEES);
    var catsRaw = getSheetData(SHEET_NAMES.CATEGORIES);
    var tasksRaw = getSheetData(SHEET_NAMES.TASKS);
    var projectsRaw = getSheetData(SHEET_NAMES.PROJECTS);
    var notifsRaw = getSheetData(SHEET_NAMES.NOTIFICATIONS);
    var settingsRaw = getSheetData(SHEET_NAMES.SETTINGS);
    
    // Process departments
    var departments = deptsRaw.map(function(d) {
        return { id: d['ID'], name: d['Tên phòng ban'], description: d['Mô tả'], created_at: d['Ngày tạo'] };
    });
    
    // Process categories
    var categories = catsRaw.map(function(c) {
        return {
            id: c['ID'], name: c['Tên danh mục'], icon: c['Icon'],
            color: c['Màu sắc'], is_default: c['Mặc định'] === true || c['Mặc định'] === 'TRUE',
            created_at: c['Ngày tạo']
        };
    });
    
    // Process employees (Manager can now see & assign across all departments)
    var filteredEmps = empsRaw;
    var employees = filteredEmps.map(function(e) {
        var deptIds = parseIds(e['ID Phòng ban']);
        var deptObjs = departments.filter(function(d) { return deptIds.indexOf(d.id) > -1; });
        return {
            id: e['ID'], name: e['Họ tên'], username: e['Tên đăng nhập'],
            password: e['Mật khẩu'],
            email: e['Email'], phone: formatPhone(e['Điện thoại']), avatar: e['Ảnh đại diện'],
            role: e['Vai trò'], department_id: deptIds[0] || '',
            department_ids: deptIds,
            telegram_chat_id: e['Telegram Chat ID'] || '',
            zalo_user_id: e['Zalo User ID'] || '',
            department: deptObjs[0] || null,
            departments: deptObjs,
            created_at: e['Ngày tạo']
        };
    });
    
    // Process projects
    var projects = projectsRaw.map(function(p) {
        var managerId = String(p['ID Người phụ trách'] || '').trim();
        var manager = employees.find(function(e) { return String(e.id).trim() === managerId; });
        
        var memberIdStr = (p['ID Thành viên'] || '').toString();
        var memberIds = memberIdStr ? memberIdStr.split(',').map(function(id) { return id.trim(); }) : [];
        var projectMembers = empsRaw.filter(function(e) { 
            var empId = String(e['ID']).trim();
            return empId && memberIds.indexOf(empId) > -1; 
        }).map(function(e) { 
            return { id: String(e['ID']).trim(), name: e['Họ tên'], avatar: e['Ảnh đại diện'] }; 
        });

        return {
            id: p['ID'], name: p['Tên dự án'], description: p['Mô tả'],
            target: p['Mục tiêu'], status: p['Trạng thái'] || 'active',
            color: p['Màu sắc'] || '#6366f1',
            start_date: p['Ngày bắt đầu'], end_date: p['Ngày kết thúc'],
            target_completed: p['Đạt mục tiêu'] === true || p['Đạt mục tiêu'] === 'TRUE',
            manager_id: managerId,
            manager: manager ? { id: manager.id, name: manager.name, avatar: manager.avatar } : null,
            member_ids: memberIds,
            members: projectMembers,
            created_by: p['ID Người tạo'], created_at: p['Ngày tạo']
        };
    });
    
    // Process tasks
    var filteredTasks = tasksRaw;
    if (userRole === 'Member') {
        filteredTasks = tasksRaw.filter(function(t) {
            var rawVal = (t['ID Người thực hiện'] || '').toString();
            var ids = rawVal.split(',').map(function(id) { return id.trim(); });
            var targetId = String(userId).trim();
            return ids.indexOf(targetId) > -1;
        });
    }
    else if (userRole === 'Manager') {
        var userDeptIds = parseIds(userDeptId);
        var uidTrim = String(userId).trim();
        filteredTasks = tasksRaw.filter(function(t) {
            var taskDeptId = String(t['ID Phòng ban'] || '').trim();
            var creatorId = String(t['ID Người tạo'] || '').trim();
            var aIds = (t['ID Người thực hiện'] || '').toString().split(',').map(function(x) { return x.trim(); });
            return userDeptIds.indexOf(taskDeptId) > -1 || creatorId === uidTrim || aIds.indexOf(uidTrim) > -1;
        });
    }
    
    var tasks = filteredTasks.map(function(t) {
        var assigneeIdStr = (t['ID Người thực hiện'] || '').toString();
        var assigneeIds = assigneeIdStr ? assigneeIdStr.split(',').map(function(id) { return id.trim(); }) : [];
        var taskAssignees = empsRaw.filter(function(e) { 
            var empId = String(e['ID']).trim();
            return empId && assigneeIds.indexOf(empId) > -1; 
        }).map(function(e) { 
            return { id: String(e['ID']).trim(), name: e['Họ tên'], avatar: e['Ảnh đại diện'] }; 
        });
        
        var projId = String(t['ID Dự án'] || '').trim();
        var projectObj = projects.find(function(p) { return String(p.id).trim() === projId; });
        var catId = String(t['ID Danh mục'] || '').trim();
        var deptId = String(t['ID Phòng ban'] || '').trim();

        return {
            id: t['ID'], title: t['Tiêu đề'], description: t['Mô tả'],
            assignee_id: assigneeIdStr, 
            assignee_ids: assigneeIds,
            department_id: deptId,
            category_id: catId, project_id: projId,
            priority: t['Độ ưu tiên'], status: t['Trạng thái'],
            start_date: t['Ngày bắt đầu'] || t['Ngày tạo'],
            due_date: t['Hạn hoàn thành'], completed_at: t['Ngày hoàn thành'],
            created_by: t['ID Người tạo'], created_at: t['Ngày tạo'],
            assignees: taskAssignees,
            assignee: taskAssignees.length > 0 ? taskAssignees[0] : null,
            category: categories.find(function(c) { return String(c.id).trim() === catId; }),
            department: departments.find(function(d) { return String(d.id).trim() === deptId; }),
            project: projectObj || null,
            checklist: (function() { try { return JSON.parse(t['Checklist'] || '[]'); } catch(e) { return []; } })(),
            tags: (t['Tags'] || '').toString().split(',').filter(function(x) { return x.trim(); }),
            starred: (t['Starred'] || '').toString().split(',').filter(function(x) { return x.trim(); }),
            recurrence: t['Lặp lại'] || 'none'
        };
    });
    
    tasks.sort(function(a, b) {
        var da = a.created_at ? new Date(a.created_at) : new Date(0);
        var db = b.created_at ? new Date(b.created_at) : new Date(0);
        return db - da;
    });
    
    // Process notifications
    var userNotifs = notifsRaw
        .filter(function(n) { return n['ID Người nhận'] === userId || n['ID Người nhận'] === 'all'; })
        .sort(function(a, b) { return new Date(b['Ngày tạo']) - new Date(a['Ngày tạo']); })
        .slice(0, 10)
        .map(function(n) {
            return {
                id: n['ID'], user_id: n['ID Người nhận'], title: n['Tiêu đề'],
                description: n['Nội dung'], link: n['Đường dẫn'],
                read: n['Đã đọc'] === true || n['Đã đọc'] === 'TRUE', created_at: n['Ngày tạo']
            };
        });
    
    // Settings
    var settingsObj = settingsRaw.length > 0 ? {
        companyName: settingsRaw[0]['Tên công ty'],
        description: settingsRaw[0]['Mô tả'],
        telegramBotToken: settingsRaw[0]['Telegram Bot Token'] || '',
        zaloBotToken: settingsRaw[0]['Zalo Bot Token'] || '',
        geminiApiKey: settingsRaw[0]['Gemini API Key'] || ''
    } : {};
    
    // Dashboard stats (computed from already-loaded data)
    var stats = computeDashboardStats(tasks, empsRaw, departments, categories, userRole);
    
    return JSON.stringify({
        departments: departments,
        employees: employees,
        categories: categories,
        tasks: tasks,
        projects: projects,
        notifications: userNotifs,
        settings: settingsObj,
        stats: stats
    });
}

function isTaskOverdueForDashboard(t, now) {
    if (!t || t.status === 'done' || !t.due_date) return false;
    var due = new Date(t.due_date);
    if (isNaN(due.getTime())) return false;
    var today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    var dueDay = new Date(due.getFullYear(), due.getMonth(), due.getDate());
    return dueDay < today;
}

function getDashboardStatusCounts(tasks, now) {
    var counts = { todo: 0, inProgress: 0, done: 0, overdue: 0 };
    tasks.forEach(function(t) {
        if (isTaskOverdueForDashboard(t, now)) {
            counts.overdue++;
        } else if (t.status === 'todo') {
            counts.todo++;
        } else if (t.status === 'in-progress') {
            counts.inProgress++;
        } else if (t.status === 'done') {
            counts.done++;
        }
    });
    return counts;
}

function computeDashboardStats(tasks, empsRaw, depts, cats, userRole) {
    var now = new Date();
    var thisMonth = now.getMonth();
    var thisYear = now.getFullYear();

    var totalTasks = tasks.length;
    var doneTasks = tasks.filter(function(t) { return t.status === 'done'; }).length;
    var inProgressTasks = tasks.filter(function(t) { return t.status === 'in-progress'; }).length;
    var todoTasks = tasks.filter(function(t) { return t.status === 'todo'; }).length;
    var overdueTasks = tasks.filter(function(t) {
        return isTaskOverdueForDashboard(t, now);
    }).length;

    var completedThisMonth = tasks.filter(function(t) {
        if (t.status !== 'done' || !t.completed_at) return false;
        var d = new Date(t.completed_at);
        return d.getMonth() === thisMonth && d.getFullYear() === thisYear;
    }).length;

    // Chart 1: Weekly
    var chartWeekly = [];
    for (var i = 6; i >= 0; i--) {
        var d = new Date(); d.setDate(d.getDate() - i);
        var dayStr = d.toLocaleDateString('vi-VN', { weekday: 'short', day: 'numeric' });
        var dayTasks = tasks.filter(function(t) { return new Date(t.created_at).toDateString() === d.toDateString(); });
        var dayCounts = getDashboardStatusCounts(dayTasks, now);
        chartWeekly.push({
            label: dayStr,
            todo: dayCounts.todo,
            inProgress: dayCounts.inProgress,
            done: dayCounts.done,
            overdue: dayCounts.overdue
        });
    }

    var chartStatus = getDashboardStatusCounts(tasks, now);

    // Chart 3: Trend 30 days
    var chartCompletionTrend = [];
    for (var ci = 29; ci >= 0; ci--) {
        var cd = new Date(); cd.setDate(cd.getDate() - ci);
        var cdStr = cd.toLocaleDateString('vi-VN', { day: '2-digit', month: '2-digit' });
        chartCompletionTrend.push({
            label: cdStr,
            completed: tasks.filter(function(t) { return t.status === 'done' && t.completed_at && new Date(t.completed_at).toDateString() === cd.toDateString(); }).length,
            created: tasks.filter(function(t) { return new Date(t.created_at).toDateString() === cd.toDateString(); }).length
        });
    }

    // Chart 4: By dept
    var deptMap = {};
    tasks.forEach(function(t) {
        var dName = (t.department && t.department.name) ? t.department.name : 'Chưa phân bổ';
        if (!deptMap[dName]) deptMap[dName] = { total: 0, done: 0 };
        deptMap[dName].total++;
        if (t.status === 'done') deptMap[dName].done++;
    });
    var chartByDept = Object.keys(deptMap).map(function(k) { return { label: k, total: deptMap[k].total, done: deptMap[k].done }; });

    // Chart 5: Priority
    var chartByPriority = {
        high: tasks.filter(function(t) { return t.priority === 'high'; }).length,
        medium: tasks.filter(function(t) { return t.priority === 'medium'; }).length,
        low: tasks.filter(function(t) { return t.priority === 'low'; }).length
    };

    // Chart 6: Category
    var catMap = {};
    tasks.forEach(function(t) {
        var cName = (t.category && t.category.name) ? t.category.name : 'Không phân loại';
        if (!catMap[cName]) catMap[cName] = { total: 0, done: 0, overdue: 0 };
        catMap[cName].total++;
        if (t.status === 'done') catMap[cName].done++;
        if (isTaskOverdueForDashboard(t, now)) catMap[cName].overdue++;
    });
    var chartByCategory = Object.keys(catMap).map(function(k) { return { label: k, total: catMap[k].total, done: catMap[k].done, overdue: catMap[k].overdue }; });

    return {
        totalTasks: totalTasks, doneTasks: doneTasks, inProgressTasks: inProgressTasks,
        todoTasks: todoTasks, overdueTasks: overdueTasks, completedThisMonth: completedThisMonth,
        totalEmployees: userRole === 'Owner' ? empsRaw.length : 0,
        totalDepartments: userRole === 'Owner' ? depts.length : 0,
        chartWeekly: chartWeekly, chartStatus: chartStatus,
        chartCompletionTrend: chartCompletionTrend, chartByDept: chartByDept,
        chartByPriority: chartByPriority, chartByCategory: chartByCategory
    };
}

// ============ AUTH ============
function login(username, password) {
    var cleanUsername = String(username).trim().toLowerCase();
    var cleanPassword = String(password).trim();

    // Check admin first
    const settingsData = getSheetData(SHEET_NAMES.SETTINGS);
    const settings = settingsData.length > 0 ? settingsData[0] : null;

    if (settings && 
        String(settings['Tên đăng nhập']).trim().toLowerCase() === cleanUsername && 
        String(settings['Mật khẩu']).trim() === cleanPassword) {
        return {
            success: true,
            user: {
                id: 'owner',
                name: settings['Họ tên'],
                username: settings['Tên đăng nhập'],
                email: settings['Email'] || '',
                phone: formatPhone(settings['Điện thoại']),
                avatar: settings['Ảnh đại diện'] || '',
                role: 'Owner',
                department_id: null
            }
        };
    }

    // Check employees
    const emps = getSheetData(SHEET_NAMES.EMPLOYEES);
    const emp = emps.find(function(e) {
        return String(e['Tên đăng nhập']).trim().toLowerCase() === cleanUsername;
    });

    if (emp && String(emp['Mật khẩu']).trim() === cleanPassword) {
        return {
            success: true,
            user: {
                id: emp['ID'],
                name: emp['Họ tên'],
                username: emp['Tên đăng nhập'],
                email: emp['Email'] || '',
                phone: formatPhone(emp['Điện thoại']),
                avatar: emp['Ảnh đại diện'] || '',
                role: emp['Vai trò'],
                department_id: emp['ID Phòng ban']
            }
        };
    }
    return { success: false, message: 'Sai tên đăng nhập hoặc mật khẩu' };
}


function getMe(userId) {
    if (userId === 'owner') {
        const settings = getSheetData(SHEET_NAMES.SETTINGS)[0];
        return {
            success: true,
            user: {
                id: 'owner', name: settings['Họ tên'], username: settings['Tên đăng nhập'],
                email: settings['Email'] || '', phone: settings['Điện thoại'] || '',
                avatar: settings['Ảnh đại diện'] || '', role: 'Owner', department_id: null
            }
        };
    }
    const emp = findRowByColumn(SHEET_NAMES.EMPLOYEES, 'ID', userId);
    if (emp) {
        return {
            success: true,
            user: {
                id: emp['ID'], name: emp['Họ tên'], username: emp['Tên đăng nhập'],
                email: emp['Email'] || '', phone: emp['Điện thoại'] || '',
                avatar: emp['Ảnh đại diện'] || '', role: emp['Vai trò'],
                department_id: emp['ID Phòng ban']
            }
        };
    }
    return { success: false };
}

// ============ SETTINGS ============
function getSettings() {
    const data = getSheetData(SHEET_NAMES.SETTINGS)[0];
    return data ? { 
        companyName: data['Tên công ty'], 
        description: data['Mô tả'], 
        telegramBotToken: data['Telegram Bot Token'] || '',
        zaloBotToken: data['Zalo Bot Token'] || '',
        geminiApiKey: data['Gemini API Key'] || ''
    } : {};
}

function updateCompanySettings(companyName, description, telegramBotToken, zaloBotToken, geminiApiKey) {
    ensureColumnExists(SHEET_NAMES.SETTINGS, 'Telegram Bot Token');
    ensureColumnExists(SHEET_NAMES.SETTINGS, 'Zalo Bot Token');
    ensureColumnExists(SHEET_NAMES.SETTINGS, 'Gemini API Key');

    const updates = {
        'Tên công ty': companyName,
        'Mô tả': description
    };
    if (telegramBotToken !== undefined) updates['Telegram Bot Token'] = telegramBotToken;
    if (zaloBotToken !== undefined) updates['Zalo Bot Token'] = zaloBotToken;
    if (geminiApiKey !== undefined) updates['Gemini API Key'] = geminiApiKey;

    updateRow(SHEET_NAMES.SETTINGS, 2, updates);
    return { success: true };
}

function updateProfile(userId, data) {
    if (userId === 'owner') {
        const updates = {};
        if (data.name) updates['Họ tên'] = data.name;
        if (data.email) updates['Email'] = data.email;
        if (data.phone) updates['Điện thoại'] = String(data.phone).startsWith('0') ? "'" + data.phone : data.phone;
        if (data.avatar) updates['Ảnh đại diện'] = data.avatar;
        if (data.username) updates['Tên đăng nhập'] = data.username;
        if (data.newPassword) updates['Mật khẩu'] = data.newPassword;
        updateRow(SHEET_NAMES.SETTINGS, 2, updates);
    } else {
        const emp = findRowByColumn(SHEET_NAMES.EMPLOYEES, 'ID', userId);
        if (emp) {
            const updates = {};
            if (data.name) updates['Họ tên'] = data.name;
            if (data.email) updates['Email'] = data.email;
            if (data.phone) {
                var phoneVal = data.phone;
                if (String(phoneVal).startsWith('0')) {
                    phoneVal = "'" + phoneVal;
                }
                updates['Điện thoại'] = phoneVal;
            }
            if (data.avatar) updates['Ảnh đại diện'] = data.avatar;
            if (data.username) updates['Tên đăng nhập'] = data.username;
            if (data.newPassword) updates['Mật khẩu'] = data.newPassword;
            updateRow(SHEET_NAMES.EMPLOYEES, emp._rowIndex, updates);
        }
    }
    return { success: true };
}

// ============ DEPARTMENTS ============
function getDepartments() {
    return getSheetData(SHEET_NAMES.DEPARTMENTS).map(d => ({
        id: d['ID'], name: d['Tên phòng ban'], description: d['Mô tả'], created_at: d['Ngày tạo']
    }));
}

function createDepartment(name, description) {
    const id = generateId('pb');
    appendRow(SHEET_NAMES.DEPARTMENTS, {
        'ID': id, 'Tên phòng ban': name, 'Mô tả': description || '', 'Ngày tạo': getLocalISOString()
    });
    return { success: true, id };
}

function updateDepartment(id, name, description) {
    const dept = findRowByColumn(SHEET_NAMES.DEPARTMENTS, 'ID', id);
    if (dept) {
        updateRow(SHEET_NAMES.DEPARTMENTS, dept._rowIndex, { 'Tên phòng ban': name, 'Mô tả': description });
        return { success: true };
    }
    return { success: false, message: 'Không tìm thấy phòng ban' };
}

function deleteDepartment(id) {
    const dept = findRowByColumn(SHEET_NAMES.DEPARTMENTS, 'ID', id);
    if (dept) { deleteRow(SHEET_NAMES.DEPARTMENTS, dept._rowIndex); return { success: true }; }
    return { success: false };
}

// ============ EMPLOYEES ============
function getEmployees(userRole, userDeptId) {
    // Manager không còn bị giới hạn xem/giao việc trong phạm vi phòng mình nữa
    let data = getSheetData(SHEET_NAMES.EMPLOYEES);
    const depts = getDepartments();
    return data.map(e => {
        var deptIds = parseIds(e['ID Phòng ban']);
        var deptObjs = depts.filter(d => deptIds.indexOf(d.id) > -1);
        return {
            id: e['ID'], name: e['Họ tên'], username: e['Tên đăng nhập'],
            password: e['Mật khẩu'],
            email: e['Email'], phone: formatPhone(e['Điện thoại']), avatar: e['Ảnh đại diện'],
            role: e['Vai trò'], department_id: deptIds[0] || '',
            department_ids: deptIds,
            telegram_chat_id: e['Telegram Chat ID'] || '',
            department: deptObjs[0] || null,
            departments: deptObjs,
            created_at: e['Ngày tạo']
        };
    });
}

function createEmployee(data) {
    const cleanUsername = String(data.username).trim().toLowerCase();
    const emps = getSheetData(SHEET_NAMES.EMPLOYEES);
    const existing = emps.find(e => String(e['Tên đăng nhập']).trim().toLowerCase() === cleanUsername);
    if (existing) return { success: false, message: 'Tên đăng nhập đã tồn tại' };
    
    ensureColumnExists(SHEET_NAMES.EMPLOYEES, 'Zalo User ID');
    
    var deptIds = Array.isArray(data.department_ids) ? data.department_ids : (data.department_id ? [data.department_id] : []);
    const id = generateId('nv');
    appendRow(SHEET_NAMES.EMPLOYEES, {
        'ID': id, 'Họ tên': data.name, 'Tên đăng nhập': cleanUsername,
        'Mật khẩu': data.password || '123456', 'Email': data.email || '',
        'Điện thoại': data.phone ? (String(data.phone).startsWith('0') ? "'" + data.phone : data.phone) : '', 'Ảnh đại diện': data.avatar || '',
        'Vai trò': data.role || 'Member', 'ID Phòng ban': deptIds.join(','),
        'Ngày tạo': getLocalISOString(),
        'Telegram Chat ID': data.telegram_chat_id || '',
        'Zalo User ID': data.zalo_user_id || ''
    });
    return { success: true, id };
}

function updateEmployee(id, data) {
    const emp = findRowByColumn(SHEET_NAMES.EMPLOYEES, 'ID', id);
    if (!emp) return { success: false, message: 'Không tìm thấy nhân viên' };
    
    ensureColumnExists(SHEET_NAMES.EMPLOYEES, 'Zalo User ID');
    
    const updates = {};
    if (data.name) updates['Họ tên'] = data.name;
    if (data.username) updates['Tên đăng nhập'] = String(data.username).trim().toLowerCase();
    if (data.password) updates['Mật khẩu'] = data.password;
    if (data.email !== undefined) updates['Email'] = data.email;
    if (data.phone !== undefined) {
        var phoneVal = data.phone || '';
        if (phoneVal && String(phoneVal).startsWith('0')) {
            phoneVal = "'" + phoneVal;
        }
        updates['Điện thoại'] = phoneVal;
    }
    if (data.avatar !== undefined) updates['Ảnh đại diện'] = data.avatar;
    if (data.role) updates['Vai trò'] = data.role;
    if (data.department_ids !== undefined) {
        var newDeptIds = Array.isArray(data.department_ids) ? data.department_ids : parseIds(data.department_ids);
        updates['ID Phòng ban'] = newDeptIds.join(',');
    } else if (data.department_id) {
        updates['ID Phòng ban'] = data.department_id;
    }
    if (data.telegram_chat_id !== undefined) updates['Telegram Chat ID'] = data.telegram_chat_id;
    if (data.zalo_user_id !== undefined) updates['Zalo User ID'] = data.zalo_user_id;
    updateRow(SHEET_NAMES.EMPLOYEES, emp._rowIndex, updates);
    return { success: true };
}

function deleteEmployee(id) {
    const emp = findRowByColumn(SHEET_NAMES.EMPLOYEES, 'ID', id);
    if (emp) { deleteRow(SHEET_NAMES.EMPLOYEES, emp._rowIndex); return { success: true }; }
    return { success: false };
}

// ============ CATEGORIES ============
function getCategories() {
    return getSheetData(SHEET_NAMES.CATEGORIES).map(c => ({
        id: c['ID'], name: c['Tên danh mục'], icon: c['Icon'],
        color: c['Màu sắc'], is_default: c['Mặc định'] === true || c['Mặc định'] === 'TRUE',
        created_at: c['Ngày tạo']
    }));
}

function createCategory(name, icon, color) {
    const id = generateId('dm');
    appendRow(SHEET_NAMES.CATEGORIES, {
        'ID': id, 'Tên danh mục': name, 'Icon': icon, 'Màu sắc': color,
        'Mặc định': false, 'Ngày tạo': getLocalISOString()
    });
    return { success: true, id };
}

function updateCategory(id, name, icon, color) {
    const cat = findRowByColumn(SHEET_NAMES.CATEGORIES, 'ID', id);
    if (cat) {
        updateRow(SHEET_NAMES.CATEGORIES, cat._rowIndex, { 'Tên danh mục': name, 'Icon': icon, 'Màu sắc': color });
        return { success: true };
    }
    return { success: false };
}

function deleteCategory(id) {
    const cat = findRowByColumn(SHEET_NAMES.CATEGORIES, 'ID', id);
    if (cat && cat['Mặc định'] !== true && cat['Mặc định'] !== 'TRUE') {
        deleteRow(SHEET_NAMES.CATEGORIES, cat._rowIndex);
        return { success: true };
    }
    return { success: false, message: 'Không thể xóa danh mục mặc định' };
}

// ============ TASKS ============
function getTasks(userRole, userId, userDeptId) {
    try {
        let data = getSheetData(SHEET_NAMES.TASKS);
        console.log('getTasks - raw data count:', data.length, 'userRole:', userRole);

        if (userRole === 'Member') {
            data = data.filter(t => {
                const assigneeIds = (t['ID Người thực hiện'] || '').toString().split(',').map(id => id.trim());
                return assigneeIds.indexOf(userId) > -1;
            });
        }
        else if (userRole === 'Manager') {
            var managerDeptIds = parseIds(userDeptId);
            var managerUid = String(userId).trim();
            data = data.filter(t => {
                var taskDeptId = String(t['ID Phòng ban'] || '').trim();
                var creatorId = String(t['ID Người tạo'] || '').trim();
                var aIds = (t['ID Người thực hiện'] || '').toString().split(',').map(x => x.trim());
                return managerDeptIds.indexOf(taskDeptId) > -1 || creatorId === managerUid || aIds.indexOf(managerUid) > -1;
            });
        }
        else if (userRole !== 'Owner') return []; // Fallback security

        const emps = getSheetData(SHEET_NAMES.EMPLOYEES);
        const cats = getCategories();
        const depts = getDepartments();

        const resultTasks = data.map(t => {
            const assigneeIdStr = (t['ID Người thực hiện'] || '').toString();
            const assigneeIds = assigneeIdStr ? assigneeIdStr.split(',').map(id => id.trim()) : [];
            const taskAssignees = emps.filter(e => assigneeIds.indexOf(e['ID']) > -1)
                                      .map(e => ({ id: e['ID'], name: e['Họ tên'], avatar: e['Ảnh đại diện'] }));
            
            return {
                id: t['ID'], title: t['Tiêu đề'], description: t['Mô tả'],
                assignee_id: assigneeIdStr, 
                assignee_ids: assigneeIds,
                department_id: t['ID Phòng ban'],
                category_id: t['ID Danh mục'], priority: t['Độ ưu tiên'], status: t['Trạng thái'],
                start_date: t['Ngày bắt đầu'] || t['Ngày tạo'],
                due_date: t['Hạn hoàn thành'], completed_at: t['Ngày hoàn thành'],
                created_by: t['ID Người tạo'], created_at: t['Ngày tạo'],
                project_id: t['ID Dự án'],
                assignees: taskAssignees,
                assignee: taskAssignees.length > 0 ? taskAssignees[0] : null, // Fallback for single-assignee logic
                category: cats.find(c => c.id === t['ID Danh mục']),
                department: depts.find(d => d.id === t['ID Phòng ban']),
                checklist: (() => { try { return JSON.parse(t['Checklist'] || '[]'); } catch(e) { return []; } })(),
                tags: (t['Tags'] || '').toString().split(',').filter(x => x.trim()),
                starred: (t['Starred'] || '').toString().split(',').filter(x => x.trim())
            };
        });
        return resultTasks.sort((a, b) => {
            const da = a.created_at ? new Date(a.created_at) : new Date(0);
            const db = b.created_at ? new Date(b.created_at) : new Date(0);
            return db - da;
        });
    } catch (e) {
        console.error('getTasks error:', e.message, e.stack);
        return [];
    }
}

function createTask(data, createdBy) {
    const assigneeIds = Array.isArray(data.assignee_ids) ? data.assignee_ids : (data.assignee_id ? [data.assignee_id] : []);
    const assigneeIdStr = assigneeIds.join(',');

    // Auto-fill department if missing
    let deptId = data.department_id;
    if (!deptId && assigneeIds.length > 0) {
        const emps = getSheetData(SHEET_NAMES.EMPLOYEES);
        const firstAssignee = emps.find(e => e['ID'] === assigneeIds[0]);
        if (firstAssignee) deptId = parseIds(firstAssignee['ID Phòng ban'])[0] || '';
    }

    // Ensure new columns exist
    ensureColumnExists(SHEET_NAMES.TASKS, 'Checklist');
    ensureColumnExists(SHEET_NAMES.TASKS, 'Tags');
    ensureColumnExists(SHEET_NAMES.TASKS, 'Starred');
    ensureColumnExists(SHEET_NAMES.TASKS, 'Lặp lại');
    ensureColumnExists(SHEET_NAMES.TASKS, 'Lần chạy cuối');
    ensureColumnExists(SHEET_NAMES.TASKS, 'Ngày bắt đầu');
    
    const id = generateId('cv');
    appendRow(SHEET_NAMES.TASKS, {
        'ID': id, 'Tiêu đề': data.title, 'Mô tả': data.description || '',
        'ID Người thực hiện': assigneeIdStr, 'ID Phòng ban': deptId || '',
        'ID Danh mục': data.category_id || '', 'ID Dự án': data.project_id || '',
        'Độ ưu tiên': data.priority || 'medium',
        'Trạng thái': 'todo', 'Hạn hoàn thành': data.due_date, 'Ngày hoàn thành': '',
        'ID Người tạo': createdBy, 'Ngày tạo': getLocalISOString(),
        'Checklist': data.checklist ? JSON.stringify(data.checklist) : '[]',
        'Tags': data.tags ? (Array.isArray(data.tags) ? data.tags.join(',') : data.tags) : '',
        'Starred': '',
        'Lặp lại': data.recurrence || 'none',
        'Lần chạy cuối': '',
        'Ngày bắt đầu': data.start_date || getLocalDateString()
    });
    
    // Create notifications for all assignees
    assigneeIds.forEach(uid => {
        createNotification(uid, 'Nhiệm vụ mới', `Bạn được giao nhiệm vụ: ${data.title}`, `?task_id=${id}`, data.description);
    });
    
    // Log creation
    try { logActivity(id, createdBy, 'created', 'Tạo nhiệm vụ: ' + data.title); } catch(e) {}
    
    return { success: true, id };
}

function updateTask(id, data, userRole, userId) {
    const task = findRowByColumn(SHEET_NAMES.TASKS, 'ID', id);
    if (!task) return { success: false };
    // Bỏ hạn chế Member chuyển trạng thái Done theo kế hoạch đã thống nhất
    const updates = {};
    if (data.title) updates['Tiêu đề'] = data.title;
    if ((userRole === 'Owner' || userRole === 'Manager') && data.description !== undefined) updates['Mô tả'] = data.description;
    
    const oldAssigneeIdStr = (task['ID Người thực hiện'] || '').toString();
    const oldAssigneeIds = oldAssigneeIdStr ? oldAssigneeIdStr.split(',').map(id => id.trim()) : [];
    
    let newAssigneeIds = oldAssigneeIds;
    if (data.assignee_ids) {
        newAssigneeIds = data.assignee_ids;
        updates['ID Người thực hiện'] = newAssigneeIds.join(',');
    } else if (data.assignee_id) {
        newAssigneeIds = [data.assignee_id];
        updates['ID Người thực hiện'] = data.assignee_id;
    }

    if (data.department_id) updates['ID Phòng ban'] = data.department_id;
    if (data.category_id !== undefined) updates['ID Danh mục'] = data.category_id;
    if (data.project_id !== undefined) updates['ID Dự án'] = data.project_id;
    if (data.priority) updates['Độ ưu tiên'] = data.priority;
    if (data.status) {
        updates['Trạng thái'] = data.status;
        if (data.status === 'done' || data.status === 'completed') {
            updates['Ngày hoàn thành'] = getLocalISOString();
        } else {
            updates['Ngày hoàn thành'] = '';
        }
    }
    if (data.due_date) updates['Hạn hoàn thành'] = data.due_date;
    if (data.start_date !== undefined) {
        ensureColumnExists(SHEET_NAMES.TASKS, 'Ngày bắt đầu');
        updates['Ngày bắt đầu'] = data.start_date;
    }
    if (data.checklist !== undefined) {
        ensureColumnExists(SHEET_NAMES.TASKS, 'Checklist');
        updates['Checklist'] = JSON.stringify(data.checklist);
    }
    if (data.tags !== undefined) {
        ensureColumnExists(SHEET_NAMES.TASKS, 'Tags');
        updates['Tags'] = Array.isArray(data.tags) ? data.tags.join(',') : data.tags;
    }
    if (data.recurrence !== undefined) {
        ensureColumnExists(SHEET_NAMES.TASKS, 'Lặp lại');
        updates['Lặp lại'] = data.recurrence;
    }
    updateRow(SHEET_NAMES.TASKS, task._rowIndex, updates);

    // Create notifications for relevant changes
    var taskTitle = data.title || task['Tiêu đề'];

    // Notify new assignees who were not previously assigned
    if (data.assignee_ids || data.assignee_id) {
        newAssigneeIds.forEach(uid => {
            if (oldAssigneeIds.indexOf(uid) === -1) {
                var taskDesc = data.description !== undefined ? data.description : task['Mô tả'];
                createNotification(uid, 'Nhiệm vụ được giao', 'Bạn được giao nhiệm vụ: ' + taskTitle, '?task_id=' + id, taskDesc);
            }
        });

        // Notify previous assignees who were removed
        const removedAssigneeIds = oldAssigneeIds.filter(uid => newAssigneeIds.indexOf(uid) === -1);
        if (removedAssigneeIds.length > 0) {
            const emps = getSheetData(SHEET_NAMES.EMPLOYEES);
            const newAssignees = emps.filter(e => newAssigneeIds.indexOf(e['ID']) > -1);
            const newAssigneeNames = newAssignees.map(e => e['Họ tên']).join(', ') || 'Chưa giao';
            
            removedAssigneeIds.forEach(uid => {
                var description = `1. Tiêu đề nhiệm vụ: ${taskTitle}\n2. Nhiệm vụ đã chuyển cho: ${newAssigneeNames}`;
                createNotification(uid, 'Chuyển giao nhiệm vụ', description, '?task_id=' + id);
            });
        }
    }

    // Log activity
    try {
        var updaterId = userId || 'system';
        if (data.status && data.status !== task['Trạng thái']) {
            logActivity(id, updaterId, 'status_change', task['Trạng thái'] + ' → ' + data.status);
        }
        if (data.assignee_ids || data.assignee_id) {
            logActivity(id, updaterId, 'assignee_change', 'Giao cho: ' + newAssigneeIds.join(', '));
        }
        if (data.checklist !== undefined) {
            logActivity(id, updaterId, 'checklist_change', 'Cập nhật danh sách công việc con');
        }
        if (data.tags !== undefined) {
            logActivity(id, updaterId, 'tags_change', 'Cập nhật danh sách nhãn');
        }
        if (data.title && data.title !== task['Tiêu đề']) {
            logActivity(id, updaterId, 'title_change', 'Đổi tiêu đề: ' + task['Tiêu đề'] + ' → ' + data.title);
        }
    } catch(logErr) { console.error('Activity log error:', logErr); }

    // Notify all current assignees if status changed
    if (data.status && data.status !== task['Trạng thái']) {
        var statusLabels = { 'todo': 'Cần làm', 'in-progress': 'Đang làm', 'done': 'Hoàn thành' };
        var statusLabel = statusLabels[data.status] || data.status;
        newAssigneeIds.forEach(uid => {
            var taskDesc = data.description !== undefined ? data.description : task['Mô tả'];
            createNotification(uid, 'Cập nhật trạng thái', 'Nhiệm vụ "' + taskTitle + '" đã chuyển sang: ' + statusLabel, '?task_id=' + id, taskDesc, true);
        });
    }

    return { success: true };
}

function toggleTaskStar(taskId, userId) {
    ensureColumnExists(SHEET_NAMES.TASKS, 'Starred');
    var task = findRowByColumn(SHEET_NAMES.TASKS, 'ID', taskId);
    if (!task) return { success: false };
    var starred = (task['Starred'] || '').toString().split(',').filter(function(x) { return x.trim(); });
    var idx = starred.indexOf(userId);
    if (idx > -1) {
        starred.splice(idx, 1);
    } else {
        starred.push(userId);
    }
    updateRow(SHEET_NAMES.TASKS, task._rowIndex, { 'Starred': starred.join(',') });
    return { success: true, starred: idx === -1 };
}

function deleteTask(id, userRole) {
    // Chỉ Owner mới được xóa nhiệm vụ
    if (userRole !== 'Owner') return { success: false, message: 'Bạn không có quyền xóa nhiệm vụ' };
    const task = findRowByColumn(SHEET_NAMES.TASKS, 'ID', id);
    if (task) { deleteRow(SHEET_NAMES.TASKS, task._rowIndex); return { success: true }; }
    return { success: false };
}

// ============ COMMENTS ============
function getComments(taskId) {
    const data = getSheetData(SHEET_NAMES.COMMENTS).filter(c => c['ID Nhiệm vụ'] === taskId);
    const emps = getSheetData(SHEET_NAMES.EMPLOYEES);
    const settingsData = getSheetData(SHEET_NAMES.SETTINGS);
    const settings = settingsData.length > 0 ? settingsData[0] : {};

    return data.map(c => {
        let user = null;
        if (c['ID Người bình luận'] === 'owner') {
            user = { id: 'owner', name: settings['Họ tên'], avatar: settings['Ảnh đại diện'] };
        } else {
            const emp = emps.find(e => e['ID'] === c['ID Người bình luận']);
            if (emp) user = { id: emp['ID'], name: emp['Họ tên'], avatar: emp['Ảnh đại diện'] };
        }
        return {
            id: c['ID'], task_id: taskId, user_id: c['ID Người bình luận'], content: c['Nội dung'],
            image: c['Hình ảnh'], created_at: c['Ngày tạo'], user
        };
    });
}

function createComment(taskId, userId, content, image) {
    const id = generateId('bl');
    appendRow(SHEET_NAMES.COMMENTS, {
        'ID': id, 'ID Nhiệm vụ': taskId, 'ID Người bình luận': userId,
        'Nội dung': content, 'Hình ảnh': image || '', 'Ngày tạo': getLocalISOString()
    });
    return { success: true, id };
}

function uploadFileToDrive(base64Data, fileName, mimeType) {
    try {
        var base64Part = base64Data.split(',')[1] || base64Data;
        var decoded = Utilities.base64Decode(base64Part);
        var blob = Utilities.newBlob(decoded, mimeType, fileName);
        
        var rootFolderName = 'coogo.vn';
        var rootFolders = DriveApp.getFoldersByName(rootFolderName);
        var rootFolder;
        if (rootFolders.hasNext()) {
            rootFolder = rootFolders.next();
        } else {
            rootFolder = DriveApp.createFolder(rootFolderName);
        }
        
        var dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT+7", "dd-MM-yyyy");
        var dateFolders = rootFolder.getFoldersByName(dateStr);
        var dateFolder;
        if (dateFolders.hasNext()) {
            dateFolder = dateFolders.next();
        } else {
            dateFolder = rootFolder.createFolder(dateStr);
        }
        
        var file = dateFolder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        
        return {
            success: true,
            id: file.getId(),
            url: file.getUrl(),
            downloadUrl: file.getDownloadUrl(),
            name: fileName,
            mimeType: mimeType
        };
    } catch (e) {
        console.error('Upload error: ', e);
        return { success: false, message: e.toString() };
    }
}

function getUploadFolderAndToken() {
    try {
        var rootFolderName = 'coogo.vn';
        var rootFolders = DriveApp.getFoldersByName(rootFolderName);
        var rootFolder;
        if (rootFolders.hasNext()) {
            rootFolder = rootFolders.next();
        } else {
            rootFolder = DriveApp.createFolder(rootFolderName);
        }
        
        var dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT+7", "dd-MM-yyyy");
        var dateFolders = rootFolder.getFoldersByName(dateStr);
        var dateFolder;
        if (dateFolders.hasNext()) {
            dateFolder = dateFolders.next();
        } else {
            dateFolder = rootFolder.createFolder(dateStr);
        }
        
        return {
            success: true,
            folderId: dateFolder.getId(),
            token: ScriptApp.getOAuthToken()
        };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

function setFileSharing(fileId) {
    try {
        var file = DriveApp.getFileById(fileId);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        return { success: true };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

// ============ NOTIFICATIONS ============
function getNotifications(userId, page, limit) {
    page = page || 1; limit = limit || 10;
    let data = getSheetData(SHEET_NAMES.NOTIFICATIONS)
        .filter(n => n['ID Người nhận'] === userId || n['ID Người nhận'] === 'all')
        .sort((a, b) => new Date(b['Ngày tạo']) - new Date(a['Ngày tạo']));
    const total = data.length;
    data = data.slice((page - 1) * limit, page * limit);
    return {
        data: data.map(n => ({
            id: n['ID'], user_id: n['ID Người nhận'], title: n['Tiêu đề'],
            description: n['Nội dung'], link: n['Đường dẫn'],
            read: n['Đã đọc'] === true || n['Đã đọc'] === 'TRUE', created_at: n['Ngày tạo']
        })),
        total, page, limit
    };
}

function createNotification(userId, title, description, link, taskDesc, skipZalo) {
    const id = generateId('tb');
    appendRow(SHEET_NAMES.NOTIFICATIONS, {
        'ID': id, 'ID Người nhận': userId, 'Tiêu đề': title,
        'Nội dung': description, 'Đường dẫn': link || '', 'Đã đọc': false,
        'Ngày tạo': getLocalISOString()
    });
    
    // === GỬI TELEGRAM & ZALO ===
    try {
        var appUrl = ScriptApp.getService().getUrl();
        var fullLink = link ? (link.indexOf('?') === 0 ? appUrl + link : appUrl) : appUrl;
        
        var extMsg = '🔔 ' + title + '\n\n📌 ' + description;
        if (taskDesc && taskDesc.trim() !== '') {
            var shortDesc = taskDesc.length > 200 ? taskDesc.substring(0, 200) + '...' : taskDesc;
            extMsg += '\n📝 Mô tả: ' + shortDesc;
        }
        extMsg += '\n🔗 Chi tiết: ' + fullLink;

        if (userId !== 'all') {
            notifyEmployeeViaTelegram(userId, extMsg);
            if (!skipZalo) {
                notifyEmployeeViaZalo(userId, extMsg);
            }
        } else {
            var allEmps = getSheetData(SHEET_NAMES.EMPLOYEES);
            allEmps.forEach(function(emp) {
                if (emp['Telegram Chat ID']) {
                    sendTelegramMessage(emp['Telegram Chat ID'], extMsg);
                }
                if (!skipZalo && emp['Zalo User ID']) {
                    sendZaloMessage(emp['Zalo User ID'], extMsg);
                }
            });
        }
    } catch (err) {
        console.error('Notification error:', err.message);
    }
}

function markNotificationRead(id) {
    const notif = findRowByColumn(SHEET_NAMES.NOTIFICATIONS, 'ID', id);
    if (notif) { updateRow(SHEET_NAMES.NOTIFICATIONS, notif._rowIndex, { 'Đã đọc': true }); }
    return { success: true };
}

function getUnreadCount(userId) {
    const data = getSheetData(SHEET_NAMES.NOTIFICATIONS)
        .filter(n => (n['ID Người nhận'] === userId || n['ID Người nhận'] === 'all')
            && n['Đã đọc'] !== true && n['Đã đọc'] !== 'TRUE');
    return { count: data.length };
}

// ============ DASHBOARD ============
function getDashboardStats(userRole, userId, userDeptId) {
    const tasks = getTasks(userRole, userId, userDeptId);
    const now = new Date();
    const thisMonth = now.getMonth();
    const thisYear = now.getFullYear();

    const totalTasks = tasks.length;
    const doneTasks = tasks.filter(t => t.status === 'done').length;
    const inProgressTasks = tasks.filter(t => t.status === 'in-progress').length;
    const todoTasks = tasks.filter(t => t.status === 'todo').length;
    const overdueTasks = tasks.filter(t => {
        return isTaskOverdueForDashboard(t, now);
    }).length;

    const completedThisMonth = tasks.filter(t => {
        if (t.status !== 'done' || !t.completed_at) return false;
        const d = new Date(t.completed_at);
        return d.getMonth() === thisMonth && d.getFullYear() === thisYear;
    }).length;

    const emps = userRole === 'Owner' ? getSheetData(SHEET_NAMES.EMPLOYEES) : [];
    const depts = userRole === 'Owner' ? getDepartments() : [];
    const cats = getCategories();

    // Chart 1: Bar - tasks by status per day (last 7 days)
    const chartWeekly = [];
    for (let i = 6; i >= 0; i--) {
        const d = new Date(); d.setDate(d.getDate() - i);
        const dayStr = d.toLocaleDateString('vi-VN', { weekday: 'short', day: 'numeric' });
        const dayTasks = tasks.filter(t => {
            const cd = new Date(t.created_at);
            return cd.toDateString() === d.toDateString();
        });
        const dayCounts = getDashboardStatusCounts(dayTasks, now);
        chartWeekly.push({
            label: dayStr,
            todo: dayCounts.todo,
            inProgress: dayCounts.inProgress,
            done: dayCounts.done,
            overdue: dayCounts.overdue
        });
    }

    // Chart 2: Doughnut - status distribution
    var chartStatus = getDashboardStatusCounts(tasks, now);

    // Chart 3: Line - completion trend (last 30 days)
    var chartCompletionTrend = [];
    for (var ci = 29; ci >= 0; ci--) {
        var cd = new Date(); cd.setDate(cd.getDate() - ci);
        var cdStr = cd.toLocaleDateString('vi-VN', { day: '2-digit', month: '2-digit' });
        var completed = tasks.filter(function(t) {
            if (t.status !== 'done' || !t.completed_at) return false;
            return new Date(t.completed_at).toDateString() === cd.toDateString();
        }).length;
        var created = tasks.filter(function(t) {
            return new Date(t.created_at).toDateString() === cd.toDateString();
        }).length;
        chartCompletionTrend.push({ label: cdStr, completed: completed, created: created });
    }

    // Chart 4: Horizontal bar - tasks by department
    var deptMap = {};
    tasks.forEach(function(t) {
        var dName = (t.department && t.department.name) ? t.department.name : 'Chưa phân bổ';
        if (!deptMap[dName]) deptMap[dName] = { total: 0, done: 0 };
        deptMap[dName].total++;
        if (t.status === 'done') deptMap[dName].done++;
    });
    var chartByDept = Object.keys(deptMap).map(function(k) {
        return { label: k, total: deptMap[k].total, done: deptMap[k].done };
    });

    // Chart 5: Polar area - tasks by priority
    var highP = tasks.filter(function(t) { return t.priority === 'high'; }).length;
    var medP = tasks.filter(function(t) { return t.priority === 'medium'; }).length;
    var lowP = tasks.filter(function(t) { return t.priority === 'low'; }).length;
    var chartByPriority = { high: highP, medium: medP, low: lowP };

    // Chart 6: Radar - performance by category
    var catMap = {};
    tasks.forEach(function(t) {
        var cName = (t.category && t.category.name) ? t.category.name : 'Không phân loại';
        if (!catMap[cName]) catMap[cName] = { total: 0, done: 0, overdue: 0 };
        catMap[cName].total++;
        if (t.status === 'done') catMap[cName].done++;
        if (isTaskOverdueForDashboard(t, now)) catMap[cName].overdue++;
    });
    var chartByCategory = Object.keys(catMap).map(function(k) {
        return { label: k, total: catMap[k].total, done: catMap[k].done, overdue: catMap[k].overdue };
    });

    return {
        totalTasks: totalTasks, doneTasks: doneTasks, inProgressTasks: inProgressTasks,
        todoTasks: todoTasks, overdueTasks: overdueTasks, completedThisMonth: completedThisMonth,
        totalEmployees: emps.length, totalDepartments: depts.length,
        chartWeekly: chartWeekly, chartStatus: chartStatus,
        chartCompletionTrend: chartCompletionTrend, chartByDept: chartByDept,
        chartByPriority: chartByPriority, chartByCategory: chartByCategory
    };
}

// ============ DEBUG ============
function debugData() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets().map(s => s.getName());

    const tasksSheet = ss.getSheetByName(SHEET_NAMES.TASKS);
    let tasksData = [];
    let tasksHeaders = [];
    if (tasksSheet) {
        const data = tasksSheet.getDataRange().getValues();
        tasksHeaders = data[0];
        tasksData = data.slice(1);
    }

    // Sheet match check
    const sheetChecks = {};
    Object.keys(SHEET_NAMES).forEach(key => {
        const expectedName = SHEET_NAMES[key];
        const exists = ss.getSheetByName(expectedName) !== null;
        sheetChecks[key] = { expected: expectedName, exists: exists };
    });

    return {
        actualSheets: sheets,
        sheetChecks: sheetChecks,
        tasksHeaders: tasksHeaders,
        tasksRowCount: tasksData.length,
        firstTask: tasksData[0] || null
    };
}

// Wrapper function for frontend - use this instead of getTasks directly
function getTasksSafe(userRole, userId, userDeptId) {
    try {
        Logger.log('getTasksSafe called with: ' + userRole + ', ' + userId + ', ' + userDeptId);
        const result = getTasks(userRole, userId, userDeptId);
        Logger.log('getTasksSafe returning: ' + (result ? result.length : 0) + ' tasks');
        // Return JSON string to avoid serialization issues
        return JSON.stringify(result || []);
    } catch (e) {
        Logger.log('getTasksSafe error: ' + e.message);
        return '[]';
    }
}

// ============ PROJECTS ============
function getProjects() {
    return getSheetData(SHEET_NAMES.PROJECTS).map(function(p) {
        return {
            id: p['ID'], name: p['Tên dự án'], description: p['Mô tả'],
            target: p['Mục tiêu'], status: p['Trạng thái'] || 'active',
            color: p['Màu sắc'] || '#6366f1',
            start_date: p['Ngày bắt đầu'], end_date: p['Ngày kết thúc'],
            target_completed: p['Đạt mục tiêu'] === true || p['Đạt mục tiêu'] === 'TRUE',
            manager_id: p['ID Người phụ trách'] || '',
            created_by: p['ID Người tạo'], created_at: p['Ngày tạo']
        };
    });
}

function createProject(data, createdBy) {
    var id = generateId('da');
    var memberIds = Array.isArray(data.member_ids) ? data.member_ids : [];
    appendRow(SHEET_NAMES.PROJECTS, {
        'ID': id, 'Tên dự án': data.name, 'Mô tả': data.description || '',
        'Mục tiêu': data.target || '', 'Trạng thái': 'active',
        'Màu sắc': data.color || '#6366f1',
        'Ngày bắt đầu': data.start_date || getLocalDateString(),
        'Ngày kết thúc': data.end_date || '',
        'Đạt mục tiêu': false,
        'ID Người phụ trách': data.manager_id || '',
        'ID Thành viên': memberIds.join(','),
        'ID Người tạo': createdBy, 'Ngày tạo': getLocalISOString()
    });
    return { success: true, id: id };
}

function updateProject(id, data) {
    var proj = findRowByColumn(SHEET_NAMES.PROJECTS, 'ID', id);
    if (!proj) return { success: false, message: 'Không tìm thấy dự án' };
    var updates = {};
    if (data.name) updates['Tên dự án'] = data.name;
    if (data.description !== undefined) updates['Mô tả'] = data.description;
    if (data.target !== undefined) updates['Mục tiêu'] = data.target;
    if (data.status) updates['Trạng thái'] = data.status;
    if (data.color) updates['Màu sắc'] = data.color;
    if (data.start_date) updates['Ngày bắt đầu'] = data.start_date;
    if (data.end_date !== undefined) updates['Ngày kết thúc'] = data.end_date;
    if (data.target_completed !== undefined) updates['Đạt mục tiêu'] = data.target_completed;
    if (data.manager_id !== undefined) updates['ID Người phụ trách'] = data.manager_id;
    if (data.member_ids !== undefined) updates['ID Thành viên'] = Array.isArray(data.member_ids) ? data.member_ids.join(',') : data.member_ids;
    updateRow(SHEET_NAMES.PROJECTS, proj._rowIndex, updates);
    return { success: true };
}

function deleteProject(id) {
    var proj = findRowByColumn(SHEET_NAMES.PROJECTS, 'ID', id);
    if (proj) { deleteRow(SHEET_NAMES.PROJECTS, proj._rowIndex); return { success: true }; }
    return { success: false };
}

function toggleProjectTarget(id, completed) {
    var proj = findRowByColumn(SHEET_NAMES.PROJECTS, 'ID', id);
    if (!proj) return { success: false };
    var updates = { 'Đạt mục tiêu': completed };
    if (completed) updates['Trạng thái'] = 'completed';
    updateRow(SHEET_NAMES.PROJECTS, proj._rowIndex, updates);
    return { success: true };
}

// ============ SETUP ALL SHEETS ============
// Helper: create or reset a sheet with styled headers
function _createSheet(ss, name, headers) {
    var sheet = ss.getSheetByName(name);
    if (sheet) {
        // Clear existing data but keep the sheet
        sheet.clear();
    } else {
        sheet = ss.insertSheet(name);
    }
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#4a86e8')
        .setFontColor('#ffffff')
        .setFontSize(10);
    sheet.setFrozenRows(1);
    // Auto-resize columns
    for (var i = 1; i <= headers.length; i++) {
        sheet.setColumnWidth(i, 150);
    }
    return sheet;
}

// Helper: bulk write rows to a sheet
function _bulkWrite(sheet, headers, dataRows) {
    if (dataRows.length === 0) return;
    var rows = dataRows.map(function(obj) {
        return headers.map(function(h) { return obj[h] !== undefined ? obj[h] : ''; });
    });
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

// Helper: date string relative to now
function _d(daysOffset) {
    var d = new Date();
    d.setDate(d.getDate() + daysOffset);
    return getLocalDateString(d); // YYYY-MM-DD
}
function _iso(daysOffset, hoursOffset) {
    var d = new Date();
    d.setDate(d.getDate() + daysOffset);
    if (hoursOffset) d.setHours(d.getHours() + hoursOffset);
    return getLocalISOString(d);
}

function setupAllSheets() {
    var ss = getSS();

    // ======== 1. CÀI ĐẶT ========
    var settingsHeaders = ['Tên công ty', 'Mô tả', 'Tên đăng nhập', 'Mật khẩu', 'Họ tên', 'Email', 'Điện thoại', 'Ảnh đại diện', 'Telegram Bot Token'];
    var settingsSheet = _createSheet(ss, SHEET_NAMES.SETTINGS, settingsHeaders);
    _bulkWrite(settingsSheet, settingsHeaders, [
        { 'Tên công ty': 'Công ty TNHH coogo.vn', 'Mô tả': 'Ứng dụng quản lý công việc thông minh', 'Tên đăng nhập': 'admin', 'Mật khẩu': '123', 'Họ tên': 'Nguyễn Văn An', 'Email': 'admin@coogo.vn.com', 'Điện thoại': '0901234567', 'Ảnh đại diện': '' }
    ]);

    // ======== 2. PHÒNG BAN ========
    var deptHeaders = ['ID', 'Tên phòng ban', 'Mô tả', 'Ngày tạo'];
    var deptSheet = _createSheet(ss, SHEET_NAMES.DEPARTMENTS, deptHeaders);
    var depts = [
        { 'ID': 'pb_001', 'Tên phòng ban': 'Ban Kỹ thuật', 'Mô tả': 'Phụ trách phát triển phần mềm và hạ tầng kỹ thuật', 'Ngày tạo': _iso(-60) },
        { 'ID': 'pb_002', 'Tên phòng ban': 'Ban Marketing', 'Mô tả': 'Phụ trách tiếp thị, truyền thông và quảng cáo', 'Ngày tạo': _iso(-60) },
        { 'ID': 'pb_003', 'Tên phòng ban': 'Ban Nhân sự', 'Mô tả': 'Phụ trách tuyển dụng, đào tạo và phúc lợi', 'Ngày tạo': _iso(-55) },
        { 'ID': 'pb_004', 'Tên phòng ban': 'Ban Kinh doanh', 'Mô tả': 'Phụ trách bán hàng và chăm sóc khách hàng', 'Ngày tạo': _iso(-50) }
    ];
    _bulkWrite(deptSheet, deptHeaders, depts);

    // ======== 3. NHÂN VIÊN ========
    var empHeaders = ['ID', 'Họ tên', 'Tên đăng nhập', 'Mật khẩu', 'Email', 'Điện thoại', 'Ảnh đại diện', 'Vai trò', 'ID Phòng ban', 'Ngày tạo', 'Telegram Chat ID'];
    var empSheet = _createSheet(ss, SHEET_NAMES.EMPLOYEES, empHeaders);
    var emps = [
        { 'ID': 'nv_001', 'Họ tên': 'Trần Minh Tuấn', 'Tên đăng nhập': 'tuan', 'Mật khẩu': '123', 'Email': 'tuan@coogo.vn.com', 'Điện thoại': '0912345001', 'Ảnh đại diện': '', 'Vai trò': 'Manager', 'ID Phòng ban': 'pb_001', 'Ngày tạo': _iso(-58) },
        { 'ID': 'nv_002', 'Họ tên': 'Lê Thị Hương', 'Tên đăng nhập': 'huong', 'Mật khẩu': '123', 'Email': 'huong@coogo.vn.com', 'Điện thoại': '0912345002', 'Ảnh đại diện': '', 'Vai trò': 'Member', 'ID Phòng ban': 'pb_001', 'Ngày tạo': _iso(-55) },
        { 'ID': 'nv_003', 'Họ tên': 'Phạm Quốc Bảo', 'Tên đăng nhập': 'bao', 'Mật khẩu': '123', 'Email': 'bao@coogo.vn.com', 'Điện thoại': '0912345003', 'Ảnh đại diện': '', 'Vai trò': 'Member', 'ID Phòng ban': 'pb_001', 'Ngày tạo': _iso(-50) },
        { 'ID': 'nv_004', 'Họ tên': 'Nguyễn Thị Lan', 'Tên đăng nhập': 'lan', 'Mật khẩu': '123', 'Email': 'lan@coogo.vn.com', 'Điện thoại': '0912345004', 'Ảnh đại diện': '', 'Vai trò': 'Manager', 'ID Phòng ban': 'pb_002', 'Ngày tạo': _iso(-56) },
        { 'ID': 'nv_005', 'Họ tên': 'Hoàng Đức Anh', 'Tên đăng nhập': 'anh', 'Mật khẩu': '123', 'Email': 'anh@coogo.vn.com', 'Điện thoại': '0912345005', 'Ảnh đại diện': '', 'Vai trò': 'Member', 'ID Phòng ban': 'pb_002', 'Ngày tạo': _iso(-48) },
        { 'ID': 'nv_006', 'Họ tên': 'Vũ Thanh Mai', 'Tên đăng nhập': 'mai', 'Mật khẩu': '123', 'Email': 'mai@coogo.vn.com', 'Điện thoại': '0912345006', 'Ảnh đại diện': '', 'Vai trò': 'Manager', 'ID Phòng ban': 'pb_003', 'Ngày tạo': _iso(-54) },
        { 'ID': 'nv_007', 'Họ tên': 'Đỗ Văn Khải', 'Tên đăng nhập': 'khai', 'Mật khẩu': '123', 'Email': 'khai@coogo.vn.com', 'Điện thoại': '0912345007', 'Ảnh đại diện': '', 'Vai trò': 'Member', 'ID Phòng ban': 'pb_003', 'Ngày tạo': _iso(-45) },
        { 'ID': 'nv_008', 'Họ tên': 'Bùi Kim Ngân', 'Tên đăng nhập': 'ngan', 'Mật khẩu': '123', 'Email': 'ngan@coogo.vn.com', 'Điện thoại': '0912345008', 'Ảnh đại diện': '', 'Vai trò': 'Manager', 'ID Phòng ban': 'pb_004', 'Ngày tạo': _iso(-52) }
    ];
    _bulkWrite(empSheet, empHeaders, emps);

    // ======== 4. DANH MỤC ========
    var catHeaders = ['ID', 'Tên danh mục', 'Icon', 'Màu sắc', 'Mặc định', 'Ngày tạo'];
    var catSheet = _createSheet(ss, SHEET_NAMES.CATEGORIES, catHeaders);
    var cats = [
        { 'ID': 'dm_001', 'Tên danh mục': 'Phát triển', 'Icon': 'code', 'Màu sắc': '#6366f1', 'Mặc định': true, 'Ngày tạo': _iso(-60) },
        { 'ID': 'dm_002', 'Tên danh mục': 'Thiết kế', 'Icon': 'palette', 'Màu sắc': '#ec4899', 'Mặc định': false, 'Ngày tạo': _iso(-58) },
        { 'ID': 'dm_003', 'Tên danh mục': 'Marketing', 'Icon': 'megaphone', 'Màu sắc': '#f59e0b', 'Mặc định': false, 'Ngày tạo': _iso(-55) },
        { 'ID': 'dm_004', 'Tên danh mục': 'Vận hành', 'Icon': 'cog', 'Màu sắc': '#10b981', 'Mặc định': false, 'Ngày tạo': _iso(-50) },
        { 'ID': 'dm_005', 'Tên danh mục': 'Nghiên cứu', 'Icon': 'search', 'Màu sắc': '#3b82f6', 'Mặc định': false, 'Ngày tạo': _iso(-48) },
        { 'ID': 'dm_006', 'Tên danh mục': 'Hỗ trợ', 'Icon': 'headset', 'Màu sắc': '#8b5cf6', 'Mặc định': false, 'Ngày tạo': _iso(-45) }
    ];
    _bulkWrite(catSheet, catHeaders, cats);

    // ======== 5. DỰ ÁN ========
    var projHeaders = ['ID', 'Tên dự án', 'Mô tả', 'Mục tiêu', 'Trạng thái', 'Màu sắc', 'Ngày bắt đầu', 'Ngày kết thúc', 'Đạt mục tiêu', 'ID Người phụ trách', 'ID Thành viên', 'ID Người tạo', 'Ngày tạo'];
    var projSheet = _createSheet(ss, SHEET_NAMES.PROJECTS, projHeaders);
    var projs = [
        { 'ID': 'da_001', 'Tên dự án': 'Website Thương mại điện tử', 'Mô tả': 'Xây dựng hệ thống website bán hàng trực tuyến đầy đủ tính năng.', 'Mục tiêu': 'Hoàn thành MVP, đạt 1000 đơn hàng/tháng', 'Trạng thái': 'active', 'Màu sắc': '#6366f1', 'Ngày bắt đầu': _d(-45), 'Ngày kết thúc': _d(60), 'Đạt mục tiêu': false, 'ID Người tạo': 'owner', 'Ngày tạo': _iso(-45) },
        { 'ID': 'da_002', 'Tên dự án': 'App Quản lý Nhân sự', 'Mô tả': 'Ứng dụng quản lý nhân sự tích hợp chấm công GPS, nghỉ phép và lương.', 'Mục tiêu': 'Launch phiên bản 1.0 trước tháng 5/2026', 'Trạng thái': 'active', 'Màu sắc': '#10b981', 'Ngày bắt đầu': _d(-30), 'Ngày kết thúc': _d(45), 'Đạt mục tiêu': false, 'ID Người tạo': 'owner', 'Ngày tạo': _iso(-30) },
        { 'ID': 'da_003', 'Tên dự án': 'Chiến dịch Marketing Q1', 'Mô tả': 'Triển khai quảng cáo MXH, Google Ads, email marketing tăng nhận diện.', 'Mục tiêu': 'Đạt 50,000 lượt tiếp cận và 500 leads mới', 'Trạng thái': 'completed', 'Màu sắc': '#f59e0b', 'Ngày bắt đầu': _d(-60), 'Ngày kết thúc': _d(-1), 'Đạt mục tiêu': true, 'ID Người tạo': 'owner', 'Ngày tạo': _iso(-60) },
        { 'ID': 'da_004', 'Tên dự án': 'Nâng cấp Hạ tầng Server', 'Mô tả': 'Di chuyển hệ thống lên cloud, tối ưu hiệu suất và bảo mật.', 'Mục tiêu': 'Giảm downtime xuống 0.1%, tăng tốc độ 50%', 'Trạng thái': 'active', 'Màu sắc': '#ef4444', 'Ngày bắt đầu': _d(-20), 'Ngày kết thúc': _d(30), 'Đạt mục tiêu': false, 'ID Người tạo': 'owner', 'Ngày tạo': _iso(-20) },
        { 'ID': 'da_005', 'Tên dự án': 'Đào tạo Nội bộ 2026', 'Mô tả': 'Chương trình đào tạo kỹ năng mềm và chuyên môn cho nhân viên.', 'Mục tiêu': '100% nhân viên hoàn thành khóa học', 'Trạng thái': 'active', 'Màu sắc': '#8b5cf6', 'Ngày bắt đầu': _d(-15), 'Ngày kết thúc': _d(90), 'Đạt mục tiêu': false, 'ID Người tạo': 'owner', 'Ngày tạo': _iso(-15) }
    ];
    _bulkWrite(projSheet, projHeaders, projs);

    // ======== 6. NHIỆM VỤ ========
    var taskHeaders = ['ID', 'Tiêu đề', 'Mô tả', 'ID Người thực hiện', 'ID Phòng ban', 'ID Danh mục', 'ID Dự án', 'Độ ưu tiên', 'Trạng thái', 'Ngày bắt đầu', 'Hạn hoàn thành', 'Ngày hoàn thành', 'ID Người tạo', 'Ngày tạo'];
    var taskSheet = _createSheet(ss, SHEET_NAMES.TASKS, taskHeaders);
    var tasks = [
        // DA 1 - Website TMDT (Kỹ thuật)
        { 'ID': 'nv_t01', 'Tiêu đề': 'Thiết kế giao diện trang chủ', 'Mô tả': 'Thiết kế UI/UX cho trang chủ website, bao gồm banner, sản phẩm nổi bật, footer.', 'ID Người thực hiện': 'nv_002', 'ID Phòng ban': 'pb_001', 'ID Danh mục': 'dm_002', 'ID Dự án': 'da_001', 'Độ ưu tiên': 'high', 'Trạng thái': 'done', 'Ngày bắt đầu': _d(-14), 'Hạn hoàn thành': _d(-5), 'Ngày hoàn thành': _iso(-6), 'ID Người tạo': 'owner', 'Ngày tạo': _iso(-14) },
        { 'ID': 'nv_t02', 'Tiêu đề': 'Xây dựng API giỏ hàng', 'Mô tả': 'Backend REST API cho chức năng giỏ hàng: thêm, xóa, cập nhật số lượng.', 'ID Người thực hiện': 'nv_001', 'ID Phòng ban': 'pb_001', 'ID Danh mục': 'dm_001', 'ID Dự án': 'da_001', 'Độ ưu tiên': 'high', 'Trạng thái': 'in-progress', 'Ngày bắt đầu': _d(-10), 'Hạn hoàn thành': _d(2), 'Ngày hoàn thành': '', 'ID Người tạo': 'owner', 'Ngày tạo': _iso(-10) },
        { 'ID': 'nv_t03', 'Tiêu đề': 'Tích hợp cổng thanh toán VNPay', 'Mô tả': 'Kết nối VNPay sandbox, xử lý callback, lưu giao dịch.', 'ID Người thực hiện': 'nv_003', 'ID Phòng ban': 'pb_001', 'ID Danh mục': 'dm_001', 'ID Dự án': 'da_001', 'Độ ưu tiên': 'high', 'Trạng thái': 'todo', 'Ngày bắt đầu': _d(-7), 'Hạn hoàn thành': _d(7), 'Ngày hoàn thành': '', 'ID Người tạo': 'owner', 'Ngày tạo': _iso(-7) },
        { 'ID': 'nv_t04', 'Tiêu đề': 'Viết unit test module đơn hàng', 'Mô tả': 'Viết test cho các chức năng tạo, cập nhật, hủy đơn hàng.', 'ID Người thực hiện': 'nv_002', 'ID Phòng ban': 'pb_001', 'ID Danh mục': 'dm_001', 'ID Dự án': 'da_001', 'Độ ưu tiên': 'medium', 'Trạng thái': 'todo', 'Ngày bắt đầu': _d(-5), 'Hạn hoàn thành': _d(10), 'Ngày hoàn thành': '', 'ID Người tạo': 'nv_001', 'Ngày tạo': _iso(-5) },

        // DA 2 - App Nhân sự (Kỹ thuật + Nhân sự)
        { 'ID': 'nv_t05', 'Tiêu đề': 'Thiết kế database schema nhân sự', 'Mô tả': 'Xây dựng ERD và SQL schema cho module nhân viên, phòng ban, chức vụ.', 'ID Người thực hiện': 'nv_001', 'ID Phòng ban': 'pb_001', 'ID Danh mục': 'dm_001', 'ID Dự án': 'da_002', 'Độ ưu tiên': 'high', 'Trạng thái': 'done', 'Ngày bắt đầu': _d(-12), 'Hạn hoàn thành': _d(-3), 'Ngày hoàn thành': _iso(-4), 'ID Người tạo': 'owner', 'Ngày tạo': _iso(-12) },
        { 'ID': 'nv_t06', 'Tiêu đề': 'Phát triển module chấm công GPS', 'Mô tả': 'Tính năng chấm công bằng GPS, xác minh vị trí văn phòng.', 'ID Người thực hiện': 'nv_003', 'ID Phòng ban': 'pb_001', 'ID Danh mục': 'dm_001', 'ID Dự án': 'da_002', 'Độ ưu tiên': 'high', 'Trạng thái': 'in-progress', 'Ngày bắt đầu': _d(-8), 'Hạn hoàn thành': _d(5), 'Ngày hoàn thành': '', 'ID Người tạo': 'owner', 'Ngày tạo': _iso(-8) },
        { 'ID': 'nv_t07', 'Tiêu đề': 'Thu thập yêu cầu nghỉ phép', 'Mô tả': 'Phỏng vấn HR để xác định yêu cầu nghiệp vụ module quản lý nghỉ phép.', 'ID Người thực hiện': 'nv_006', 'ID Phòng ban': 'pb_003', 'ID Danh mục': 'dm_005', 'ID Dự án': 'da_002', 'Độ ưu tiên': 'medium', 'Trạng thái': 'done', 'Ngày bắt đầu': _d(-15), 'Hạn hoàn thành': _d(-7), 'Ngày hoàn thành': _iso(-8), 'ID Người tạo': 'owner', 'Ngày tạo': _iso(-15) },
        { 'ID': 'nv_t08', 'Tiêu đề': 'Tạo landing page giới thiệu app', 'Mô tả': 'Thiết kế và code trang giới thiệu ứng dụng HRMS.', 'ID Người thực hiện': 'nv_002', 'ID Phòng ban': 'pb_001', 'ID Danh mục': 'dm_002', 'ID Dự án': 'da_002', 'Độ ưu tiên': 'low', 'Trạng thái': 'todo', 'Ngày bắt đầu': _d(-3), 'Hạn hoàn thành': _d(14), 'Ngày hoàn thành': '', 'ID Người tạo': 'nv_001', 'Ngày tạo': _iso(-3) },

        // DA 3 - Marketing Q1 (Marketing - completed)
        { 'ID': 'nv_t09', 'Tiêu đề': 'Chạy quảng cáo Facebook Ads', 'Mô tả': 'Setup và tối ưu campaign Facebook Ads cho sản phẩm mới.', 'ID Người thực hiện': 'nv_005', 'ID Phòng ban': 'pb_002', 'ID Danh mục': 'dm_003', 'ID Dự án': 'da_003', 'Độ ưu tiên': 'high', 'Trạng thái': 'done', 'Ngày bắt đầu': _d(-25), 'Hạn hoàn thành': _d(-10), 'Ngày hoàn thành': _iso(-11), 'ID Người tạo': 'nv_004', 'Ngày tạo': _iso(-25) },
        { 'ID': 'nv_t10', 'Tiêu đề': 'Viết content SEO blog', 'Mô tả': 'Viết 10 bài blog tối ưu SEO cho website công ty.', 'ID Người thực hiện': 'nv_004', 'ID Phòng ban': 'pb_002', 'ID Danh mục': 'dm_003', 'ID Dự án': 'da_003', 'Độ ưu tiên': 'medium', 'Trạng thái': 'done', 'Ngày bắt đầu': _d(-20), 'Hạn hoàn thành': _d(-8), 'Ngày hoàn thành': _iso(-9), 'ID Người tạo': 'owner', 'Ngày tạo': _iso(-20) },
        { 'ID': 'nv_t11', 'Tiêu đề': 'Phân tích báo cáo chiến dịch', 'Mô tả': 'Tổng hợp và phân tích KPI chiến dịch Marketing Q1.', 'ID Người thực hiện': 'nv_004', 'ID Phòng ban': 'pb_002', 'ID Danh mục': 'dm_005', 'ID Dự án': 'da_003', 'Độ ưu tiên': 'medium', 'Trạng thái': 'done', 'Ngày bắt đầu': _d(-10), 'Hạn hoàn thành': _d(-2), 'Ngày hoàn thành': _iso(-3), 'ID Người tạo': 'owner', 'Ngày tạo': _iso(-10) },

        // DA 4 - Server (Kỹ thuật)
        { 'ID': 'nv_t12', 'Tiêu đề': 'Đánh giá nhà cung cấp cloud', 'Mô tả': 'So sánh AWS, GCP, Azure về giá, hiệu suất và hỗ trợ khu vực APAC.', 'ID Người thực hiện': 'nv_001', 'ID Phòng ban': 'pb_001', 'ID Danh mục': 'dm_005', 'ID Dự án': 'da_004', 'Độ ưu tiên': 'high', 'Trạng thái': 'done', 'Ngày bắt đầu': _d(-15), 'Hạn hoàn thành': _d(-4), 'Ngày hoàn thành': _iso(-5), 'ID Người tạo': 'owner', 'Ngày tạo': _iso(-15) },
        { 'ID': 'nv_t13', 'Tiêu đề': 'Migrate database lên cloud', 'Mô tả': 'Di chuyển PostgreSQL từ on-premises lên RDS, kiểm tra tính toàn vẹn dữ liệu.', 'ID Người thực hiện': 'nv_003', 'ID Phòng ban': 'pb_001', 'ID Danh mục': 'dm_004', 'ID Dự án': 'da_004', 'Độ ưu tiên': 'high', 'Trạng thái': 'in-progress', 'Ngày bắt đầu': _d(-7), 'Hạn hoàn thành': _d(3), 'Ngày hoàn thành': '', 'ID Người tạo': 'owner', 'Ngày tạo': _iso(-7) },
        { 'ID': 'nv_t14', 'Tiêu đề': 'Thiết lập CI/CD pipeline', 'Mô tả': 'Cấu hình GitHub Actions cho auto deploy staging/production.', 'ID Người thực hiện': 'nv_001', 'ID Phòng ban': 'pb_001', 'ID Danh mục': 'dm_004', 'ID Dự án': 'da_004', 'Độ ưu tiên': 'medium', 'Trạng thái': 'todo', 'Ngày bắt đầu': _d(-4), 'Hạn hoàn thành': _d(8), 'Ngày hoàn thành': '', 'ID Người tạo': 'nv_001', 'Ngày tạo': _iso(-4) },
        { 'ID': 'nv_t15', 'Tiêu đề': 'Cấu hình monitoring & alerting', 'Mô tả': 'Setup Grafana + Prometheus cho giám sát server, cảnh báo qua Slack.', 'ID Người thực hiện': 'nv_002', 'ID Phòng ban': 'pb_001', 'ID Danh mục': 'dm_004', 'ID Dự án': 'da_004', 'Độ ưu tiên': 'medium', 'Trạng thái': 'todo', 'Ngày bắt đầu': _d(-2), 'Hạn hoàn thành': _d(12), 'Ngày hoàn thành': '', 'ID Người tạo': 'nv_001', 'Ngày tạo': _iso(-2) },

        // DA 5 - Đào tạo (Nhân sự)
        { 'ID': 'nv_t16', 'Tiêu đề': 'Lập kế hoạch đào tạo Q2', 'Mô tả': 'Xây dựng lộ trình đào tạo cho toàn bộ nhân viên Q2/2026.', 'ID Người thực hiện': 'nv_006', 'ID Phòng ban': 'pb_003', 'ID Danh mục': 'dm_004', 'ID Dự án': 'da_005', 'Độ ưu tiên': 'high', 'Trạng thái': 'in-progress', 'Ngày bắt đầu': _d(-10), 'Hạn hoàn thành': _d(4), 'Ngày hoàn thành': '', 'ID Người tạo': 'owner', 'Ngày tạo': _iso(-10) },
        { 'ID': 'nv_t17', 'Tiêu đề': 'Tổ chức workshop Agile/Scrum', 'Mô tả': 'Workshop 2 ngày cho các team, mời chuyên gia bên ngoài.', 'ID Người thực hiện': 'nv_007', 'ID Phòng ban': 'pb_003', 'ID Danh mục': 'dm_006', 'ID Dự án': 'da_005', 'Độ ưu tiên': 'medium', 'Trạng thái': 'todo', 'Ngày bắt đầu': _d(-5), 'Hạn hoàn thành': _d(15), 'Ngày hoàn thành': '', 'ID Người tạo': 'nv_006', 'Ngày tạo': _iso(-5) },

        // Không thuộc dự án
        { 'ID': 'nv_t18', 'Tiêu đề': 'Xử lý ticket hỗ trợ khách hàng', 'Mô tả': 'Giải quyết 15 ticket support từ tuần trước, ưu tiên ticket urgent.', 'ID Người thực hiện': 'nv_008', 'ID Phòng ban': 'pb_004', 'ID Danh mục': 'dm_006', 'ID Dự án': '', 'Độ ưu tiên': 'high', 'Trạng thái': 'in-progress', 'Ngày bắt đầu': _d(-3), 'Hạn hoàn thành': _d(1), 'Ngày hoàn thành': '', 'ID Người tạo': 'owner', 'Ngày tạo': _iso(-3) },
        { 'ID': 'nv_t19', 'Tiêu đề': 'Chuẩn bị báo cáo doanh thu tháng 3', 'Mô tả': 'Tổng hợp doanh thu theo sản phẩm, kênh bán hàng, so sánh vs target.', 'ID Người thực hiện': 'nv_008', 'ID Phòng ban': 'pb_004', 'ID Danh mục': 'dm_004', 'ID Dự án': '', 'Độ ưu tiên': 'medium', 'Trạng thái': 'todo', 'Ngày bắt đầu': _d(-2), 'Hạn hoàn thành': _d(3), 'Ngày hoàn thành': '', 'ID Người tạo': 'owner', 'Ngày tạo': _iso(-2) },
        { 'ID': 'nv_t20', 'Tiêu đề': 'Cập nhật tài liệu onboarding', 'Mô tả': 'Cập nhật handbook cho nhân viên mới, thêm quy trình mới.', 'ID Người thực hiện': 'nv_007', 'ID Phòng ban': 'pb_003', 'ID Danh mục': 'dm_004', 'ID Dự án': '', 'Độ ưu tiên': 'low', 'Trạng thái': 'in-progress', 'Ngày bắt đầu': _d(-4), 'Hạn hoàn thành': _d(6), 'Ngày hoàn thành': '', 'ID Người tạo': 'nv_006', 'Ngày tạo': _iso(-4) },
        { 'ID': 'nv_t21', 'Tiêu đề': 'Thiết kế banner quảng cáo Q2', 'Mô tả': 'Tạo bộ banner cho chiến dịch Q2 trên Facebook, Google, Zalo.', 'ID Người thực hiện': 'nv_005', 'ID Phòng ban': 'pb_002', 'ID Danh mục': 'dm_002', 'ID Dự án': '', 'Độ ưu tiên': 'medium', 'Trạng thái': 'todo', 'Ngày bắt đầu': _d(-1), 'Hạn hoàn thành': _d(9), 'Ngày hoàn thành': '', 'ID Người tạo': 'nv_004', 'Ngày tạo': _iso(-1) },
        { 'ID': 'nv_t22', 'Tiêu đề': 'Review code sprint 15', 'Mô tả': 'Review và merge các pull requests từ sprint 15.', 'ID Người thực hiện': 'nv_001', 'ID Phòng ban': 'pb_001', 'ID Danh mục': 'dm_001', 'ID Dự án': 'da_001', 'Độ ưu tiên': 'medium', 'Trạng thái': 'in-progress', 'Ngày bắt đầu': _d(-1), 'Hạn hoàn thành': _d(1), 'Ngày hoàn thành': '', 'ID Người tạo': 'nv_001', 'Ngày tạo': _iso(-1) },
        { 'ID': 'nv_t23', 'Tiêu đề': 'Tối ưu SEO trang sản phẩm', 'Mô tả': 'Cải thiện meta tags, schema markup, tốc độ tải trang sản phẩm.', 'ID Người thực hiện': 'nv_005', 'ID Phòng ban': 'pb_002', 'ID Danh mục': 'dm_003', 'ID Dự án': 'da_001', 'Độ ưu tiên': 'low', 'Trạng thái': 'todo', 'Ngày bắt đầu': _d(0), 'Hạn hoàn thành': _d(11), 'Ngày hoàn thành': '', 'ID Người tạo': 'nv_004', 'Ngày tạo': _iso(0) }
    ];
    _bulkWrite(taskSheet, taskHeaders, tasks);

    // ======== 7. BÌNH LUẬN ========
    var cmtHeaders = ['ID', 'ID Nhiệm vụ', 'ID Người bình luận', 'Nội dung', 'Hình ảnh', 'Ngày tạo'];
    var cmtSheet = _createSheet(ss, SHEET_NAMES.COMMENTS, cmtHeaders);
    var cmts = [
        { 'ID': 'bl_001', 'ID Nhiệm vụ': 'nv_t01', 'ID Người bình luận': 'nv_001', 'Nội dung': 'Giao diện đẹp lắm, approve!', 'Hình ảnh': '', 'Ngày tạo': _iso(-6) },
        { 'ID': 'bl_002', 'ID Nhiệm vụ': 'nv_t02', 'ID Người bình luận': 'owner', 'Nội dung': 'Nhớ xử lý edge case khi giỏ hàng trống nhé.', 'Hình ảnh': '', 'Ngày tạo': _iso(-5) },
        { 'ID': 'bl_003', 'ID Nhiệm vụ': 'nv_t02', 'ID Người bình luận': 'nv_001', 'Nội dung': 'Đã xử lý, đang test thêm với số lượng lớn.', 'Hình ảnh': '', 'Ngày tạo': _iso(-4) },
        { 'ID': 'bl_004', 'ID Nhiệm vụ': 'nv_t06', 'ID Người bình luận': 'nv_003', 'Nội dung': 'GPS có sai số ~10m, cần hỏi lại bán kính chấp nhận.', 'Hình ảnh': '', 'Ngày tạo': _iso(-3) },
        { 'ID': 'bl_005', 'ID Nhiệm vụ': 'nv_t06', 'ID Người bình luận': 'owner', 'Nội dung': 'Bán kính 50m là ok nhé.', 'Hình ảnh': '', 'Ngày tạo': _iso(-2) },
        { 'ID': 'bl_006', 'ID Nhiệm vụ': 'nv_t13', 'ID Người bình luận': 'nv_001', 'Nội dung': 'Đã backup xong, đang tiến hành migrate. ETA 2 ngày.', 'Hình ảnh': '', 'Ngày tạo': _iso(-1) },
        { 'ID': 'bl_007', 'ID Nhiệm vụ': 'nv_t18', 'ID Người bình luận': 'nv_008', 'Nội dung': 'Đã xử lý 8/15 tickets, còn lại đang chờ phản hồi KH.', 'Hình ảnh': '', 'Ngày tạo': _iso(-1) }
    ];
    _bulkWrite(cmtSheet, cmtHeaders, cmts);

    // ======== 8. THÔNG BÁO ========
    var notifHeaders = ['ID', 'ID Người nhận', 'Tiêu đề', 'Nội dung', 'Đường dẫn', 'Đã đọc', 'Ngày tạo'];
    var notifSheet = _createSheet(ss, SHEET_NAMES.NOTIFICATIONS, notifHeaders);
    var notifs = [
        { 'ID': 'tb_001', 'ID Người nhận': 'all', 'Tiêu đề': 'Chào mừng đến với coogo.vn!', 'Nội dung': 'Hệ thống quản lý công việc đã được thiết lập thành công.', 'Đường dẫn': '/dashboard', 'Đã đọc': false, 'Ngày tạo': _iso(-14) },
        { 'ID': 'tb_002', 'ID Người nhận': 'nv_001', 'Tiêu đề': 'Nhiệm vụ mới được giao', 'Nội dung': 'Bạn được giao nhiệm vụ: Xây dựng API giỏ hàng', 'Đường dẫn': '/progress', 'Đã đọc': true, 'Ngày tạo': _iso(-10) },
        { 'ID': 'tb_003', 'ID Người nhận': 'nv_003', 'Tiêu đề': 'Nhiệm vụ mới được giao', 'Nội dung': 'Bạn được giao nhiệm vụ: Tích hợp cổng thanh toán VNPay', 'Đường dẫn': '/progress', 'Đã đọc': false, 'Ngày tạo': _iso(-7) },
        { 'ID': 'tb_004', 'ID Người nhận': 'nv_002', 'Tiêu đề': 'Cập nhật trạng thái', 'Nội dung': 'Nhiệm vụ "Thiết kế giao diện trang chủ" đã chuyển sang: Hoàn thành', 'Đường dẫn': '/progress', 'Đã đọc': true, 'Ngày tạo': _iso(-6) },
        { 'ID': 'tb_005', 'ID Người nhận': 'nv_006', 'Tiêu đề': 'Nhiệm vụ mới được giao', 'Nội dung': 'Bạn được giao nhiệm vụ: Lập kế hoạch đào tạo Q2', 'Đường dẫn': '/progress', 'Đã đọc': false, 'Ngày tạo': _iso(-5) },
        { 'ID': 'tb_006', 'ID Người nhận': 'nv_008', 'Tiêu đề': 'Nhiệm vụ cần gấp', 'Nội dung': 'Bạn được giao nhiệm vụ ưu tiên cao: Xử lý ticket hỗ trợ khách hàng', 'Đường dẫn': '/progress', 'Đã đọc': false, 'Ngày tạo': _iso(-3) },
        { 'ID': 'tb_007', 'ID Người nhận': 'all', 'Tiêu đề': 'Thông báo chung', 'Nội dung': 'Chiến dịch Marketing Q1 đã hoàn thành và đạt mục tiêu. Chúc mừng team Marketing!', 'Đường dẫn': '/projects', 'Đã đọc': false, 'Ngày tạo': _iso(-1) }
    ];
    _bulkWrite(notifSheet, notifHeaders, notifs);

    return { success: true, message: 'Đã thiết lập toàn bộ ' + Object.keys(SHEET_NAMES).length + ' sheets với dữ liệu mẫu thành công!' };
}

// Keep backward compatibility
function setupProjectsSheet() {
    return setupAllSheets();
}

// ============ ACTIVITY LOG ============
function logActivity(taskId, userId, action, details) {
    var sheet = getSheet('Nhật Ký');
    if (!sheet) {
        var ss = getSS();
        var headers = ['ID', 'ID Nhiệm vụ', 'ID Người thực hiện', 'Hành động', 'Chi tiết', 'Ngày tạo'];
        sheet = ss.insertSheet('Nhật Ký');
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4a86e8').setFontColor('#ffffff');
        sheet.setFrozenRows(1);
    }
    var id = generateId('log');
    sheet.appendRow([id, taskId, userId, action, details || '', getLocalISOString()]);
}

function getActivityLog(taskId) {
    var sheet = getSheet('Nhật Ký');
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    var headers = data[0].map(function(h) { return String(h).trim(); });
    var emps = getSheetData(SHEET_NAMES.EMPLOYEES);
    var settingsData = getSheetData(SHEET_NAMES.SETTINGS);
    var ownerName = settingsData.length > 0 ? settingsData[0]['Họ tên'] : 'Admin';
    
    return data.slice(1).filter(function(row) {
        return row[headers.indexOf('ID Nhiệm vụ')] === taskId;
    }).map(function(row) {
        var userId = row[headers.indexOf('ID Người thực hiện')];
        var userName = userId === 'owner' ? ownerName : '';
        if (!userName) {
            var emp = emps.find(function(e) { return e['ID'] === userId; });
            if (emp) userName = emp['Họ tên'];
        }
        return {
            id: row[headers.indexOf('ID')],
            task_id: taskId,
            user_id: userId,
            user_name: userName || 'Hệ thống',
            action: row[headers.indexOf('Hành động')],
            details: row[headers.indexOf('Chi tiết')],
            created_at: row[headers.indexOf('Ngày tạo')]
        };
    }).sort(function(a, b) { return new Date(b.created_at) - new Date(a.created_at); });
}

// ============ TELEGRAM NOTIFICATION ============
function getTelegramBotToken() {
    var data = getSheetData(SHEET_NAMES.SETTINGS);
    if (data && data.length > 0) {
        return data[0]['Telegram Bot Token'] || '';
    }
    return '';
}

function sendTelegramMessage(chatId, message) {
    var token = getTelegramBotToken();
    if (!token) return 'Lỗi: Chưa có Bot Token';
    if (!chatId) return 'Lỗi: Chưa có Chat ID';
    
    try {
        var url = 'https://api.telegram.org/bot' + token + '/sendMessage';
        var payload = {
            'chat_id': String(chatId),
            'text': message,
            'parse_mode': 'HTML',
            'disable_web_page_preview': true
        };
        
        var options = {
            'method': 'post',
            'contentType': 'application/json',
            'payload': JSON.stringify(payload),
            'muteHttpExceptions': true
        };
        
        var response = UrlFetchApp.fetch(url, options);
        var result = JSON.parse(response.getContentText());
        
        if (!result.ok) {
            console.error('Telegram error:', result.description);
            return 'Telegram Error: ' + result.description;
        }
        return 'Thành công';
    } catch (e) {
        console.error('sendTelegramMessage error:', e.message);
        return 'Lỗi: ' + e.message;
    }
}

function notifyEmployeeViaTelegram(employeeId, message) {
    var emp = findRowByColumn(SHEET_NAMES.EMPLOYEES, 'ID', employeeId);
    if (!emp) return 'Lỗi: Không tìm thấy nhân viên ID ' + employeeId;
    
    var chatId = emp['Telegram Chat ID'];
    if (!chatId) {
        console.log('Employee ' + employeeId + ' has no Telegram Chat ID');
        return 'Lỗi: Nhân viên không có Chat ID';
    }
    
    return sendTelegramMessage(chatId, message);
}

function testTelegramConnection(botToken, chatId) {
    if (!botToken) return { success: false, message: 'Chưa nhập Bot Token' };
    if (!chatId) return { success: false, message: 'Chưa nhập Chat ID để test' };
    
    try {
        var url = 'https://api.telegram.org/bot' + botToken + '/sendMessage';
        var payload = {
            'chat_id': String(chatId),
            'text': '✅ Kết nối Telegram thành công!\n\n🤖 Bot QLCV đã sẵn sàng gửi thông báo.\n📅 ' + new Date().toLocaleString('vi-VN'),
            'parse_mode': 'HTML'
        };
        
        var options = {
            'method': 'post',
            'contentType': 'application/json',
            'payload': JSON.stringify(payload),
            'muteHttpExceptions': true
        };
        
        var response = UrlFetchApp.fetch(url, options);
        var result = JSON.parse(response.getContentText());
        
        if (result.ok) {
            return { success: true, message: 'Gửi tin nhắn test thành công! Kiểm tra Telegram.' };
        } else {
            return { success: false, message: 'Lỗi Telegram: ' + (result.description || 'Unknown error') };
        }
    } catch (e) {
        return { success: false, message: 'Lỗi kết nối: ' + e.message };
    }
}

// Nhắc nhở deadline - chạy tự động mỗi ngày (cần tạo trigger)
function dailyDeadlineReminder() {
    var tasks = getSheetData(SHEET_NAMES.TASKS);
    var now = new Date();
    var tomorrow = new Date();
    tomorrow.setDate(tomorrow.getDate() + 1);
    
    var count = 0;
    tasks.forEach(function(t) {
        if (t['Trạng thái'] === 'done') return;
        var due = new Date(t['Hạn hoàn thành']);
        
        // Nếu deadline là ngày mai
        if (due.toDateString() === tomorrow.toDateString()) {
            var ids = (t['ID Người thực hiện'] || '').toString().split(',');
            ids.forEach(function(id) {
                id = id.trim();
                if (id) {
                    notifyEmployeeViaTelegram(id, 
                        '⏰ <b>Nhắc nhở deadline</b>\n\n' +
                        'Nhiệm vụ "' + t['Tiêu đề'] + '" sẽ hết hạn <b>ngày mai</b>!\n' +
                        '📅 Hạn: ' + due.toLocaleDateString('vi-VN')
                    );
                    count++;
                }
            });
        }
        
        // Nếu đã quá hạn
        if (due < now && due.toDateString() !== now.toDateString()) {
            var ids2 = (t['ID Người thực hiện'] || '').toString().split(',');
            ids2.forEach(function(id) {
                id = id.trim();
                if (id) {
                    notifyEmployeeViaTelegram(id, 
                        '🚨 <b>Quá hạn!</b>\n\n' +
                        'Nhiệm vụ "' + t['Tiêu đề'] + '" đã quá hạn!\n' +
                        '📅 Hạn: ' + due.toLocaleDateString('vi-VN') + '\n' +
                        'Vui lòng hoàn thành ngay.'
                    );
                    count++;
                }
            });
        }
    });
    
    return { success: true, message: 'Đã gửi ' + count + ' thông báo nhắc nhở.' };
}

// ============ ZALO NOTIFICATION ============

function sendZaloMessage(chatId, message) {
    var settings = getSheetData(SHEET_NAMES.SETTINGS)[0] || {};
    var botToken = (settings['Zalo Bot Token'] || '').replace(/\s+/g, '');
    if (!botToken) return 'Lỗi: Chưa có Zalo Bot Token';
    if (!chatId) return 'Lỗi: Chưa có Zalo Chat ID';
    
    var url = 'https://bot-api.zaloplatforms.com/bot' + botToken + '/sendMessage';
    var payload = {
        chat_id: chatId,
        text: message
    };
    
    var options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
    };
    
    try {
        var response = UrlFetchApp.fetch(url, options);
        var responseText = response.getContentText();
        if (responseText.trim().indexOf('<') === 0) {
            return 'Lỗi Zalo (Sai Token hoặc server lỗi): ' + responseText.substring(0, 50);
        }
        var result = JSON.parse(responseText);
        if (!result.ok) {
            return 'Zalo Error: ' + (result.description || JSON.stringify(result));
        }
        return 'Thành công';
    } catch (e) {
        return 'Lỗi ngoại lệ Zalo: ' + e.message;
    }
}

function notifyEmployeeViaZalo(employeeId, message) {
    var emp = findRowByColumn(SHEET_NAMES.EMPLOYEES, 'ID', employeeId);
    if (!emp) return 'Lỗi: Không tìm thấy nhân viên ID ' + employeeId;
    
    var zaloId = emp['Zalo User ID'];
    if (!zaloId) return 'Lỗi: Nhân viên không có Zalo ID';
    
    return sendZaloMessage(zaloId, message);
}

function testZaloConnection(botToken, testUserId) {
    botToken = (botToken || '').replace(/\s+/g, '');
    if (!botToken) return { success: false, message: 'Vui lòng nhập Zalo Bot Token' };
    if (!testUserId) return { success: false, message: 'Vui lòng nhập Zalo User ID để gửi test' };
    
    try {
        var url = 'https://bot-api.zaloplatforms.com/bot' + botToken + '/sendMessage';
        var payload = {
            chat_id: testUserId,
            text: '✅ Kết nối Zalo thành công!\n\n🤖 Bot QLCV đã sẵn sàng.\n📅 ' + new Date().toLocaleString('vi-VN')
        };
        var options = {
            method: 'post',
            contentType: 'application/json',
            payload: JSON.stringify(payload),
            muteHttpExceptions: true
        };
        var response = UrlFetchApp.fetch(url, options);
        var responseText = response.getContentText();
        if (responseText.trim().indexOf('<') === 0) {
            return { success: false, message: 'Lỗi Zalo API (Sai Token hoặc server lỗi): Đảm bảo Token chính xác và không có khoảng trắng thừa.' };
        }
        var result = JSON.parse(responseText);
        
        if (result.ok) {
            // Save token if successful
            ensureColumnExists(SHEET_NAMES.SETTINGS, 'Zalo Bot Token');
            updateRow(SHEET_NAMES.SETTINGS, 2, { 'Zalo Bot Token': botToken });
            return { success: true, message: 'Thành công! Đã gửi tin nhắn test qua Zalo.' };
        } else {
            return { success: false, message: 'Lỗi Zalo: ' + (result.description || JSON.stringify(result)) };
        }
    } catch (e) {
        return { success: false, message: 'Lỗi ngoại lệ: ' + e.message };
    }
}

function doPost(e) {
    if (!e || !e.postData) return ContentService.createTextOutput("OK");
    try {
        var data = JSON.parse(e.postData.contents);
        
        // Debug: log raw webhook data để kiểm tra cấu trúc
        try {
            ensureColumnExists(SHEET_NAMES.SETTINGS, 'Webhook Debug');
            updateRow(SHEET_NAMES.SETTINGS, 2, { 'Webhook Debug': JSON.stringify(data).substring(0, 500) });
        } catch(dbg) {}
        
        // Phát hiện Zalo Bot Platform: kiểm tra event_name ở mọi cấp
        var isZalo = false;
        var zaloChatId = null;
        var zaloText = null;
        
        // Format 1: { ok: true, result: { message: {...}, event_name: "..." } }
        if (data.result && data.result.event_name && data.result.message) {
            isZalo = true;
            zaloChatId = data.result.message.chat ? data.result.message.chat.id : (data.result.message.from ? data.result.message.from.id : null);
            zaloText = data.result.message.text;
        }
        // Format 2: { message: {...}, event_name: "..." } (không có ok/result bọc ngoài)
        else if (data.event_name && data.message) {
            isZalo = true;
            zaloChatId = data.message.chat ? data.message.chat.id : (data.message.from ? data.message.from.id : null);
            zaloText = data.message.text;
        }
        
        if (isZalo && zaloChatId && zaloText) {
            zaloText = zaloText.trim();
            if (zaloText.toLowerCase().indexOf('/start') === 0) {
                var employeeId = zaloText.replace('/start', '').trim();
                var emps = getSheetData(SHEET_NAMES.EMPLOYEES);
                var emp = emps.find(function(em) { return String(em['ID']).toLowerCase() === employeeId.toLowerCase(); });
                
                if (emp) {
                    ensureColumnExists(SHEET_NAMES.EMPLOYEES, 'Zalo User ID');
                    updateRow(SHEET_NAMES.EMPLOYEES, emp._rowIndex, { 'Zalo User ID': zaloChatId });
                    sendZaloMessage(zaloChatId, '✅ Chào ' + emp['Họ tên'] + ', hệ thống đã liên kết thành công tài khoản Zalo của bạn với hệ thống Quản Lý Công Việc!');
                } else {
                    sendZaloMessage(zaloChatId, '❌ Không tìm thấy nhân viên với mã: ' + employeeId + '. Vui lòng kiểm tra lại!');
                }
            }
        }
        // Telegram Webhook format (có update_id, không có event_name)
        else if (!isZalo && data.message && data.message.chat && data.message.text) {
            var text = data.message.text.trim();
            var chatId = data.message.chat.id;
            
            if (text.toLowerCase().indexOf('/start') === 0) {
                var employeeId = text.replace('/start', '').trim();
                var emps = getSheetData(SHEET_NAMES.EMPLOYEES);
                var emp = emps.find(function(em) { return String(em['ID']).toLowerCase() === employeeId.toLowerCase(); });
                
                if (emp) {
                    ensureColumnExists(SHEET_NAMES.EMPLOYEES, 'Telegram Chat ID');
                    updateRow(SHEET_NAMES.EMPLOYEES, emp._rowIndex, { 'Telegram Chat ID': chatId });
                    sendTelegramMessage(chatId, '✅ Chào ' + emp['Họ tên'] + ', hệ thống đã liên kết thành công tài khoản Telegram của bạn với hệ thống Quản Lý Công Việc!');
                } else {
                    sendTelegramMessage(chatId, '❌ Không tìm thấy nhân viên với mã: ' + employeeId + '. Vui lòng kiểm tra lại!');
                }
            }
        }
    } catch (err) {
        console.error('Webhook error:', err);
    }
    return ContentService.createTextOutput("OK");
}

function setupZaloWebhook(botToken) {
    botToken = (botToken || '').replace(/\s+/g, '');
    if (!botToken) return { success: false, message: 'Vui lòng nhập Zalo Bot Token trước' };
    
    try {
        // Lấy URL Web App hiện tại của Google Apps Script
        var webAppUrl = ScriptApp.getService().getUrl();
        if (!webAppUrl) {
            return { success: false, message: 'Chưa triển khai Web App. Vui lòng vào Apps Script > Triển khai > Triển khai mới > Ứng dụng Web, rồi thử lại.' };
        }
        
        var url = 'https://bot-api.zaloplatforms.com/bot' + botToken + '/setWebhook';
        var secretToken = Utilities.getUuid().replace(/-/g, '');
        var options = {
            method: 'post',
            contentType: 'application/json',
            payload: JSON.stringify({ url: webAppUrl, secret_token: secretToken }),
            muteHttpExceptions: true
        };
        var response = UrlFetchApp.fetch(url, options);
        var httpCode = response.getResponseCode();
        var responseText = response.getContentText();
        
        if (httpCode === 401 || httpCode === 403) {
            return { success: false, message: '❌ Token không hợp lệ. Hãy kiểm tra lại Zalo Bot Token.' };
        }
        if (httpCode !== 200) {
            return { success: false, message: 'Lỗi HTTP ' + httpCode + ': ' + responseText.substring(0, 100) };
        }
        
        var result = JSON.parse(responseText);
        if (result.ok) {
            return { success: true, message: '✅ Đã kết nối Webhook thành công!\nURL: ' + webAppUrl };
        } else {
            return { success: false, message: 'Lỗi Zalo: ' + (result.description || JSON.stringify(result)) };
        }
    } catch (e) {
        return { success: false, message: 'Lỗi ngoại lệ: ' + e.message };
    }
}

function getLatestZaloChatId(botToken) {
    botToken = (botToken || '').replace(/\s+/g, '');
    if (!botToken) return { success: false, message: 'Vui lòng nhập Zalo Bot Token trước' };
    try {
        var url = 'https://bot-api.zaloplatforms.com/bot' + botToken + '/getUpdates';
        var options = {
            method: 'post',
            contentType: 'application/json',
            payload: JSON.stringify({ timeout: 0 }),
            muteHttpExceptions: true
        };
        var response = UrlFetchApp.fetch(url, options);
        var httpCode = response.getResponseCode();
        var responseText = response.getContentText();
        
        // Xử lý theo HTTP status code trước
        if (httpCode === 504 || httpCode === 502) {
            return { success: false, message: '⏳ Token hợp lệ nhưng chưa có tin nhắn mới (server timeout). Hãy mở Zalo gửi "hello" cho Bot, rồi bấm lại nút này NGAY trong vòng 30 giây.' };
        }
        if (httpCode === 401 || httpCode === 403) {
            return { success: false, message: '❌ Token không hợp lệ (Unauthorized). Hãy kiểm tra lại Zalo Bot Token.' };
        }
        if (httpCode === 404) {
            return { success: false, message: '❌ Không tìm thấy Bot. Token có thể sai hoặc Bot đã bị xóa.' };
        }
        if (httpCode !== 200) {
            return { success: false, message: 'Lỗi HTTP ' + httpCode + ': ' + responseText.substring(0, 100) };
        }
        
        // HTTP 200 - parse JSON
        var result = JSON.parse(responseText);
        
        if (!result.ok) {
            if (result.description && result.description.toLowerCase().indexOf('timeout') !== -1) {
                return { success: false, message: '⏳ Token hợp lệ nhưng chưa có tin nhắn mới. Hãy mở Zalo gửi "hello" cho Bot, rồi bấm lại nút này NGAY.' };
            }
            if (result.description && result.description.indexOf('webhook') !== -1) {
                return { success: false, message: '⚠️ Đang có Webhook hoạt động. Vui lòng vào bot.zaloplatforms.com xóa Webhook trước.' };
            }
            return { success: false, message: 'Lỗi Zalo: ' + (result.description || JSON.stringify(result)) };
        }
        
        var data = result.result;
        if (!data) {
            return { success: false, message: '⏳ Token hợp lệ nhưng chưa có tin nhắn mới. Hãy mở Zalo gửi "hello" cho Bot, rồi bấm lại nút này NGAY.' };
        }
        
        // Zalo Bot Platform: result có thể là object đơn hoặc mảng
        var chatId = null;
        if (data.message && data.message.chat && data.message.chat.id) {
            chatId = data.message.chat.id;
        } else if (data.message && data.message.from && data.message.from.id) {
            chatId = data.message.from.id;
        } else if (Array.isArray(data) && data.length > 0) {
            var last = data[data.length - 1];
            if (last && last.message && last.message.chat) chatId = last.message.chat.id;
            else if (last && last.message && last.message.from) chatId = last.message.from.id;
            else if (last && last.result && last.result.message && last.result.message.chat) chatId = last.result.message.chat.id;
        }
        
        if (chatId) {
            return { success: true, chatId: chatId };
        } else {
            return { success: false, message: 'Không thể trích xuất ID. Dữ liệu: ' + JSON.stringify(result).substring(0, 300) };
        }
    } catch (e) {
        if (e.message && e.message.indexOf('Timeout') !== -1) {
            return { success: false, message: '⏳ Token có vẻ hợp lệ nhưng server timeout. Hãy gửi "hello" cho Bot trên Zalo, rồi bấm lại nút này NGAY.' };
        }
        return { success: false, message: 'Lỗi ngoại lệ: ' + e.message };
    }
}

// ============ LAZY LOAD APIS ============
function getCoreData(userRole, userId, userDeptId) {
    var ss = getSS();
    var deptsRaw = getSheetData(SHEET_NAMES.DEPARTMENTS);
    var empsRaw = getSheetData(SHEET_NAMES.EMPLOYEES);
    var catsRaw = getSheetData(SHEET_NAMES.CATEGORIES);
    var tasksRaw = getSheetData(SHEET_NAMES.TASKS);
    var projectsRaw = getSheetData(SHEET_NAMES.PROJECTS);
    var notifsRaw = getSheetData(SHEET_NAMES.NOTIFICATIONS);
    var settingsRaw = getSheetData(SHEET_NAMES.SETTINGS);
    
    var departments = deptsRaw.map(function(d) {
        return { id: d['ID'], name: d['Tên phòng ban'], description: d['Mô tả'], created_at: d['Ngày tạo'] };
    });
    var categories = catsRaw.map(function(c) {
        return { id: c['ID'], name: c['Tên danh mục'], icon: c['Icon'], color: c['Màu sắc'], is_default: c['Mặc định'] === true || c['Mặc định'] === 'TRUE', created_at: c['Ngày tạo'] };
    });
    var filteredEmps = empsRaw;
    var employees = filteredEmps.map(function(e) {
        var deptIds = parseIds(e['ID Phòng ban']);
        return { id: e['ID'], name: e['Họ tên'], username: e['Tên đăng nhập'], role: e['Vai trò'], department_id: deptIds[0] || '', department_ids: deptIds, avatar: e['Ảnh đại diện'] };
    });
    var projects = projectsRaw.map(function(p) {
        return { id: p['ID'], name: p['Tên dự án'], color: p['Màu sắc'] || '#6366f1' };
    });
    
    var filteredTasks = tasksRaw;
    if (userRole === 'Member') {
        filteredTasks = tasksRaw.filter(function(t) {
            var rawVal = (t['ID Người thực hiện'] || '').toString();
            var ids = rawVal.split(',').map(function(id) { return id.trim(); });
            return ids.indexOf(String(userId).trim()) > -1;
        });
    } else if (userRole === 'Manager') {
        var managerDeptIds2 = parseIds(userDeptId);
        var managerUid2 = String(userId).trim();
        filteredTasks = tasksRaw.filter(function(t) {
            var taskDeptId = String(t['ID Phòng ban'] || '').trim();
            var creatorId = String(t['ID Người tạo'] || '').trim();
            var aIds = (t['ID Người thực hiện'] || '').toString().split(',').map(function(x) { return x.trim(); });
            return managerDeptIds2.indexOf(taskDeptId) > -1 || creatorId === managerUid2 || aIds.indexOf(managerUid2) > -1;
        });
    }
    
    var tasks = filteredTasks.map(function(t) {
        var assigneeIdStr = (t['ID Người thực hiện'] || '').toString();
        var assigneeIds = assigneeIdStr ? assigneeIdStr.split(',').map(function(id) { return id.trim(); }) : [];
        var taskAssignees = empsRaw.filter(function(e) { return assigneeIds.indexOf(String(e['ID']).trim()) > -1; })
                                  .map(function(e) { return { id: String(e['ID']).trim(), name: e['Họ tên'], avatar: e['Ảnh đại diện'] }; });
        var projId = String(t['ID Dự án'] || '').trim();
        var projectObj = projects.find(function(p) { return String(p.id).trim() === projId; });
        var catId = String(t['ID Danh mục'] || '').trim();
        var deptId = String(t['ID Phòng ban'] || '').trim();
        
        return {
            id: t['ID'], title: t['Tiêu đề'], description: t['Mô tả'],
            assignee_id: assigneeIdStr, assignee_ids: assigneeIds,
            department_id: deptId, category_id: catId, project_id: projId,
            priority: t['Độ ưu tiên'], status: t['Trạng thái'],
            start_date: t['Ngày bắt đầu'] || t['Ngày tạo'],
            due_date: t['Hạn hoàn thành'], completed_at: t['Ngày hoàn thành'],
            created_by: t['ID Người tạo'], created_at: t['Ngày tạo'],
            assignees: taskAssignees,
            assignee: taskAssignees.length > 0 ? taskAssignees[0] : null,
            category: categories.find(function(c) { return String(c.id).trim() === catId; }),
            department: departments.find(function(d) { return String(d.id).trim() === deptId; }),
            project: projectObj || null,
            checklist: (function() { try { return JSON.parse(t['Checklist'] || '[]'); } catch(e) { return []; } })(),
            tags: (t['Tags'] || '').toString().split(',').filter(function(x) { return x.trim(); }),
            starred: (t['Starred'] || '').toString().split(',').filter(function(x) { return x.trim(); }),
            recurrence: t['Lặp lại'] || 'none'
        };
    });
    
    tasks.sort(function(a, b) {
        return (b.created_at ? new Date(b.created_at) : new Date(0)) - (a.created_at ? new Date(a.created_at) : new Date(0));
    });
    
    var userNotifs = notifsRaw.filter(function(n) { return n['ID Người nhận'] === userId || n['ID Người nhận'] === 'all'; })
                             .sort(function(a, b) { return new Date(b['Ngày tạo']) - new Date(a['Ngày tạo']); })
                             .slice(0, 10)
                             .map(function(n) {
                                 return { id: n['ID'], user_id: n['ID Người nhận'], title: n['Tiêu đề'], description: n['Nội dung'], link: n['Đường dẫn'], read: n['Đã đọc'] === true || n['Đã đọc'] === 'TRUE', created_at: n['Ngày tạo'] };
                             });
                             
    var settingsObj = settingsRaw.length > 0 ? {
        companyName: settingsRaw[0]['Tên công ty'],
        description: settingsRaw[0]['Mô tả'],
        telegramBotToken: settingsRaw[0]['Telegram Bot Token'] || '',
        zaloBotToken: settingsRaw[0]['Zalo Bot Token'] || '',
        geminiApiKey: settingsRaw[0]['Gemini API Key'] || ''
    } : {};
    
    var stats = computeDashboardStats(tasks, empsRaw, departments, categories, userRole);
    
    return JSON.stringify({
        tasks: tasks,
        settings: settingsObj,
        stats: stats,
        notifications: userNotifs
    });
}

function getExtendedData(userRole, userId, userDeptId) {
    var deptsRaw = getSheetData(SHEET_NAMES.DEPARTMENTS);
    var empsRaw = getSheetData(SHEET_NAMES.EMPLOYEES);
    var catsRaw = getSheetData(SHEET_NAMES.CATEGORIES);
    var projectsRaw = getSheetData(SHEET_NAMES.PROJECTS);
    
    var departments = deptsRaw.map(function(d) {
        return { id: d['ID'], name: d['Tên phòng ban'], description: d['Mô tả'], created_at: d['Ngày tạo'] };
    });
    
    var categories = catsRaw.map(function(c) {
        return {
            id: c['ID'], name: c['Tên danh mục'], icon: c['Icon'],
            color: c['Màu sắc'], is_default: c['Mặc định'] === true || c['Mặc định'] === 'TRUE',
            created_at: c['Ngày tạo']
        };
    });
    
    var filteredEmps = empsRaw;
    var employees = filteredEmps.map(function(e) {
        var deptIds = parseIds(e['ID Phòng ban']);
        var deptObjs = departments.filter(function(d) { return deptIds.indexOf(d.id) > -1; });
        return {
            id: e['ID'], name: e['Họ tên'], username: e['Tên đăng nhập'],
            email: e['Email'], phone: formatPhone(e['Điện thoại']), avatar: e['Ảnh đại diện'],
            role: e['Vai trò'], department_id: deptIds[0] || '',
            department_ids: deptIds,
            telegram_chat_id: e['Telegram Chat ID'] || '',
            zalo_user_id: e['Zalo User ID'] || '',
            department: deptObjs[0] || null,
            departments: deptObjs,
            created_at: e['Ngày tạo']
        };
    });
    
    var projects = projectsRaw.map(function(p) {
        var managerId = String(p['ID Người phụ trách'] || '').trim();
        var manager = employees.find(function(e) { return String(e.id).trim() === managerId; });
        var memberIdStr = (p['ID Thành viên'] || '').toString();
        var memberIds = memberIdStr ? memberIdStr.split(',').map(function(id) { return id.trim(); }) : [];
        var projectMembers = empsRaw.filter(function(e) { 
            return memberIds.indexOf(String(e['ID']).trim()) > -1; 
        }).map(function(e) { 
            return { id: String(e['ID']).trim(), name: e['Họ tên'], avatar: e['Ảnh đại diện'] }; 
        });

        return {
            id: p['ID'], name: p['Tên dự án'], description: p['Mô tả'],
            target: p['Mục tiêu'], status: p['Trạng thái'] || 'active',
            color: p['Màu sắc'] || '#6366f1',
            start_date: p['Ngày bắt đầu'], end_date: p['Ngày kết thúc'],
            target_completed: p['Đạt mục tiêu'] === true || p['Đạt mục tiêu'] === 'TRUE',
            manager_id: managerId,
            manager: manager ? { id: manager.id, name: manager.name, avatar: manager.avatar } : null,
            member_ids: memberIds,
            members: projectMembers,
            created_by: p['ID Người tạo'], created_at: p['Ngày tạo']
        };
    });
    
    return JSON.stringify({
        departments: departments,
        employees: employees,
        categories: categories,
        projects: projects
    });
}

// ============ RECURRING TASKS CRON ENGINE ============
function createRecurringTasks() {
    ensureColumnExists(SHEET_NAMES.TASKS, 'Lặp lại');
    ensureColumnExists(SHEET_NAMES.TASKS, 'Lần chạy cuối');
    var tasks = getSheetData(SHEET_NAMES.TASKS);
    var now = new Date();
    var todayStr = Utilities.formatDate(now, Session.getScriptTimeZone() || "GMT+7", "yyyy-MM-dd");
    
    tasks.forEach(function(task) {
        var recurrence = task['Lặp lại'];
        if (!recurrence || recurrence === '' || recurrence === 'none') return;
        
        var lastRun = task['Lần chạy cuối'] || '';
        if (lastRun === todayStr) return; // Already run today
        
        var shouldSpawn = false;
        var createdDate = task['Ngày tạo'] ? new Date(task['Ngày tạo']) : new Date();
        
        if (recurrence === 'daily') {
            shouldSpawn = true;
        } else if (recurrence === 'weekly') {
            if (now.getDay() === createdDate.getDay()) {
                shouldSpawn = true;
            }
        } else if (recurrence === 'monthly') {
            if (now.getDate() === createdDate.getDate()) {
                shouldSpawn = true;
            }
        }
        
        if (shouldSpawn) {
            var newDueDateStr = '';
            if (task['Hạn hoàn thành'] && task['Ngày tạo']) {
                var origDue = new Date(task['Hạn hoàn thành']);
                var origCreated = new Date(task['Ngày tạo']);
                var diffTime = Math.abs(origDue - origCreated);
                var diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
                
                var newDue = new Date(now.getTime() + diffDays * 24 * 60 * 60 * 1000);
                newDueDateStr = Utilities.formatDate(newDue, Session.getScriptTimeZone() || "GMT+7", "yyyy-MM-dd");
            } else {
                newDueDateStr = todayStr;
            }
            
            var newChecklist = [];
            if (task['Checklist']) {
                try {
                    var parsed = JSON.parse(task['Checklist']);
                    newChecklist = parsed.map(function(item) {
                        return { text: item.text, done: false };
                    });
                } catch(e) {}
            }
            
            createTask({
                title: task['Tiêu đề'],
                description: task['Mô tả'] || '',
                assignee_ids: (task['ID Người thực hiện'] || '').toString().split(',').filter(Boolean),
                department_id: task['ID Phòng ban'] || '',
                category_id: task['ID Danh mục'] || '',
                project_id: task['ID Dự án'] || '',
                priority: task['Độ ưu tiên'] || 'medium',
                due_date: newDueDateStr,
                checklist: newChecklist,
                tags: task['Tags'] || ''
            }, task['ID Người tạo'] || 'system');
            
            updateRow(SHEET_NAMES.TASKS, task._rowIndex, { 'Lần chạy cuối': todayStr });
        }
    });
}

function setupRecurringTasksTrigger() {
    var triggers = ScriptApp.getProjectTriggers();
    var hasTrigger = false;
    for (var i = 0; i < triggers.length; i++) {
        if (triggers[i].getHandlerFunction() === 'createRecurringTasks') {
            hasTrigger = true;
            break;
        }
    }
    if (!hasTrigger) {
        ScriptApp.newTrigger('createRecurringTasks')
            .timeBased()
            .everyDays(1)
            .atHour(2)
            .create();
        return "Trigger created successfully!";
    }
    return "Trigger already exists.";
}

// ============ GEMINI AI INTEGRATION ============
function testGeminiConnection(apiKey) {
    if (!apiKey) {
        return { success: false, message: "Khóa API không được để trống." };
    }
    const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + apiKey;
    const payload = {
        contents: [{ parts: [{ text: "Hello, reply in 2 words." }] }]
    };
    const options = {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
    };
    try {
        const response = UrlFetchApp.fetch(url, options);
        const code = response.getResponseCode();
        const text = response.getContentText();
        if (code === 200) {
            return { success: true, message: "Kết nối thành công tới Gemini API!" };
        } else {
            return { success: false, message: "Lỗi kết nối (" + code + "): " + text };
        }
    } catch (e) {
        return { success: false, message: "Lỗi kết nối hệ thống: " + e.toString() };
    }
}

function parseTaskWithGemini(prompt, employees, categories, projects) {
    const settingsRaw = getSheetData(SHEET_NAMES.SETTINGS);
    const apiKey = settingsRaw.length > 0 ? settingsRaw[0]['Gemini API Key'] : null;
    if (!apiKey) {
        throw new Error("Vui lòng cấu hình Gemini API Key trong phần Cài đặt trước.");
    }

    const context = {
        employees: employees.map(e => ({ id: e.id, name: e.name })),
        categories: categories.map(c => ({ id: c.id, name: c.name })),
        projects: projects.map(p => ({ id: p.id, name: p.name }))
    };

    const systemInstruction = "Bạn là trợ lý AI thông minh tích hợp trong hệ thống quản lý công việc.\n" +
        "Nhiệm vụ của bạn là phân tích yêu cầu tạo nhiệm vụ bằng ngôn ngữ tự nhiên của người dùng và trích xuất thành thông tin chi tiết dưới dạng JSON.\n" +
        "Bạn phải tìm và khớp người thực hiện, danh mục, và dự án từ danh sách được cung cấp dưới đây.\n" +
        "Nếu khớp, hãy trả về ID của đối tượng đó. Nếu không khớp hoặc không đề cập, hãy để trống hoặc mảng rỗng.\n" +
        "Danh sách dữ liệu hiện tại:\n" +
        JSON.stringify(context) + "\n\n" +
        "Các trường trong JSON kết quả:\n" +
        "- title: Tiêu đề nhiệm vụ ngắn gọn (bắt buộc, ví dụ: 'Báo cáo nhiệm vụ tháng 5').\n" +
        "- description: Mô tả chi tiết nhiệm vụ (nếu có).\n" +
        "- assignee_ids: Mảng chứa các ID người thực hiện phù hợp (khớp từ tên người dùng, ví dụ: ['nv_123']). Nếu không tìm thấy hoặc không có người phù hợp, hãy trả về mảng rỗng.\n" +
        "- category_id: ID của danh mục phù hợp (khớp từ tên danh mục). Nếu không khớp, để trống.\n" +
        "- project_id: ID của dự án phù hợp (khớp từ tên dự án). Nếu không khớp, để trống.\n" +
        "- priority: Độ ưu tiên, chỉ nhận giá trị 'low', 'medium', hoặc 'high' (mặc định là 'medium' nếu không nói rõ. 'quan trọng', 'khẩn cấp' -> 'high').\n" +
        "- due_date: Hạn hoàn thành định dạng 'YYYY-MM-DD'. Nếu có thời gian cụ thể (ví dụ: 'tháng 5', 'ngày 15/6', 'ngày mai'), hãy tính toán ngày tương ứng dựa trên thời gian hiện tại là: " + getLocalISOString() + ".\n" +
        "- recurrence: Lặp lại, chỉ nhận 'none', 'daily', 'weekly', 'monthly'. Mặc định 'none'.\n\n" +
        "LƯU Ý: Chỉ trả về chuỗi JSON thô, không bọc trong markdown (```json ... ```), không thêm bất kỳ văn bản giải thích nào khác.";

    const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + apiKey;
    const payload = {
        contents: [
            {
                parts: [
                    { text: "Hãy phân tích yêu cầu sau: \"" + prompt + "\"" }
                ]
            }
        ],
        systemInstruction: {
            parts: [
                { text: systemInstruction }
            ]
        },
        generationConfig: {
            responseMimeType: "application/json"
        }
    };

    const options = {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    const responseText = response.getContentText();
    
    if (code !== 200) {
        throw new Error("Lỗi gọi Gemini API (Mã lỗi " + code + "): " + responseText);
    }

    try {
        const resJson = JSON.parse(responseText);
        const textResult = resJson.candidates[0].content.parts[0].text;
        return JSON.parse(textResult.trim());
    } catch (e) {
        throw new Error("Lỗi phân tích cú pháp kết quả AI: " + e.toString() + "\nRaw response: " + responseText);
    }
}

// ============ MIGRATION: chuẩn hóa tên vai trò 'Leader' -> 'Manager' ============
// Chạy 1 lần thủ công trên spreadsheet đang dùng thật (OS) sau khi deploy bản code này.
// Lý do cần: code cũ lưu/kiểm tra vai trò quản lý phòng ban bằng chuỗi 'Leader',
// trong khi form tạo nhân viên trên UI luôn ghi 'Manager' -> lệch chuỗi gây sai quyền.
// Đổi code không tự sửa dữ liệu chữ đã có sẵn trong sheet, nên cần chạy hàm này 1 lần.
function migrateLeaderRoleToManager() {
    var sheet = getSheet(SHEET_NAMES.EMPLOYEES);
    if (!sheet) return { success: false, message: 'Không tìm thấy sheet Nhân Viên' };

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h) { return String(h).trim(); });
    var roleCol = headers.indexOf('Vai trò') + 1;
    if (roleCol === 0) return { success: false, message: 'Không tìm thấy cột Vai trò' };

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, updated: 0 };

    var range = sheet.getRange(2, roleCol, lastRow - 1, 1);
    var values = range.getValues();
    var updated = 0;
    for (var i = 0; i < values.length; i++) {
        if (String(values[i][0]).trim() === 'Leader') {
            values[i][0] = 'Manager';
            updated++;
        }
    }
    if (updated > 0) {
        range.setValues(values);
        clearCache('data_v2_' + SHEET_NAMES.EMPLOYEES);
    }
    return { success: true, updated: updated, message: 'Đã đổi ' + updated + ' tài khoản từ Leader sang Manager' };
}

// ============ THÔNG BÁO NỘI BỘ (ANNOUNCEMENTS) ============
// Trạng thái: draft | scheduled | published
// Mức độ: normal | important | urgent (urgent = bắt buộc xác nhận đã đọc)
// Phạm vi: all | departments | employees
var ANN_HEADERS = [
    'ID', 'Tiêu đề', 'Nội dung', 'Mức độ', 'Trạng thái', 'Phạm vi',
    'ID Phòng ban nhận', 'ID Người nhận', 'Nhắc tên', 'Thẻ', 'Đính kèm',
    'ID Người đăng', 'ID Phòng ban phát hành', 'Thời gian đăng',
    'Đã đọc bởi', 'Đã xác nhận bởi', 'Ngày tạo', 'Ngày cập nhật'
];
var ANN_FILTER_HEADERS = ['ID', 'ID Nhân viên', 'Tên bộ lọc', 'Cấu hình', 'Ngày tạo'];

function ensureAnnouncementSheets() {
    var ss = getSS();
    if (!ss.getSheetByName(SHEET_NAMES.ANNOUNCEMENTS)) {
        _createSheet(ss, SHEET_NAMES.ANNOUNCEMENTS, ANN_HEADERS);
    }
    if (!ss.getSheetByName(SHEET_NAMES.SAVED_FILTERS)) {
        _createSheet(ss, SHEET_NAMES.SAVED_FILTERS, ANN_FILTER_HEADERS);
    }
    return { success: true };
}

// Lấy danh sách phòng ban của 1 nhân viên (nguồn chuẩn từ sheet Nhân Viên)
function _userDeptIds(userId) {
    var emp = findRowByColumn(SHEET_NAMES.EMPLOYEES, 'ID', userId);
    return emp ? parseIds(emp['ID Phòng ban']) : [];
}

// Sheets tự ép chuỗi datetime thành Date. Khi trả về client, Date bị serialize sang ISO UTC
// làm lệch giờ lúc mở lại để sửa -> luôn chuẩn hóa về ISO kèm offset múi giờ của script.
function _annDateStr(v) {
    if (!v) return '';
    if (v instanceof Date) return getLocalISOString(v);
    return String(v);
}

// Chiều ghi: input datetime-local là "YYYY-MM-DDTHH:mm" (không có offset).
// Chuyển sang ISO đầy đủ kèm offset để không phụ thuộc cách Sheets/JS đoán múi giờ.
function _annParsePublishAt(v) {
    if (!v) return '';
    var d = (v instanceof Date) ? v : new Date(v);
    if (isNaN(d.getTime())) return String(v);
    return getLocalISOString(d);
}

function _annIsPublished(a, now) {
    if (a['Trạng thái'] === 'published') return true;
    if (a['Trạng thái'] === 'scheduled' && a['Thời gian đăng']) {
        return new Date(a['Thời gian đăng']) <= now;
    }
    return false;
}

function _annVisibleTo(a, userId, deptIds) {
    var scope = a['Phạm vi'] || 'all';
    if (scope === 'all') return true;
    if (scope === 'departments') {
        var target = parseIds(a['ID Phòng ban nhận']);
        return deptIds.some(function(d) { return target.indexOf(d) > -1; });
    }
    if (scope === 'employees') {
        return parseIds(a['ID Người nhận']).indexOf(userId) > -1;
    }
    return false;
}

// Danh sách nhân viên thực sự nhận được thông báo
function _annRecipientIds(a) {
    var scope = a['Phạm vi'] || 'all';
    if (scope === 'employees') return parseIds(a['ID Người nhận']);
    var emps = getSheetData(SHEET_NAMES.EMPLOYEES);
    if (scope === 'all') return emps.map(function(e) { return e['ID']; });
    var target = parseIds(a['ID Phòng ban nhận']);
    return emps.filter(function(e) {
        return parseIds(e['ID Phòng ban']).some(function(d) { return target.indexOf(d) > -1; });
    }).map(function(e) { return e['ID']; });
}

function _mapAnnouncement(a, userId) {
    var readBy = parseIds(a['Đã đọc bởi']);
    var confirmedBy = parseIds(a['Đã xác nhận bởi']);
    var mentioned = parseIds(a['Nhắc tên']);
    var attachments = [];
    try { attachments = a['Đính kèm'] ? JSON.parse(a['Đính kèm']) : []; } catch (e) {}
    return {
        id: a['ID'],
        title: a['Tiêu đề'],
        content: a['Nội dung'],
        priority: a['Mức độ'] || 'normal',
        status: a['Trạng thái'] || 'draft',
        scope: a['Phạm vi'] || 'all',
        department_ids: parseIds(a['ID Phòng ban nhận']),
        employee_ids: parseIds(a['ID Người nhận']),
        mentioned_ids: mentioned,
        tags: parseIds(a['Thẻ']),
        attachments: attachments,
        created_by: a['ID Người đăng'],
        publisher_dept_id: a['ID Phòng ban phát hành'],
        published_at: _annDateStr(a['Thời gian đăng']),
        created_at: _annDateStr(a['Ngày tạo']),
        updated_at: _annDateStr(a['Ngày cập nhật']),
        read: readBy.indexOf(userId) > -1,
        confirmed: confirmedBy.indexOf(userId) > -1,
        mentioned_me: mentioned.indexOf(userId) > -1,
        read_count: readBy.length,
        confirmed_count: confirmedBy.length
    };
}

function getAnnouncements(userRole, userId) {
    ensureAnnouncementSheets();
    var now = new Date();
    var deptIds = _userDeptIds(userId);
    var isManager = userRole === 'Owner' || userRole === 'Manager';

    return getSheetData(SHEET_NAMES.ANNOUNCEMENTS)
        .filter(function(a) {
            if (!a['ID']) return false;
            // Owner quản trị toàn hệ thống: thấy mọi thông báo
            if (userRole === 'Owner') return true;
            // Người quản lý luôn thấy thông báo do chính mình đăng (kể cả nháp/hẹn giờ) để theo dõi
            if (isManager && a['ID Người đăng'] === userId) return true;
            return _annIsPublished(a, now) && _annVisibleTo(a, userId, deptIds);
        })
        .sort(function(x, y) {
            return new Date(y['Thời gian đăng'] || y['Ngày tạo']) - new Date(x['Thời gian đăng'] || x['Ngày tạo']);
        })
        .map(function(a) { return _mapAnnouncement(a, userId); });
}

function createAnnouncement(data, createdBy) {
    ensureAnnouncementSheets();
    var id = generateId('tbnb');
    var nowIso = getLocalISOString();
    var status = data.status || 'published';
    var publishAt = status === 'scheduled' ? (_annParsePublishAt(data.published_at) || nowIso) : nowIso;

    appendRow(SHEET_NAMES.ANNOUNCEMENTS, {
        'ID': id,
        'Tiêu đề': data.title || '',
        'Nội dung': data.content || '',
        'Mức độ': data.priority || 'normal',
        'Trạng thái': status,
        'Phạm vi': data.scope || 'all',
        'ID Phòng ban nhận': (data.department_ids || []).join(','),
        'ID Người nhận': (data.employee_ids || []).join(','),
        'Nhắc tên': (data.mentioned_ids || []).join(','),
        'Thẻ': (data.tags || []).join(','),
        'Đính kèm': data.attachments && data.attachments.length ? JSON.stringify(data.attachments) : '',
        'ID Người đăng': createdBy,
        'ID Phòng ban phát hành': (_userDeptIds(createdBy)[0] || ''),
        'Thời gian đăng': status === 'draft' ? '' : publishAt,
        'Đã đọc bởi': '',
        'Đã xác nhận bởi': '',
        'Ngày tạo': nowIso,
        'Ngày cập nhật': nowIso
    });

    if (status === 'published') _pushAnnouncementToChannels(id);
    return { success: true, id: id };
}

function updateAnnouncement(id, data) {
    var a = findRowByColumn(SHEET_NAMES.ANNOUNCEMENTS, 'ID', id);
    if (!a) return { success: false, message: 'Không tìm thấy thông báo' };

    var updates = { 'Ngày cập nhật': getLocalISOString() };
    if (data.title !== undefined) updates['Tiêu đề'] = data.title;
    if (data.content !== undefined) updates['Nội dung'] = data.content;
    if (data.priority !== undefined) updates['Mức độ'] = data.priority;
    if (data.scope !== undefined) updates['Phạm vi'] = data.scope;
    if (data.department_ids !== undefined) updates['ID Phòng ban nhận'] = data.department_ids.join(',');
    if (data.employee_ids !== undefined) updates['ID Người nhận'] = data.employee_ids.join(',');
    if (data.mentioned_ids !== undefined) updates['Nhắc tên'] = data.mentioned_ids.join(',');
    if (data.tags !== undefined) updates['Thẻ'] = data.tags.join(',');
    if (data.attachments !== undefined) updates['Đính kèm'] = data.attachments.length ? JSON.stringify(data.attachments) : '';
    if (data.published_at !== undefined) updates['Thời gian đăng'] = _annParsePublishAt(data.published_at);

    var wasPublished = a['Trạng thái'] === 'published';
    if (data.status !== undefined) {
        updates['Trạng thái'] = data.status;
        if (data.status === 'published') {
            // Sửa tin đã đăng thì giữ nguyên thời điểm đăng gốc;
            // chuyển từ nháp/hẹn giờ sang đăng ngay mới lấy thời điểm hiện tại.
            if (wasPublished) delete updates['Thời gian đăng'];
            else updates['Thời gian đăng'] = getLocalISOString();
        }
    }

    updateRow(SHEET_NAMES.ANNOUNCEMENTS, a._rowIndex, updates);

    if (!wasPublished && data.status === 'published') _pushAnnouncementToChannels(id);
    return { success: true };
}

function deleteAnnouncement(id) {
    var a = findRowByColumn(SHEET_NAMES.ANNOUNCEMENTS, 'ID', id);
    if (!a) return { success: false, message: 'Không tìm thấy thông báo' };
    deleteRow(SHEET_NAMES.ANNOUNCEMENTS, a._rowIndex);
    return { success: true };
}

// Ghi nhận đã đọc / đã xác nhận. Dùng LockService vì nhiều người có thể đọc cùng lúc.
function markAnnouncementRead(id, userId, confirmed) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
    } catch (e) {
        return { success: false, message: 'Hệ thống đang bận, vui lòng thử lại' };
    }
    try {
        clearCache('data_v2_' + SHEET_NAMES.ANNOUNCEMENTS);
        var a = findRowByColumn(SHEET_NAMES.ANNOUNCEMENTS, 'ID', id);
        if (!a) return { success: false, message: 'Không tìm thấy thông báo' };

        var updates = {};
        var readBy = parseIds(a['Đã đọc bởi']);
        if (readBy.indexOf(userId) === -1) {
            readBy.push(userId);
            updates['Đã đọc bởi'] = readBy.join(',');
        }
        if (confirmed) {
            var confirmedBy = parseIds(a['Đã xác nhận bởi']);
            if (confirmedBy.indexOf(userId) === -1) {
                confirmedBy.push(userId);
                updates['Đã xác nhận bởi'] = confirmedBy.join(',');
            }
        }
        if (Object.keys(updates).length > 0) {
            updateRow(SHEET_NAMES.ANNOUNCEMENTS, a._rowIndex, updates);
        }
        return { success: true };
    } finally {
        lock.releaseLock();
    }
}

// Bảng theo dõi Đã đọc / Chưa đọc của 1 thông báo
function getAnnouncementReadReport(id) {
    var a = findRowByColumn(SHEET_NAMES.ANNOUNCEMENTS, 'ID', id);
    if (!a) return { success: false, message: 'Không tìm thấy thông báo' };

    var readBy = parseIds(a['Đã đọc bởi']);
    var confirmedBy = parseIds(a['Đã xác nhận bởi']);
    var recipientIds = _annRecipientIds(a);
    var deptMap = {};
    getSheetData(SHEET_NAMES.DEPARTMENTS).forEach(function(d) { deptMap[d['ID']] = d['Tên phòng ban']; });

    var rows = getSheetData(SHEET_NAMES.EMPLOYEES)
        .filter(function(e) { return recipientIds.indexOf(e['ID']) > -1; })
        .map(function(e) {
            return {
                id: e['ID'],
                name: e['Họ tên'],
                avatar: e['Ảnh đại diện'] || '',
                department: parseIds(e['ID Phòng ban']).map(function(d) { return deptMap[d] || ''; })
                    .filter(function(x) { return x; }).join(', '),
                read: readBy.indexOf(e['ID']) > -1,
                confirmed: confirmedBy.indexOf(e['ID']) > -1
            };
        });

    return {
        success: true,
        title: a['Tiêu đề'],
        priority: a['Mức độ'] || 'normal',
        total: rows.length,
        read_count: rows.filter(function(r) { return r.read; }).length,
        confirmed_count: rows.filter(function(r) { return r.confirmed; }).length,
        rows: rows
    };
}

// 1-Click gửi nhắc nhở cho những người chưa đọc
function sendAnnouncementReminder(id) {
    var report = getAnnouncementReadReport(id);
    if (!report.success) return report;

    var unread = report.rows.filter(function(r) { return !r.read; });
    if (unread.length === 0) return { success: true, sent: 0, message: 'Tất cả mọi người đã đọc thông báo này' };

    var appUrl = '';
    try { appUrl = ScriptApp.getService().getUrl(); } catch (e) {}
    var msg = '🔔 NHẮC ĐỌC THÔNG BÁO\n\n📌 ' + report.title +
        '\n\nBạn chưa đọc thông báo này. Vui lòng xem sớm.' +
        (appUrl ? '\n🔗 ' + appUrl + '?ann_id=' + id : '');

    var sent = 0;
    unread.forEach(function(u) {
        try {
            notifyEmployeeViaTelegram(u.id, msg);
            notifyEmployeeViaZalo(u.id, msg);
            sent++;
        } catch (e) {
            console.error('Reminder error for ' + u.id + ': ' + e.message);
        }
    });
    return { success: true, sent: sent, message: 'Đã gửi nhắc nhở tới ' + sent + ' người' };
}

// Đẩy thông báo ra Telegram/Zalo cho đúng đối tượng nhận
function _pushAnnouncementToChannels(id) {
    try {
        var a = findRowByColumn(SHEET_NAMES.ANNOUNCEMENTS, 'ID', id);
        if (!a) return;

        var appUrl = '';
        try { appUrl = ScriptApp.getService().getUrl(); } catch (e) {}
        var icon = a['Mức độ'] === 'urgent' ? '🚨' : (a['Mức độ'] === 'important' ? '⚠️' : '📢');
        var plain = String(a['Nội dung'] || '').replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim();
        if (plain.length > 300) plain = plain.substring(0, 300) + '...';

        var msg = icon + ' THÔNG BÁO NỘI BỘ\n\n📌 ' + a['Tiêu đề'] + '\n\n' + plain +
            (appUrl ? '\n\n🔗 Xem chi tiết: ' + appUrl + '?ann_id=' + id : '');

        _annRecipientIds(a).forEach(function(uid) {
            try {
                notifyEmployeeViaTelegram(uid, msg);
                notifyEmployeeViaZalo(uid, msg);
            } catch (e) {}
        });
    } catch (err) {
        console.error('Push announcement error: ' + err.message);
    }
}

// Trigger: đăng các thông báo đã tới giờ hẹn
function publishScheduledAnnouncements() {
    ensureAnnouncementSheets();
    var now = new Date();
    var published = 0;
    getSheetData(SHEET_NAMES.ANNOUNCEMENTS).forEach(function(a) {
        if (a['Trạng thái'] !== 'scheduled' || !a['Thời gian đăng']) return;
        if (new Date(a['Thời gian đăng']) > now) return;
        updateRow(SHEET_NAMES.ANNOUNCEMENTS, a._rowIndex, { 'Trạng thái': 'published' });
        _pushAnnouncementToChannels(a['ID']);
        published++;
    });
    return { success: true, published: published };
}

function setupAnnouncementTrigger() {
    ScriptApp.getProjectTriggers().forEach(function(t) {
        if (t.getHandlerFunction() === 'publishScheduledAnnouncements') ScriptApp.deleteTrigger(t);
    });
    ScriptApp.newTrigger('publishScheduledAnnouncements').timeBased().everyMinutes(5).create();
    return { success: true, message: 'Đã bật tự động đăng thông báo hẹn giờ (5 phút/lần)' };
}

// ============ BỘ LỌC ĐÃ LƯU (SAVED SEARCH FILTERS) ============
function getSavedFilters(userId) {
    ensureAnnouncementSheets();
    return getSheetData(SHEET_NAMES.SAVED_FILTERS)
        .filter(function(f) { return f['ID Nhân viên'] === userId; })
        .map(function(f) {
            var config = {};
            try { config = f['Cấu hình'] ? JSON.parse(f['Cấu hình']) : {}; } catch (e) {}
            return { id: f['ID'], name: f['Tên bộ lọc'], config: config, created_at: f['Ngày tạo'] };
        });
}

function saveFilter(userId, name, config) {
    ensureAnnouncementSheets();
    var id = generateId('bl');
    appendRow(SHEET_NAMES.SAVED_FILTERS, {
        'ID': id, 'ID Nhân viên': userId, 'Tên bộ lọc': name,
        'Cấu hình': JSON.stringify(config || {}), 'Ngày tạo': getLocalISOString()
    });
    return { success: true, id: id };
}

function deleteSavedFilter(id) {
    var f = findRowByColumn(SHEET_NAMES.SAVED_FILTERS, 'ID', id);
    if (!f) return { success: false, message: 'Không tìm thấy bộ lọc' };
    deleteRow(SHEET_NAMES.SAVED_FILTERS, f._rowIndex);
    return { success: true };
}