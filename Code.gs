const THEME_MAP = {
  '#FF69B4': { light: '#fbcfe8' }, // Pink
  '#2196F3': { light: '#bbdefb' }, // Blue
  '#3F51B5': { light: '#c5cae9' }, // Navy
  '#4CAF50': { light: '#c8e6c9' }, // Green
  '#FFC107': { light: '#fff9c4' }, // Yellow
  '#9C27B0': { light: '#e1bee7' }, // Purple
  '#FF9800': { light: '#ffe0b2' }, // Orange
  '#F44336': { light: '#ffcdd2' }  // Red
};

function doGet(e) {
  Logger.log('doGet called with params: ' + JSON.stringify(e));
  try {
    var htmlContent = loadLoginPage();
    return HtmlService
      .createHtmlOutput(htmlContent)
      .setTitle('ล็อกอิน - ระบบเช็คชื่อนักเรียนออนไลน์')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    Logger.log('Error in doGet: ' + error.stack);
    throw new Error('ไม่สามารถโหลดหน้า login ได้: ' + error.message);
  }
}

function initializeSheet(sheetName) {
  Logger.log('initializeSheet called: ' + sheetName);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      Logger.log('Created new sheet: ' + sheetName);
    }
    var headers = [];
    if (sheetName === 'subjects') {
      headers = ['id', 'code', 'name'];
    } else if (sheetName === 'teachers') {
      headers = ['id', 'username', 'password', 'name', 'subjectIds', 'classLevels', 'resetRequested', 'subjectClassPairs'];
    } else if (sheetName === 'students') {
      headers = ['id', 'code', 'name', 'class', 'classroom'];
    } else if (sheetName === 'attendance') {
      headers = ['id', 'studentId', 'subjectId', 'date', 'status', 'studentName', 'class', 'classroom', 'teacherName', 'subjectName', 'remark'];
    } else if (sheetName === 'settings') {
      headers = ['key', 'value'];
    }
    var range = sheet.getRange(1, 1, 1, headers.length);
    if (sheet.getLastRow() === 0 || !range.getValues()[0].every((val, i) => val === headers[i])) {
      range.setValues([headers]);
      if (sheetName === 'settings') {
        sheet.appendRow(['system_name', 'ระบบเช็คชื่อนักเรียนออนไลน์ โรงเรียนห้องสื่อครูคอม']);
        sheet.appendRow(['logo_url', 'https://img2.pic.in.th/jsz.png']);
        sheet.appendRow(['header_text', 'ระบบเช็คชื่อนักเรียนออนไลน์ โรงเรียนห้องสื่อครูคอม']);
        sheet.appendRow(['footer_text', '© 2025 ระบบเช็คชื่อนักเรียนออนไลน์ โรงเรียนห้องสื่อครูคอม | กฤษฎา คำมา']);
        sheet.appendRow(['theme_color', '#FF69B4']); // เพิ่มค่าสีเริ่มต้น
      }
      Logger.log('Set headers for sheet: ' + sheetName);
    }
    return sheet;
  } catch (error) {
    Logger.log('Error in initializeSheet: ' + error.message);
    throw new Error('ไม่สามารถสร้างหรือเข้าถึง Sheet ได้');
  }
}

function getInitialData() {
  Logger.log('getInitialData: Bundling all necessary data for the client.');
  try {
    var user = {};
    // ในสถานการณ์จริง ควรมีการตรวจสอบ Session หรือ Token เพื่อระบุผู้ใช้
    // แต่ในโครงสร้างนี้ เราจะส่งข้อมูลทั้งหมดที่จำเป็นไปก่อน
    
    var dataBundle = {
      settings: getSettings(),
      classLevelSettings: getClassLevelSettings(),
      subjects: getData('subjects'),
      students: getData('students'),
      teachers: getData('teachers'),
      attendance: getData('attendance')
    };
    Logger.log('getInitialData: Successfully bundled all data.');
    return dataBundle;
  } catch (error) {
    Logger.log('Error in getInitialData: ' + error.stack);
    throw new Error('ไม่สามารถรวบรวมข้อมูลเริ่มต้นของระบบได้: ' + error.message);
  }
}

function getInitialDataForUser(user) {
    Logger.log('getInitialDataForUser called for role: ' + user.role);
    try {
        var dataBundle = {
            settings: getSettings(),
            classLevelSettings: getClassLevelSettings()
        };

        if (user.role === 'admin') {
            dataBundle.subjects = getData('subjects');
            dataBundle.students = getData('students');
            dataBundle.teachers = getData('teachers');
            dataBundle.attendance = getData('attendance');
        } else if (user.role === 'teacher') {
            var pairs = [];
            try {
                // แยกข้อมูลคู่ วิชา-ชั้นเรียน ที่ครูสอน
                pairs = JSON.parse(user.subjectClassPairs || '[]');
            } catch (e) {
                Logger.log('Invalid subjectClassPairs for teacher: ' + user.id);
            }

            // สร้าง list ของ ID วิชา, ระดับชั้น, และห้องเรียนที่ครูคนนี้สอนทั้งหมด
            var validSubjectIds = pairs.map(p => p.subjectId);
            var validClasses = pairs.flatMap(p => p.classLevels || []);
            var validClassrooms = pairs.flatMap(p => p.classrooms || []);

            dataBundle.subjects = getData('subjects').filter(s => validSubjectIds.includes(s.id));
            
            // ปรับปรุงการกรองนักเรียนให้แม่นยำขึ้น (ดูคำอธิบายเพิ่มเติมด้านล่าง)
            dataBundle.students = getData('students').filter(student => {
                // ตรวจสอบว่านักเรียนคนนี้อยู่ในชั้นเรียนและห้องที่ครูสอนจริงหรือไม่
                return pairs.some(pair => 
                    (pair.classLevels || []).includes(student.class) && 
                    (pair.classrooms || []).includes(student.classroom || '')
                );
            });

            dataBundle.teachers = []; // ครูไม่จำเป็นต้องเห็นข้อมูลครูคนอื่น
            
            // แก้ไขจุดที่ผิด: เอา 'return' ออกจาก arrow function แบบย่อ
            dataBundle.attendance = getData('attendance').filter(a =>
                validClasses.includes(a.class) &&
                validClassrooms.includes(a.classroom || '') &&
                validSubjectIds.includes(String(a.subjectId)) &&
                a.teacherName === user.name
            );
        }

        Logger.log('getInitialDataForUser: Successfully bundled data for ' + user.role);
        return dataBundle;
    } catch (error) {
        Logger.log('Error in getInitialDataForUser: ' + error.stack);
        throw new Error('ไม่สามารถรวบรวมข้อมูลได้: ' + error.message);
    }
}

function initializeClassLevelsSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = 'class_levels_config';
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        var headers = ['groupName', 'levelName'];
        sheet.getRange(1, 1, 1, 2).setValues([headers]);
        var levels = [
            ['ระดับปฐมวัย', 'เตรียมอนุบาล'], ['ระดับปฐมวัย', 'อนุบาล 1'], ['ระดับปฐมวัย', 'อนุบาล 2'], ['ระดับปฐมวัย', 'อนุบาล 3'],
            ['ระดับประถมศึกษาตอนต้น', 'ประถมศึกษาปีที่ 1'], ['ระดับประถมศึกษาตอนต้น', 'ประถมศึกษาปีที่ 2'], ['ระดับประถมศึกษาตอนต้น', 'ประถมศึกษาปีที่ 3'],
            ['ระดับประถมศึกษาตอนปลาย', 'ประถมศึกษาปีที่ 4'], ['ระดับประถมศึกษาตอนปลาย', 'ประถมศึกษาปีที่ 5'], ['ระดับประถมศึกษาตอนปลาย', 'ประถมศึกษาปีที่ 6'],
            ['ระดับมัธยมศึกษาตอนต้น', 'มัธยมศึกษาปีที่ 1'], ['ระดับมัธยมศึกษาตอนต้น', 'มัธยมศึกษาปีที่ 2'], ['ระดับมัธยมศึกษาตอนต้น', 'มัธยมศึกษาปีที่ 3'],
            ['ระดับมัธยมศึกษาตอนปลาย', 'มัธยมศึกษาปีที่ 4'], ['ระดับมัธยมศึกษาตอนปลาย', 'มัธยมศึกษาปีที่ 5'], ['ระดับมัธยมศึกษาตอนปลาย', 'มัธยมศึกษาปีที่ 6'],
            ['ระดับอาชีวศึกษา', 'ปวช. ปี 1'], ['ระดับอาชีวศึกษา', 'ปวช. ปี 2'], ['ระดับอาชีวศึกษา', 'ปวช. ปี 3'], ['ระดับอาชีวศึกษา', 'ปวส. ปี 1'], ['ระดับอาชีวศึกษา', 'ปวส. ปี 2'],
            ['ระดับอุดมศึกษา', 'ปี 1'], ['ระดับอุดมศึกษา', 'ปี 2'], ['ระดับอุดมศึกษา', 'ปี 3'], ['ระดับอุดมศึกษา', 'ปี 4'], ['ระดับอุดมศึกษา', 'ปี 5'], ['ระดับอุดมศึกษา', 'ปี 6']
        ];
        sheet.getRange(2, 1, levels.length, 2).setValues(levels);
    }
    return sheet;
}

function getClassLevelSettings() {
    try {
        var configSheet = initializeClassLevelsSheet();
        var allLevelsData = configSheet.getDataRange().getValues();
        var settings = getSettings();

        var allLevelsGrouped = {};
        for (var i = 1; i < allLevelsData.length; i++) {
            var group = allLevelsData[i][0];
            var level = allLevelsData[i][1];
            if (!allLevelsGrouped[group]) {
                allLevelsGrouped[group] = [];
            }
            allLevelsGrouped[group].push(level);
        }

        var enabledLevels = [];
        // ตรวจสอบว่ามีค่า setting นี้อยู่หรือไม่ และเป็น JSON ที่ถูกต้องหรือไม่
        if (settings.enabled_class_levels) {
            try {
                var parsedLevels = JSON.parse(settings.enabled_class_levels);
                if (Array.isArray(parsedLevels)) {
                     enabledLevels = parsedLevels;
                }
            } catch (e) {
                Logger.log('Could not parse enabled_class_levels, defaulting to empty. Value was: ' + settings.enabled_class_levels);
                enabledLevels = []; // หาก parse ไม่ได้ ให้ใช้ค่าว่าง
            }
        }

        return {
            allSettings: settings,
            allLevels: allLevelsGrouped,
            enabledLevels: enabledLevels
        };
    } catch (e) {
        Logger.log('Error in getClassLevelSettings: ' + e.stack);
        throw new Error('ไม่สามารถดึงข้อมูลการตั้งค่าระดับชั้นได้');
    }
}

function login(username, password) {
  Logger.log('login called for teacher/admin: ' + username);
  try {
    username = username ? username.trim() : '';
    password = password ? password.trim() : '';

    if (username === 'admin' && password === 'admin1234') {
      Logger.log('Admin login successful');
      return { id: 'admin', name: 'Administrator', role: 'admin' };
    }

    var sheet = initializeSheet('teachers');
    var data = sheet.getDataRange().getValues();
    Logger.log('Teachers sheet data length: ' + data.length);

    if (data.length <= 1) {
      Logger.log('Login failed: No teacher data in sheet');
      throw new Error('ไม่มีข้อมูลครูในระบบ กรุณาติดต่อผู้ดูแลระบบ');
    }

    for (var i = 1; i < data.length; i++) {
      if (data[i].length < 8) {
        Logger.log('Invalid row format at index ' + i + ': ' + JSON.stringify(data[i]));
        continue;
      }
      var sheetUsername = data[i][1] ? data[i][1].toString().trim() : '';
      var sheetPassword = data[i][2] ? data[i][2].toString().trim() : '';
      if (sheetUsername === username && sheetPassword === password) {
        var subjectClassPairs = data[i][7] ? data[i][7].toString().trim() : '[]';
        var parsedPairs = [];
        try {
          parsedPairs = JSON.parse(subjectClassPairs);
        } catch (e) {
          Logger.log('Invalid subjectClassPairs for user ' + username + ': ' + subjectClassPairs);
        }
        var classrooms = parsedPairs.flatMap(p => p.classrooms || []);
        Logger.log('Teacher login successful: ' + username);
        return {
          id: data[i][0],
          name: data[i][3] || 'Unknown',
          role: 'teacher',
          subjectIds: data[i][4] ? data[i][4].toString().split(',') : [],
          classLevels: data[i][5] ? data[i][5].toString().split(',') : [],
          subjectClassPairs: subjectClassPairs,
          classrooms: classrooms
        };
      }
    }
    Logger.log('Login failed: Invalid credentials for username: ' + username);
    throw new Error('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง');
  } catch (error) {
    Logger.log('Error in login: ' + error.message);
    throw new Error(error.message || 'การล็อกอินล้มเหลว กรุณาลองใหม่');
  }
}


function loadIndexPage() {
  Logger.log('loadIndexPage called');
  try {
    var settings = getSettings();
    var htmlContent = HtmlService.createHtmlOutputFromFile('index').getContent();

    var placeholders = {
      '{{SYSTEM_NAME}}': settings.system_name || 'ระบบเช็คชื่อนักเรียนออนไลน์ โรงเรียนห้องสื่อครูคอม',
      '{{LOGO_URL}}': settings.logo_url || 'https://img2.pic.in.th/jsz.png',
      '{{HEADER_TEXT}}': settings.header_text || 'ระบบเช็คชื่อนักเรียนออนไลน์ โรงเรียนห้องสื่อครูคอม',
      '{{FOOTER_TEXT}}': settings.footer_text || '© 2025 ระบบเช็คชื่อนักเรียนออนไลน์ โรงเรียนห้องสื่อครูคอม | กฤษฎา คำมา',
      '{{THEME_COLOR}}': settings.theme_color || '#FF69B4' // เพิ่ม placeholder สำหรับสี
    };

    Object.keys(placeholders).forEach(function(key) {
      htmlContent = htmlContent.replace(new RegExp(key, 'g'), placeholders[key]);
    });

    if (!htmlContent || htmlContent.trim() === '') {
      throw new Error('เนื้อหา HTML ว่างเปล่า');
    }

    Logger.log('Index page loaded successfully');
    return htmlContent;
  } catch (error) {
    Logger.log('Error in loadIndexPage: ' + error.stack);
    throw new Error('ไม่สามารถโหลดหน้า index ได้: ' + error.message);
  }
}

function loadLoginPage() {
    Logger.log('loadLoginPage called');
    try {
        var settings = getSettings();
        var html = HtmlService.createHtmlOutputFromFile('login').getContent();

        var primaryColor = settings.theme_color || '#FF69B4';
        var lightColor = THEME_MAP[primaryColor] ? THEME_MAP[primaryColor].light : '#fbcfe8';

        html = html.replace(/{{SYSTEM_NAME}}/g, settings.system_name || 'ระบบเช็คชื่อนักเรียนออนไลน์ โรงเรียนห้องสื่อครูคอม');
        html = html.replace(/{{LOGO_URL}}/g, settings.logo_url || 'https://img2.pic.in.th/jsz.png');
        html = html.replace(/{{HEADER_TEXT}}/g, settings.header_text || 'ระบบเช็คชื่อนักเรียนออนไลน์ โรงเรียนห้องสื่อครูคอม');
        html = html.replace(/{{FOOTER_TEXT}}/g, settings.footer_text || '© 2025 ระบบเช็คชื่อนักเรียนออนไลน์ โรงเรียนห้องสื่อครูคอม | กฤษฎา คำมา');
        html = html.replace(/{{THEME_COLOR}}/g, primaryColor);
        html = html.replace(/{{THEME_LIGHT_COLOR}}/g, lightColor);

        return html;
    } catch (error) {
        Logger.log('Error in loadLoginPage: ' + error.message);
        throw new Error('ไม่สามารถโหลดหน้า login ได้');
    }
}


function loginAndLoad(username, password) {
    Logger.log('loginAndLoad called: ' + username);
    try {
        var user = login(username, password);
        var htmlContent = loadIndexPage();
        var initialData = getInitialDataForUser(user);
        return {
            user: user,
            htmlContent: htmlContent,
            initialData: initialData
        };
    } catch (error) {
        Logger.log('Error in loginAndLoad: ' + error.stack);
        throw new Error(error.message || 'การล็อกอินล้มเหลว กรุณาลองใหม่');
    }
}

function studentLoginAndLoad(studentCode) {
    Logger.log('studentLoginAndLoad called for student code: ' + studentCode);
    try {
        if (!studentCode) {
            throw new Error('กรุณากรอกรหัสนักเรียน');
        }

        var allStudents = getData('students');
        var student = allStudents.find(function(s) {
            return s.code === studentCode.trim();
        });

        if (!student) {
            throw new Error('ไม่พบรหัสนักเรียนนี้ในระบบ');
        }

        var user = {
            id: student.id,
            name: student.name,
            role: 'student',
            code: student.code,
            class: student.class,
            classroom: student.classroom
        };

        var allAttendance = getData('attendance');
        var studentAttendance = allAttendance.filter(function(a) {
            return a.studentId === student.id;
        });

        // --- START: ส่วนที่แก้ไข ---
        var allSubjects = getData('subjects');
        var allTeachers = getData('teachers');
        var relevantSubjectIds = new Set();

        // ค้นหารายวิชาทั้งหมดที่เกี่ยวข้องกับระดับชั้นและห้องของนักเรียน
        allTeachers.forEach(function(teacher) {
            var pairs = [];
            try {
                pairs = JSON.parse(teacher.subjectClassPairs || '[]');
            } catch (e) { /* Ignore parsing errors */ }
            
            pairs.forEach(function(pair) {
                var teachesThisClass = (pair.classLevels || []).includes(student.class);
                var teachesThisClassroom = (pair.classrooms || []).includes(student.classroom || '');
                
                if (teachesThisClass && teachesThisClassroom) {
                    relevantSubjectIds.add(pair.subjectId);
                }
            });
        });

        // กรองรายวิชาจาก Set ที่สร้างขึ้น
        var relevantSubjects = allSubjects.filter(function(s) {
            return relevantSubjectIds.has(s.id);
        });
        
        var initialData = {
            settings: getSettings(), // เพิ่มการตั้งค่าระบบทั้งหมด
            classLevelSettings: getClassLevelSettings(), // เพิ่มการตั้งค่าระดับชั้น
            attendance: studentAttendance,
            subjects: relevantSubjects,
            teachers: allTeachers // ส่งข้อมูลครูทั้งหมดเผื่อการแสดงชื่อครูผู้สอน
        };
        // --- END: ส่วนที่แก้ไข ---

        var htmlContent = loadIndexPage();

        return {
            user: user,
            htmlContent: htmlContent,
            initialData: initialData
        };

    } catch (error) {
        Logger.log('Error in studentLoginAndLoad: ' + error.stack);
        throw new Error(error.message || 'การล็อกอินล้มเหลว กรุณาลองใหม่');
    }
}

function requestPasswordReset(username) {
  Logger.log('requestPasswordReset called: ' + username);
  try {
    var sheet = initializeSheet('teachers');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === username) {
        sheet.getRange(i + 1, 7).setValue(true);
        Logger.log('Password reset requested for: ' + username);
        return;
      }
    }
    Logger.log('Password reset failed: Username not found');
    throw new Error('ไม่พบชื่อผู้ใช้');
  } catch (error) {
    Logger.log('Error in requestPasswordReset: ' + error.message);
    throw new Error('ไม่พบชื่อผู้ใช้ การขอรีเซ็ตรหัสผ่านล้มเหลว');
  }
}

function resetTeacherPassword(teacherId, newPassword) {
  Logger.log('resetTeacherPassword called: ' + teacherId);
  try {
    var sheet = initializeSheet('teachers');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === teacherId) {
        sheet.getRange(i + 1, 3).setValue(newPassword);
        sheet.getRange(i + 1, 7).setValue(false);
        Logger.log('Password reset for teacher: ' + teacherId);
        return;
      }
    }
    Logger.log('Password reset failed: Teacher not found');
    throw new Error('ไม่พบครู');
  } catch (error) {
    Logger.log('Error in resetTeacherPassword: ' + error.message);
    throw new Error('การรีเซ็ตรหัสผ่านล้มเหลว');
  }
}

function getData(sheetName) {
  Logger.log('getData called: ' + sheetName);
  try {
    var sheet = initializeSheet(sheetName);
    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    var result = [];

    if (values.length <= 1) {
      Logger.log('No data found in sheet: ' + sheetName);
      return [];
    }

    for (var i = 1; i < values.length; i++) {
      if (!values[i][0]) continue;

      var row = { id: values[i][0] };

      if (sheetName === 'subjects') {
        row.code = values[i][1] ? values[i][1].toString().trim() : '';
        row.name = values[i][2] ? values[i][2].toString().trim() : '';
      } else if (sheetName === 'teachers') {
        row.username = values[i][1] ? values[i][1].toString().trim() : '';
        row.password = values[i][2] ? values[i][2].toString().trim() : '';
        row.name = values[i][3] ? values[i][3].toString().trim() : '';
        row.subjectIds = values[i][4] ? values[i][4].toString().trim() : '';
        row.classLevels = values[i][5] ? values[i][5].toString().trim() : '';
        row.resetRequested = values[i][6] || false;
        row.subjectClassPairs = values[i][7] ? values[i][7].toString().trim() : '';
      } else if (sheetName === 'students') {
        row.code = values[i][1] ? values[i][1].toString().trim() : '';
        row.name = values[i][2] ? values[i][2].toString().trim() : '';
        row.class = values[i][3] ? values[i][3].toString().trim() : '';
        row.classroom = values[i][4] ? values[i][4].toString().trim() : '';
      } else if (sheetName === 'attendance') {
        row.studentId = values[i][1] ? values[i][1].toString().trim() : '';
        row.subjectId = values[i][2] ? values[i][2].toString().trim() : '';
        var dateCell = values[i][3];
        if (dateCell instanceof Date) {
          var tz = Session.getScriptTimeZone();
          row.date = Utilities.formatDate(dateCell, tz, 'yyyy-MM-dd');
        } else {
          row.date = dateCell ? dateCell.toString().trim() : '';
        }
        row.status = values[i][4] ? values[i][4].toString().trim() : '';
        row.studentName = values[i][5] ? values[i][5].toString().trim() : '';
        row.class = values[i][6] ? values[i][6].toString().trim() : '';
        row.classroom = values[i][7] ? values[i][7].toString().trim() : '';
        row.teacherName = values[i][8] ? values[i][8].toString().trim() : '';
        row.subjectName = values[i][9] ? values[i][9].toString().trim() : '';
        row.remark = values[i][10] ? values[i][10].toString().trim() : '';
      } else if (sheetName === 'settings') {
        row.key = values[i][0] ? values[i][0].toString().trim() : '';
        row.value = values[i][1] ? values[i][1].toString().trim() : '';
      }

      result.push(row);
    }

    Logger.log('Data retrieved: ' + sheetName + ', count: ' + result.length);
    return result;
  } catch (error) {
    Logger.log('Error in getData: ' + error.message);
    throw new Error('ไม่สามารถดึงข้อมูลได้: ' + error.message);
  }
}

function addData(sheetName, data) {
    Logger.log('addData called: ' + sheetName);
    try {
        var sheet = initializeSheet(sheetName);
        if (sheetName === 'teachers') {
            if (!/^[a-zA-Z0-9]{3,}$/.test(data.username)) {
                throw new Error('ชื่อผู้ใช้ต้องประกอบด้วยภาษาอังกฤษและตัวเลขอย่างน้อย 3 ตัวอักษร');
            }
            if (data.password && data.password.length < 6) {
                throw new Error('รหัสผ่านต้องมีอย่างน้อย 6 ตัวอักษร');
            }
            var teachers = sheet.getDataRange().getValues();
            for (var i = 1; i < teachers.length; i++) {
                var existingUsername = teachers[i][1] ? teachers[i][1].toString().trim() : '';
                if (existingUsername === data.username.trim()) {
                    throw new Error('ชื่อผู้ใช้นี้มีอยู่แล้วในระบบ');
                }
            }
        } else if (sheetName === 'students') {
            var students = getData('students');
            var codeExists = students.some(s => s.code === data.code);
            if (codeExists) {
                throw new Error('รหัสนักเรียนนี้มีอยู่แล้วในระบบ');
            }
        }
        var id = Utilities.getUuid();
        var row = [id];
        if (sheetName === 'subjects') {
            row.push(data.code || '', data.name || '');
        } else if (sheetName === 'teachers') {
            row.push(
                data.username || '',
                data.password || '',
                data.name || '',
                data.subjectIds || '',
                data.classLevels || '',
                false,
                data.subjectClassPairs || '[]'
            );
        } else if (sheetName === 'students') {
            row.push(
                data.code || '',
                data.name || '',
                data.class || '',
                data.classroom || ''
            );
        }
        sheet.appendRow(row);
        Logger.log('Data added: ' + sheetName + ', id: ' + id);
        return { success: true, id: id };
    } catch (error) {
        Logger.log('Error in addData: ' + error.message);
        throw new Error(error.message || 'ไม่สามารถเพิ่มข้อมูลได้');
    }
}

function updateData(sheetName, id, data) {
    Logger.log('updateData called: ' + sheetName + ', id: ' + id);
    try {
        var sheet = initializeSheet(sheetName);
        var dataRange = sheet.getDataRange();
        var values = dataRange.getValues();
        if (sheetName === 'teachers' && data.username) {
            if (!/^[a-zA-Z0-9]{3,}$/.test(data.username)) {
                throw new Error('ชื่อผู้ใช้ต้องประกอบด้วยภาษาอังกฤษและตัวเลขอย่างน้อย 3 ตัวอักษร');
            }
            if (data.password && data.password.length < 6) {
                throw new Error('รหัสผ่านต้องมีอย่างน้อย 6 ตัวอักษร');
            }
            for (var r = 1; r < values.length; r++) {
                var rowId = values[r][0];
                var existingUsername = values[r][1] ? values[r][1].toString().trim() : '';
                if (rowId !== id && existingUsername === data.username.trim()) {
                    throw new Error('ชื่อผู้ใช้นี้มีอยู่แล้วในระบบ');
                }
            }
        } else if (sheetName === 'students' && data.code) {
            var students = getData('students');
            var codeExists = students.some(s => s.code === data.code && s.id !== id);
            if (codeExists) {
                throw new Error('รหัสนักเรียนนี้มีอยู่ในระบบแล้ว');
            }
        }
        for (var i = 1; i < values.length; i++) {
            if (values[i][0] === id) {
                if (sheetName === 'subjects') {
                    values[i][1] = data.code || '';
                    values[i][2] = data.name || '';
                } else if (sheetName === 'teachers') {
                    values[i][1] = data.username || '';
                    if (data.password) values[i][2] = data.password;
                    values[i][3] = data.name || '';
                    values[i][4] = data.subjectIds || '';
                    values[i][5] = data.classLevels || '';
                    values[i][6] = data.resetRequested || false;
                    values[i][7] = data.subjectClassPairs || '[]';
                } else if (sheetName === 'students') {
                    values[i][1] = data.code || '';
                    values[i][2] = data.name || '';
                    if (data.class) values[i][3] = data.class;
                    if (data.classroom) values[i][4] = data.classroom;
                } else if (sheetName === 'attendance') {
                    values[i][1] = data.studentId || '';
                    values[i][2] = data.subjectId || '';
                    values[i][3] = data.date || '';
                    values[i][4] = data.status || '';
                    values[i][5] = data.studentName || '';
                    values[i][6] = data.class || '';
                    values[i][7] = data.classroom || '';
                    values[i][8] = data.teacherName || '';
                    values[i][9] = data.subjectName || '';
                    values[i][10] = data.remark || '';
                }
                dataRange.setValues(values);
                Logger.log('Data updated: ' + sheetName + ', id: ' + id);
                return { success: true };
            }
        }
        throw new Error('ไม่พบข้อมูลที่ต้องการอัปเดต');
    } catch (error) {
        Logger.log('Error in updateData: ' + error.message);
        throw new Error(error.message || 'การอัปเดตข้อมูลล้มเหลว');
    }
}

function deleteData(sheetName, id) {
  Logger.log('deleteData called: ' + sheetName + ', id: ' + id);
  try {
    var sheet = initializeSheet(sheetName);
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.deleteRow(i + 1);
        
        if (sheetName === 'subjects') {
          var attendanceSheet = initializeSheet('attendance');
          var attendanceData = attendanceSheet.getDataRange().getValues();
          var rowsToDelete = [];
          
          for (var k = 1; k < attendanceData.length; k++) {
            if (attendanceData[k][2] === id) {
              rowsToDelete.push(k + 1);
            }
          }
          
          rowsToDelete.sort((a, b) => b - a).forEach(rowIndex => {
            attendanceSheet.deleteRow(rowIndex);
            Logger.log('Deleted attendance row: ' + rowIndex + ' due to subject deletion');
          });
        } else if (sheetName === 'teachers') {
          var attendanceSheet = initializeSheet('attendance');
          var attendanceData = attendanceSheet.getDataRange().getValues();
          for (var k = attendanceData.length - 1; k >= 1; k--) {
            if (attendanceData[k][8] === data[i][3]) {
              attendanceSheet.deleteRow(k + 1);
              Logger.log('Deleted attendance row: ' + (k + 1) + ' due to teacher deletion');
            }
          }
        } else if (sheetName === 'students') {
          var attendanceSheet = initializeSheet('attendance');
          var attendanceData = attendanceSheet.getDataRange().getValues();
          for (var k = attendanceData.length - 1; k >= 1; k--) {
            if (attendanceData[k][1] === id) {
              attendanceSheet.deleteRow(k + 1);
              Logger.log('Deleted attendance row: ' + (k + 1) + ' due to student deletion');
            }
          }
        }
        
        Logger.log('Data deleted: ' + sheetName + ', id: ' + id);
        return;
      }
    }
    Logger.log('Delete failed: ID not found');
    throw new Error('ไม่พบข้อมูลที่ต้องการลบ');
  } catch (error) {
    Logger.log('Error in deleteData: ' + error.message);
    throw new Error('การลบข้อมูลล้มเหลว: ' + error.message);
  }
}

function saveAllAttendance(records) {
  Logger.log('saveAllAttendance called: %s records', records.length);
  try {
    var sheet = initializeSheet('attendance');
    var data = sheet.getDataRange().getValues();
    var count = 0;
    var savedRecords = []; //--- เพิ่มบรรทัดนี้

    var dataMap = {};
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) {
        dataMap[data[i][0]] = i + 1;
      }
    }

    var newRows = [];
    
    records.forEach(function(r) {
      var isNew = !r.id || !dataMap[r.id];
      var recordId = isNew ? Utilities.getUuid() : r.id;

      var rowData = [
        recordId, r.studentId || '', r.subjectId || '', r.date || '', r.status || '',
        r.studentName || '', r.class || '', r.classroom || '',
        r.teacherName || '', r.subjectName || '', r.remark || ''
      ];
      
      if (isNew) {
        newRows.push(rowData);
      } else {
        var rowIndex = dataMap[r.id];
        var range = sheet.getRange(rowIndex, 1, 1, 11);
        range.setValues([rowData]);
      }
      
      //--- เพิ่มส่วนนี้ เพื่อสร้าง object ที่จะส่งกลับ
      savedRecords.push({
        id: rowData[0],
        studentId: rowData[1],
        subjectId: rowData[2],
        date: rowData[3],
        status: rowData[4],
        studentName: rowData[5],
        class: rowData[6],
        classroom: rowData[7],
        teacherName: rowData[8],
        subjectName: rowData[9],
        remark: rowData[10]
      });
      //--- สิ้นสุดส่วนที่เพิ่ม

      count++;
    });

    if (newRows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 11).setValues(newRows);
    }

    SpreadsheetApp.flush();
    Logger.log('Saved %s records.', count);
    //--- แก้ไขบรรทัดนี้ เพื่อส่งข้อมูลที่บันทึกแล้วกลับไป
    return { success: true, count: count, savedRecords: savedRecords };
  } catch (e) {
    Logger.log('Error in saveAllAttendance: ' + e.message + " stack: " + e.stack);
    throw new Error('การบันทึกเวลาเรียนล้มเหลว: ' + e.message);
  }
}


function importCSV(sheetName, csvContent) {
  Logger.log('importCSV called: ' + sheetName);
  try {
    var sheet = initializeSheet(sheetName);
    var csvData = Utilities.parseCsv(csvContent);
    var count = 0;

    if (sheetName === 'subjects') {
      var existingCodes = getData('subjects').map(s => s.code);
      for (var i = 1; i < csvData.length; i++) {
        try {
          var row = csvData[i];
          if (row.length >= 2 && row[0] && row[1]) {
            if (existingCodes.includes(row[0])) {
              Logger.log('Skipping duplicate subject code: ' + row[0]);
              continue;
            }
            sheet.appendRow([Utilities.getUuid(), row[0], row[1]]);
            existingCodes.push(row[0]);
            count++;
          }
        } catch (e) {
          Logger.log('Error importing row ' + i + ': ' + e.message);
        }
      }
    } else if (sheetName === 'students') {
      var existingCodes = getData('students').map(s => s.code);
      for (var i = 1; i < csvData.length; i++) {
        try {
          var row = csvData[i];
          if (row.length >= 4 && row[0] && row[1] && row[2] && row[3]) {
            if (existingCodes.includes(row[0])) {
              Logger.log('Skipping duplicate student code: ' + row[0]);
              continue;
            }
            sheet.appendRow([
              Utilities.getUuid(),
              row[0],
              row[1],
              row[2],
              row[3]
            ]);
            existingCodes.push(row[0]);
            count++;
          }
        } catch (e) {
          Logger.log('Error importing row ' + i + ': ' + e.message);
        }
      }
    }

    Logger.log('Imported ' + count + ' records to ' + sheetName);
    return { count: count };
  } catch (error) {
    Logger.log('Error in importCSV: ' + error.message);
    throw new Error('การนำเข้าข้อมูลล้มเหลว: ' + error.message);
  }
}

function checkStudentCode(params) {
  Logger.log('checkStudentCode called: ' + JSON.stringify(params));
  try {
    var sheet = initializeSheet('students');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === params.code && (!params.studentId || data[i][0] !== params.studentId)) {
        Logger.log('Student code exists: ' + params.code);
        return { result: false };
      }
    }
    Logger.log('Student code is unique: ' + params.code);
    return { result: true };
  } catch (error) {
    Logger.log('Error in checkStudentCode: ' + error.message);
    throw new Error('การตรวจสอบรหัสนักเรียนล้มเหลว: ' + error.message);
  }
}

function getClassroomReport(params) {
    Logger.log('getClassroomReport called: ' + JSON.stringify(params));
    try {
        var startDate = new Date(params.startDate);
        var endDate = new Date(params.endDate);
        var classLevelFilter = params.classLevel || '';
        var classroomFilter = params.classroom || '';
        var subjectIdFilter = params.subjectId || '';
        var teacherId = params.teacherId || null;

        var attendanceData = getData('attendance');
        var subjects = getData('subjects');
        var teachers = getData('teachers');

        if (teacherId) {
            var teacher = teachers.find(function(t) { return t.id === teacherId; });
            if (teacher && teacher.subjectClassPairs) {
                var pairs = [];
                try {
                    pairs = JSON.parse(teacher.subjectClassPairs);
                } catch (e) {
                    Logger.log('Could not parse subjectClassPairs for teacher ' + teacherId + ': ' + e.message);
                }
                var validClasses = pairs.flatMap(p => p.classLevels || []);
                var validClassrooms = pairs.flatMap(p => p.classrooms || []);
                var validSubjectIds = pairs.map(p => p.subjectId);

                attendanceData = attendanceData.filter(function(a) {
                    return validClasses.includes(a.class) &&
                        validClassrooms.includes(a.classroom || '') &&
                        validSubjectIds.includes(String(a.subjectId));
                });
            }
        }

        var reportMap = {};

        attendanceData.forEach(function(a) {
            var attDateStr = normalizeDate(a.date);
            if (!attDateStr) return;

            var dObj = new Date(attDateStr);
            if (dObj < startDate || dObj > endDate) return;

            // *** CORRECTED LOGIC FOR "ALL" OPTION ***
            if (classLevelFilter && classLevelFilter !== 'all' && a.class !== classLevelFilter) return;

            if (classroomFilter && (a.classroom || '').toString().trim() !== classroomFilter) return;
            if (subjectIdFilter && a.subjectId !== subjectIdFilter) return;
            if (['present', 'absent', 'leave', 'late'].indexOf(a.status) === -1) return;

            var key = attDateStr + '_' + (a.class || '') + '_' + (a.classroom || '') + '_' + a.subjectId;
            if (!reportMap[key]) {
                var subjObj = subjects.find(function(s) { return s.id === a.subjectId; }) || {};
                reportMap[key] = {
                    date: attDateStr,
                    class: a.class || '',
                    classroom: a.classroom || '',
                    subjectId: a.subjectId,
                    subjectCode: subjObj.code || '',
                    subjectName: subjObj.name || '',
                    totalStudents: 0,
                    present: 0,
                    absent: 0,
                    leave: 0,
                    late: 0,
                    studentSet: new Set() // Use a Set to count unique students
                };
            }
            var entry = reportMap[key];

            // Only increment total students if the student hasn't been counted for this entry yet
            if (!entry.studentSet.has(a.studentId)) {
                entry.studentSet.add(a.studentId);
                entry.totalStudents++;
            }

            if (a.status === 'present') entry.present++;
            else if (a.status === 'absent') entry.absent++;
            else if (a.status === 'leave') entry.leave++;
            else if (a.status === 'late') entry.late++;
        });

        // Convert map to array and remove the temporary Set
        var result = Object.values(reportMap).map(function(item) {
            delete item.studentSet;
            return item;
        });

        Logger.log('Classroom report generated with ' + result.length + ' rows.');
        return result;
    } catch (error) {
        Logger.log('Error in getClassroomReport: ' + error.message);
        throw new Error('ไม่สามารถสร้างรายงานห้องเรียนได้: ' + error.message);
    }
}

function getSchoolReport(params) {
    Logger.log('getSchoolReport called: ' + JSON.stringify(params));
    try {
        var startDate = new Date(params.startDate);
        var endDate = new Date(params.endDate);
        var classLevelFilter = params.classLevel || '';
        var classroomFilter = params.classroom || '';
        var subjectIdFilter = params.subjectId || '';

        var attendance = getData('attendance');
        var subjects = getData('subjects');

        var filteredAttendance = attendance.filter(a => {
            var attDate = new Date(a.date);
            return attDate >= startDate && attDate <= endDate;
        });
        
        // *** CORRECTED LOGIC FOR "ALL" OPTION ***
        if (classLevelFilter && classLevelFilter !== 'all') {
            filteredAttendance = filteredAttendance.filter(a => a.class === classLevelFilter);
        }
        if (classroomFilter) {
            filteredAttendance = filteredAttendance.filter(a => (a.classroom || '').toString().trim() === classroomFilter);
        }
        if (subjectIdFilter) {
            filteredAttendance = filteredAttendance.filter(a => a.subjectId === subjectIdFilter);
        }

        var reportMap = {};
        
        // Iterate through the already filtered data
        filteredAttendance.forEach(a => {
            if (['present', 'absent', 'leave', 'late'].indexOf(a.status) === -1) return;

            var key = normalizeDate(a.date) + '_' + a.class + '_' + (a.classroom || '') + '_' + a.subjectId;
            if (!reportMap[key]) {
                var subject = subjects.find(s => s.id === a.subjectId) || { code: 'N/A', name: 'N/A' };
                reportMap[key] = {
                    date: normalizeDate(a.date),
                    subjectCode: subject.code,
                    subjectName: subject.name,
                    classLevel: a.class,
                    classroom: a.classroom || '',
                    studentIds: new Set(),
                    present: 0,
                    absent: 0,
                    leave: 0,
                    late: 0,
                };
            }
            reportMap[key].studentIds.add(a.studentId);
            if (reportMap[key].hasOwnProperty(a.status)) {
                reportMap[key][a.status]++;
            }
        });

        var result = Object.values(reportMap).map(item => {
            item.totalStudents = item.studentIds.size;
            delete item.studentIds;
            return item;
        });

        Logger.log('School report generated: ' + result.length);
        return result;
    } catch (error) {
        Logger.log('Error in getSchoolReport: ' + error.message);
        throw new Error('ไม่สามารถสร้างรายงานภาพรวมโรงเรียนได้: ' + error.message);
    }
}

function getSettings() {
  Logger.log('getSettings called');
  try {
    var sheet = initializeSheet('settings');
    var data = getData('settings');
    var settings = {};
    data.forEach(row => {
      settings[row.key] = row.value;
    });
    Logger.log('Settings retrieved: ' + JSON.stringify(settings));
    return settings;
  } catch (error) {
    Logger.log('Error in getSettings: ' + error.message);
    throw new Error('ไม่สามารถดึงการตั้งค่าได้: ' + error.message);
  }
}

function saveSettings(data) {
    Logger.log('saveSettings called with data: ' + JSON.stringify(data));
    try {
        if (!data || !data.admin_password) {
            return { success: false, message: 'ข้อมูลรหัสผ่านผู้ดูแลระบบไม่ครบถ้วน' };
        }
        if (data.admin_password !== 'admin1234') {
            return { success: false, message: 'รหัสผ่านผู้ดูแลระบบไม่ถูกต้อง' };
        }
        if (data.hasOwnProperty('enabled_class_levels')) {
            var currentSettings = getSettings();
            var currentEnabledLevels = [];
            if (currentSettings.enabled_class_levels) {
                try {
                    currentEnabledLevels = JSON.parse(currentSettings.enabled_class_levels);
                } catch(e) {
                    Logger.log('Could not parse enabled_class_levels, defaulting to empty. Error: ' + e.message);
                }
            }
            var newEnabledLevels = JSON.parse(data.enabled_class_levels);
            var levelsToBeDisabled = currentEnabledLevels.filter(level => !newEnabledLevels.includes(level));
            if (levelsToBeDisabled.length > 0) {
                var allStudents = getData('students');
                var conflictingLevels = levelsToBeDisabled.filter(level => 
                    allStudents.some(student => student.class === level)
                );
                if (conflictingLevels.length > 0) {
                    return { success: false, message: 'ไม่สามารถปิดระดับชั้น: ' + conflictingLevels.join(', ') + ' ได้ เพราะยังมีข้อมูลนักเรียนอยู่' };
                }
            }
        }
        var sheet = initializeSheet('settings');
        var dataRange = sheet.getDataRange();
        var values = dataRange.getValues();
        var settingsInSheet = {};
        for (var i = 1; i < values.length; i++) {
            if (values[i][0]) {
                settingsInSheet[values[i][0]] = i + 1;
            }
        }
        for (var key in data) {
            if (data.hasOwnProperty(key) && key !== 'admin_password') {
                var valueToSave = data[key];
                var rowNum = settingsInSheet[key];
                if (rowNum) {
                    sheet.getRange(rowNum, 2).setValue(valueToSave);
                } else {
                    sheet.appendRow([key, valueToSave]);
                }
            }
        }
        SpreadsheetApp.flush(); // บังคับบันทึกข้อมูลลง Sheet
        Logger.log('Settings saved successfully');
        return { success: true };
    } catch (error) {
        Logger.log('Error in saveSettings: ' + error.stack);
        return { success: false, message: error.message || 'การบันทึกการตั้งค่าล้มเหลว' };
    }
}

function normalizeDate(dateStr) {
  try {
    const date = new Date(dateStr);
    if (isNaN(date.getTime())) {
      return '';
    }
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch (e) {
    return '';
  }
}
