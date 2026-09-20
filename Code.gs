/**
 * ==============================================================================
 * Google Apps Script (Code.gs) - KKU FIS Union Backend (Sheets & Drive Complete Sync)
 * สโมสรนักศึกษาคณะสหวิทยาการ มหาวิทยาลัยขอนแก่น (KKU FIS Union)
 * ==============================================================================
 * 
 * คำแนะนำการติดตั้งและการใช้งาน:
 * 1. เปิด Google Sheets ที่ต้องการใช้เก็บข้อมูลสโมสรนักศึกษา
 * 2. ไปที่เมนู "ส่วนขยาย" (Extensions) -> "Apps Script"
 * 3. ลบโค้ดเก่าในไฟล์ Code.gs ทั้งหมด แล้ววางชุดโค้ดด้านล่างนี้ลงไปแทน
 * 4. กดปุ่มบันทึก (Ctrl + S)
 * 5. กดปุ่ม "ทำให้ใช้งานได้อย่างเป็นทางการ" (Deploy) -> "การทำให้ใช้งานได้อย่างเป็นทางการรายการใหม่" (New deployment)
 * 6. ตั้งค่าการเผยแพร่:
 *    - ประเภท (Select type): "เว็บแอป" (Web app)
 *    - ผู้เรียกใช้การทำงาน (Execute as): "ตัวฉัน" (Me)
 *    - ผู้ที่มีสิทธิ์เข้าถึง (Who has access): "ทุกคน" (Anyone)  <-- *** สำคัญมาก ***
 * 7. กด "ทำให้ใช้งานได้อย่างเป็นทางการ" (Deploy) แล้วให้สิทธิ์เข้าถึง (Authorize Access) ให้เรียบร้อย
 */

function doGet(e) {
  try {
    var params = e ? e.parameter : {};
    var action = params.action;

    // 1.1 ดึงข้อมูลนักศึกษาจากรหัสนักศึกษา (getStudent)
    if (action === 'getStudent') {
      var searchId = (params.studentId || '').replace(/[^0-9]/g, '');
      var ss = getSpreadsheet();
      var found = null;

      // 1. ค้นหาจากชีท "ฐานข้อมูลนักศึกษา" ก่อน
      var dbSheet = ss ? ss.getSheetByName("ฐานข้อมูลนักศึกษา") : null;
      if (dbSheet && searchId) {
        var lastRowDb = dbSheet.getLastRow();
        if (lastRowDb > 1) {
          var dbValues = dbSheet.getRange(2, 1, lastRowDb - 1, 7).getValues();
          for (var k = dbValues.length - 1; k >= 0; k--) {
            var dbRow = dbValues[k];
            var sIdDb = (dbRow[0] || '').toString().replace(/[^0-9]/g, '');
            if (sIdDb === searchId) {
              found = {
                studentId: dbRow[0],
                name: dbRow[1],
                major: dbRow[2],
                year: (dbRow[3] || '').toString().replace(/[^0-9]/g, ''),
                phone: dbRow[4],
                medical: dbRow[5]
              };
              break;
            }
          }
        }
      }

      // 2. ถ้าไม่พบ ให้ค้นหาจากชีท "ข้อมูลการลงทะเบียนกิจกรรม"
      if (!found && ss) {
        var regSheet = ss.getSheetByName("ข้อมูลการลงทะเบียนกิจกรรม");
        if (regSheet && searchId) {
          var lastRow = regSheet.getLastRow();
          if (lastRow > 1) {
            var values = regSheet.getRange(2, 1, lastRow - 1, 9).getValues();
            for (var i = values.length - 1; i >= 0; i--) {
              var row = values[i];
              var sId = (row[2] || '').toString().replace(/[^0-9]/g, '');
              if (sId === searchId) {
                found = {
                  studentId: row[2],
                  name: row[3],
                  major: row[4],
                  year: (row[5] || '').toString().replace(/[^0-9]/g, ''),
                  phone: row[6],
                  medical: row[7]
                };
                break;
              }
            }
          }
        }
      }

      // 3. ถ้ายังไม่พบ ให้ค้นหาจากชีท "คำร้องขอยืมอุปกรณ์"
      if (!found && ss) {
        var bSheet = ss.getSheetByName("คำร้องขอยืมอุปกรณ์");
        if (bSheet && searchId) {
          var lastRowB = bSheet.getLastRow();
          if (lastRowB > 1) {
            var bValues = bSheet.getRange(2, 1, lastRowB - 1, 14).getValues();
            for (var j = bValues.length - 1; j >= 0; j--) {
              var bRow = bValues[j];
              var bsId = (bRow[2] || '').toString().replace(/[^0-9]/g, '');
              if (bsId === searchId) {
                found = {
                  studentId: bRow[2],
                  name: bRow[3],
                  major: bRow[4],
                  year: (bRow[5] || '').toString().replace(/[^0-9]/g, ''),
                  phone: bRow[6]
                };
                break;
              }
            }
          }
        }
      }

      return createJsonResponse({ status: 'success', student: found });
    }

    // 1.2 หน้าเว็บขอข้อมูลสถานะระบบ (getState) -> อ่านจาก Google Sheets ก่อนเป็นอันดับแรกเพื่อให้การแก้ไข/ลบในชีทแสดงผลหน้าระบบทันที
    if (action === 'getState') {
      var ss = getSpreadsheet();
      var state = ss ? getStateFromSheets(ss) : null;
      if (!state) {
        state = getStateFromProperties();
      }
      if (!state) {
        state = getStateFromDriveBackup();
      }
      return createJsonResponse({ status: 'success', state: state });
    }

    // 1.3 ดึงข้อมูลรายชื่อผู้ลงทะเบียนกิจกรรม (getActivityRegistrations)
    if (action === 'getActivityRegistrations') {
      var targetTitle = (params.activityTitle || '').toString().trim();
      var ss = getSpreadsheet();
      var registrations = [];

      if (ss && targetTitle) {
        var regSheet = ss.getSheetByName("ข้อมูลการลงทะเบียนกิจกรรม");
        if (regSheet) {
          var lastRow = regSheet.getLastRow();
          if (lastRow > 1) {
            var values = regSheet.getRange(2, 1, lastRow - 1, 9).getValues();
            for (var r = 0; r < values.length; r++) {
              var row = values[r];
              var actTitle = (row[0] || '').toString().trim();
              if (actTitle.toLowerCase() === targetTitle.toLowerCase() || targetTitle === 'all') {
                registrations.push({
                  activityTitle: row[0],
                  timestamp: row[1],
                  studentId: row[2],
                  name: row[3],
                  major: row[4],
                  year: row[5],
                  phone: row[6],
                  medical: row[7],
                  customAnswers: row[8]
                });
              }
            }
          }
        }
      }

      return createJsonResponse({ status: 'success', registrations: registrations });
    }

    // 2. หน้าเว็บขอซิงค์สถิติจำนวนผู้สมัครจริงจากชีทลงทะเบียนกิจกรรม (syncApplicantsFromGoogleSheets)
    var ss = getSpreadsheet();
    var stats = {};
    if (ss) {
      var regSheet = ss.getSheetByName("ข้อมูลการลงทะเบียนกิจกรรม");
      if (regSheet) {
        var lastRow = regSheet.getLastRow();
        if (lastRow > 1) {
          var data = regSheet.getRange(2, 1, lastRow - 1, 1).getValues();
          data.forEach(function(row) {
            var title = (row[0] || '').toString().trim();
            if (title) {
              stats[title] = (stats[title] || 0) + 1;
            }
          });
        }
      }
    }

    return createJsonResponse({ status: 'success', stats: stats });

  } catch (err) {
    return createJsonResponse({ status: 'error', message: err.toString() });
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return createJsonResponse({ status: 'error', message: 'No post data received' });
    }

    var data = JSON.parse(e.postData.contents);

    // 1. บันทึกโครงสร้างข้อมูลระบบทั้งหมด (saveState) -> Google Sheets & Google Drive
    if (data.action === 'saveState') {
      var state = data.state;
      saveStateToProperties(state);

      // 1.1 บันทึกลง Google Sheets
      saveAllStateToSheets(state);

      // 1.2 บันทึกสำรองลง Google Drive
      saveDriveBackup(state);

      return createJsonResponse({ status: 'success', message: 'State saved to Google Sheets & Google Drive successfully' });
    }

    // 2. คำร้องขอยืมอุปกรณ์สโมสรฯ (borrowRequest) -> Google Sheets & Google Drive
    if (data.action === 'borrowRequest' && data.request) {
      var req = data.request;
      var ss = getSpreadsheet();
      var sheetName = "คำร้องขอยืมอุปกรณ์";
      var sheet = getOrCreateSheet(ss, sheetName);

      if (sheet.getLastRow() === 0) {
        var headers = [
          "รหัสคำร้อง", "วัน-เวลาที่ยื่นคำร้อง", "รหัสนักศึกษา", "ชื่อ-นามสกุล",
          "สาขาวิชา", "ชั้นปี", "เบอร์โทรศัพท์/Line", "อุปกรณ์ที่ยืม",
          "จำนวนที่ยืม", "วันที่ยืม", "กำหนดวันคืน", "โครงการ/สังกัด",
          "วัตถุประสงค์", "สถานะคำร้อง"
        ];
        createSheetHeader(sheet, headers);
      }

      var rowData = [
        req.id || '-', req.createdAt || '-', req.studentId || '-', req.studentName || '-',
        req.major || '-', req.year ? 'ปี ' + req.year : '-', req.phone || '-', req.equipmentName || '-',
        req.qty || 1, req.borrowDate || '-', req.returnDate || '-', req.orgName || '-',
        req.purpose || '-', req.status || 'รออนุมัติ'
      ];

      sheet.appendRow(rowData);

      // อัปเดตฐานข้อมูลนักศึกษา (ชีท "ฐานข้อมูลนักศึกษา")
      try {
        updateOrAddStudentProfile(ss, {
          studentId: req.studentId,
          name: req.studentName,
          major: req.major,
          year: req.year,
          phone: req.phone
        });
      } catch (e) {
        Logger.log("ไม่สามารถอัปเดตฐานข้อมูลนักศึกษาได้: " + e.toString());
      }

      // สำรองข้อมูลคำร้องลง Google Drive
      saveBorrowRequestToDrive(req);

      return createJsonResponse({ status: 'success', message: 'Borrow request recorded to Sheets & Drive' });
    }

    // 3. แก้ไขข้อมูลผู้ลงทะเบียนกิจกรรม (updateRegistration) -> Google Sheets
    if (data.action === 'updateRegistration') {
      var targetTitle = (data.activityTitle || '').toString().trim();
      var targetStudentId = (data.oldStudentId || data.studentId || '').toString().trim();
      var targetTimestamp = (data.timestamp || '').toString().trim();

      var ss = getSpreadsheet();
      var regSheet = ss ? ss.getSheetByName("ข้อมูลการลงทะเบียนกิจกรรม") : null;
      var updated = false;

      if (regSheet) {
        var lastRow = regSheet.getLastRow();
        if (lastRow > 1) {
          var values = regSheet.getRange(2, 1, lastRow - 1, 9).getValues();
          for (var u = 0; u < values.length; u++) {
            var rTitle = (values[u][0] || '').toString().trim();
            var rTime = (values[u][1] || '').toString().trim();
            var rStudentId = (values[u][2] || '').toString().trim();

            var matchesTitle = (rTitle.toLowerCase() === targetTitle.toLowerCase());
            var matchesStudent = (rStudentId === targetStudentId || (targetTimestamp && rTime === targetTimestamp));

            if (matchesTitle && matchesStudent) {
              var rowIndex = u + 2;
              var newYear = data.year ? (data.year.toString().indexOf('ปี') !== -1 ? data.year : 'ปี ' + data.year) : '-';
              regSheet.getRange(rowIndex, 3, 1, 7).setValues([[
                data.studentId || rStudentId,
                data.name || '-',
                data.major || '-',
                newYear,
                data.phone || '-',
                data.medical || '-',
                data.customAnswers || '-'
              ]]);
              updated = true;
              break;
            }
          }
        }
      }

      if (updated) {
        try { updateOrAddStudentProfile(ss, { studentId: data.studentId, name: data.name, major: data.major, year: data.year, phone: data.phone, medical: data.medical }); } catch (e) {}
        try { updateActivitySummarySheet(ss); } catch (e) {}
        return createJsonResponse({ status: 'success', message: 'แก้ไขข้อมูลผู้ลงทะเบียนเรียบร้อยแล้ว' });
      } else {
        return createJsonResponse({ status: 'error', message: 'ไม่พบรายการผู้ลงทะเบียนที่ต้องการแก้ไข' });
      }
    }

    // 4. ลบข้อมูลผู้ลงทะเบียนกิจกรรม (deleteRegistration) -> Google Sheets
    if (data.action === 'deleteRegistration') {
      var targetTitle = (data.activityTitle || '').toString().trim();
      var targetStudentId = (data.studentId || '').toString().trim();
      var targetTimestamp = (data.timestamp || '').toString().trim();

      var ss = getSpreadsheet();
      var regSheet = ss ? ss.getSheetByName("ข้อมูลการลงทะเบียนกิจกรรม") : null;
      var deleted = false;

      if (regSheet) {
        var lastRow = regSheet.getLastRow();
        if (lastRow > 1) {
          var values = regSheet.getRange(2, 1, lastRow - 1, 9).getValues();
          for (var d = values.length - 1; d >= 0; d--) {
            var rTitle = (values[d][0] || '').toString().trim();
            var rTime = (values[d][1] || '').toString().trim();
            var rStudentId = (values[d][2] || '').toString().trim();

            var matchesTitle = (rTitle.toLowerCase() === targetTitle.toLowerCase());
            var matchesStudent = (rStudentId === targetStudentId || (targetTimestamp && rTime === targetTimestamp));

            if (matchesTitle && matchesStudent) {
              regSheet.deleteRow(d + 2);
              deleted = true;
              break;
            }
          }
        }
      }

      if (deleted) {
        try { updateActivitySummarySheet(ss); } catch (e) {}
        return createJsonResponse({ status: 'success', message: 'ลบข้อมูลผู้ลงทะเบียนเรียบร้อยแล้ว' });
      } else {
        return createJsonResponse({ status: 'error', message: 'ไม่พบรายการผู้ลงทะเบียนที่ต้องการลบ' });
      }
    }

    // 5. นักศึกษาลงทะเบียนเข้าร่วมกิจกรรม (registerActivity) -> Google Sheets & Google Drive
    var activityTitle = data.activityTitle || (data.data ? data.data.activityTitle : 'กิจกรรมทั่วไป');
    var regInfo = data.data || data;

    var ss = getSpreadsheet();
    var regSheetName = "ข้อมูลการลงทะเบียนกิจกรรม";
    var sheet = getOrCreateSheet(ss, regSheetName);

    if (sheet.getLastRow() === 0) {
      var headers = [
        "ชื่อกิจกรรม", "วัน-เวลาที่ลงทะเบียน", "รหัสนักศึกษา", "ชื่อ-นามสกุล", "สาขาวิชา",
        "ชั้นปี", "เบอร์โทรศัพท์", "โรคประจำตัว/แพ้อาหาร", "คำตอบเพิ่มเติม"
      ];
      createSheetHeader(sheet, headers);
    }

    var timestamp = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss");
    var rowData = [
      activityTitle,
      timestamp,
      regInfo.studentId || regInfo.student_id || '-',
      regInfo.name || '-',
      regInfo.major || '-',
      regInfo.year ? 'ปี ' + regInfo.year : '-',
      regInfo.phone || '-',
      regInfo.medical || '-',
      regInfo.customAnswers || '-'
    ];

    sheet.appendRow(rowData);

    // อัปเดตฐานข้อมูลนักศึกษา (ชีท "ฐานข้อมูลนักศึกษา")
    try {
      updateOrAddStudentProfile(ss, {
        studentId: regInfo.studentId || regInfo.student_id,
        name: regInfo.name,
        major: regInfo.major,
        year: regInfo.year,
        phone: regInfo.phone,
        medical: regInfo.medical
      });
    } catch (e) {
      Logger.log("ไม่สามารถอัปเดตฐานข้อมูลนักศึกษาได้: " + e.toString());
    }

    // อัปเดตชีทสรุปผู้เข้าร่วมกิจกรรม (แยกสาขา และชั้นปี)
    try {
      updateActivitySummarySheet(ss);
    } catch (e) {
      Logger.log("ไม่สามารถอัปเดตชีทสรุปผู้เข้าร่วมกิจกรรมได้: " + e.toString());
    }

    // สำรองข้อมูลผู้ลงทะเบียนลง Google Drive
    saveRegistrationToDrive(activityTitle, regInfo, timestamp);

    return createJsonResponse({ 
      status: 'success', 
      message: 'Registered successfully to Sheets & Drive',
      activityTitle: activityTitle 
    });

  } catch (err) {
    return createJsonResponse({ status: 'error', message: err.toString() });
  }
}

/**
 * ==============================================================================
 * ฟังก์ชันบันทึกข้อมูลลง Google Sheets (แยกหมวดหมู่อย่างเป็นระเบียบ)
 * ==============================================================================
 */
function saveAllStateToSheets(state) {
  if (!state) return;
  var ss = getSpreadsheet();
  if (!ss) return;

  // 1. ชีทสรุปสถิติระบบ (Overview Dashboard Sheet)
  var summarySheet = getOrCreateSheet(ss, "สถิติระบบ");
  summarySheet.clear();
  createSheetHeader(summarySheet, ["รายการข้อมูล", "จำนวนในระบบ", "อัปเดตล่าสุด"]);
  var nowStr = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss");
  summarySheet.appendRow(["กิจกรรมทั้งหมด", (state.activities || []).length, nowStr]);
  summarySheet.appendRow(["คำร้องขอยืมอุปกรณ์", (state.borrowRequests || []).length, nowStr]);
  summarySheet.appendRow(["อุปกรณ์ในคลัง", (state.equipmentList || []).length, nowStr]);
  summarySheet.appendRow(["ข่าวประชาสัมพันธ์", (state.newsList || []).length, nowStr]);
  summarySheet.appendRow(["กำหนดการปฏิทิน", (state.calendarEvents || []).length, nowStr]);
  summarySheet.appendRow(["องค์กรและชมรม", (state.organizations || []).length, nowStr]);
  summarySheet.appendRow(["ผู้ดูแลระบบ (Admins)", (state.adminUsers || []).length, nowStr]);
  summarySheet.appendRow(["สาขาวิชา", (state.majors || []).length, nowStr]);

  // 2. ชีทรายการกิจกรรม (Activities)
  var actSheet = getOrCreateSheet(ss, "รายการกิจกรรม");
  actSheet.clear();
  createSheetHeader(actSheet, [
    "ID", "ชื่อกิจกรรม", "หมวดหมู่", "รูปภาพแบนเนอร์ URL", "วันที่จัดงาน", "สถานที่", 
    "กำหนดปิดรับ", "สถานะ", "ผู้สมัคร (คน)", "เก็บข้อมูลสุขภาพ", "รายละเอียดกิจกรรม", "ส่วนเพิ่มเติม (Custom Sections JSON)"
  ]);

  var customSecSheet = getOrCreateSheet(ss, "โครงสร้างฟอร์มเพิ่มเติม");
  customSecSheet.clear();
  createSheetHeader(customSecSheet, [
    "ID กิจกรรม", "ชื่อกิจกรรม", "ชื่อส่วนเพิ่มเติม (Section Title)", "ข้อคำถาม (Question Label)", "ประเภทข้อคำถาม (Type)", "ตัวเลือก (Options)", "จำเป็นต้องตอบ (Required)"
  ]);

  if (state.activities && state.activities.length > 0) {
    state.activities.forEach(function(a) {
      var customSecs = a.customSections || (a.customQuestions ? [{ id: 'sec_1', title: 'ส่วนเพิ่มเติม', questions: a.customQuestions }] : []);
      var customJsonStr = (customSecs && customSecs.length > 0) ? JSON.stringify(customSecs) : '';

      actSheet.appendRow([
        a.id, 
        a.title, 
        a.category, 
        a.image || '', 
        a.date, 
        a.location, 
        a.deadline, 
        a.status, 
        a.applicants || 0, 
        a.collectMedical ? 'เก็บ' : 'ไม่เก็บ', 
        a.description || '', 
        customJsonStr
      ]);

      if (customSecs && customSecs.length > 0) {
        customSecs.forEach(function(sec) {
          if (sec.questions && sec.questions.length > 0) {
            sec.questions.forEach(function(q) {
              var optsStr = Array.isArray(q.options) ? q.options.join(', ') : (q.options || '');
              customSecSheet.appendRow([
                a.id,
                a.title,
                sec.title || 'ส่วนเพิ่มเติม',
                q.label || '',
                q.type || 'text',
                optsStr,
                q.required ? 'ใช่' : 'ไม่'
              ]);
            });
          }
        });
      }
    });
  }

  // 3. ชีทรายการอุปกรณ์ในคลัง (Equipment Inventory)
  var eqSheet = getOrCreateSheet(ss, "รายการอุปกรณ์ในคลัง");
  eqSheet.clear();
  createSheetHeader(eqSheet, ["ID", "ชื่ออุปกรณ์", "หมวดหมู่", "จำนวนรวมคลัง", "รูปภาพ URL", "รายละเอียด"]);
  if (state.equipmentList && state.equipmentList.length > 0) {
    state.equipmentList.forEach(function(e) {
      eqSheet.appendRow([e.id, e.name, e.category, e.totalQty, e.image, e.description || '-']);
    });
  }

  // 4. ชีทคำร้องขอยืมอุปกรณ์ (Borrow Requests)
  var bSheet = getOrCreateSheet(ss, "คำร้องขอยืมอุปกรณ์");
  bSheet.clear();
  createSheetHeader(bSheet, ["รหัสคำร้อง", "วัน-เวลาที่ยื่น", "รหัสนักศึกษา", "ชื่อ-นามสกุล", "สาขาวิชา", "ชั้นปี", "เบอร์โทร/Line", "อุปกรณ์", "จำนวน", "วันที่ยืม", "กำหนดคืน", "โครงการ/สังกัด", "วัตถุประสงค์", "สถานะ"]);
  if (state.borrowRequests && state.borrowRequests.length > 0) {
    state.borrowRequests.forEach(function(r) {
      bSheet.appendRow([r.id, r.createdAt, r.studentId, r.studentName, r.major, 'ปี ' + r.year, r.phone, r.equipmentName, r.qty, r.borrowDate, r.returnDate, r.orgName, r.purpose || '-', r.status]);
    });
  }

  // 5. ชีทข่าวประชาสัมพันธ์ (News)
  var nSheet = getOrCreateSheet(ss, "ข่าวประชาสัมพันธ์");
  nSheet.clear();
  createSheetHeader(nSheet, ["ID", "หัวข้อข่าวสาร", "หมวดหมู่", "สรุปข่าว", "เนื้อหาฉบับเต็ม", "รูปภาพ URL"]);
  if (state.newsList && state.newsList.length > 0) {
    state.newsList.forEach(function(n) {
      var imgStr = (Array.isArray(n.images) && n.images.length > 0) ? n.images.join(', ') : (n.image || '');
      nSheet.appendRow([n.id, n.title, n.category, n.summary, n.content || '-', imgStr]);
    });
  }

  // 6. ชีทกำหนดการปฏิทิน (Calendar Events)
  var cSheet = getOrCreateSheet(ss, "กำหนดการปฏิทิน");
  cSheet.clear();
  createSheetHeader(cSheet, ["ID", "วันที่", "เดือน", "ชื่อเดือน", "หัวข้อกิจกรรม", "เวลาจัดงาน", "สถานที่"]);
  if (state.calendarEvents && state.calendarEvents.length > 0) {
    state.calendarEvents.forEach(function(c) {
      cSheet.appendRow([c.id, c.date, c.month, c.monthName, c.title, c.time, c.location]);
    });
  }

  // 7. ชีทโครงสร้างองค์กรและสมาชิก (Organizations & Members)
  var oSheet = getOrCreateSheet(ss, "โครงสร้างองค์กร");
  oSheet.clear();
  createSheetHeader(oSheet, ["ชื่อองค์กร/ชมรม", "ชื่อ-นามสกุลสมาชิก", "ชื่อเล่น", "ตำแหน่ง", "ช่องทางติดต่อ", "รูปภาพ URL"]);
  if (state.organizations && state.organizations.length > 0) {
    state.organizations.forEach(function(org) {
      if (org.members && org.members.length > 0) {
        org.members.forEach(function(m) {
          oSheet.appendRow([org.name, m.fullname, m.nickname, m.position, m.contact || '-', m.image]);
        });
      } else {
        oSheet.appendRow([org.name, 'ยังไม่มีสมาชิก', '-', '-', '-', '-']);
      }
    });
  }

  // 8. ชีทผู้ดูแลระบบ (Admin Users)
  var uSheet = getOrCreateSheet(ss, "ผู้ดูแลระบบ");
  uSheet.clear();
  createSheetHeader(uSheet, ["ID", "ชื่อ-นามสกุล", "Username", "Password", "บทบาท (Role)"]);
  if (state.adminUsers && state.adminUsers.length > 0) {
    state.adminUsers.forEach(function(u) {
      uSheet.appendRow([u.id, u.fullname, u.username, u.password || 'kku123', u.role]);
    });
  }

  // 9. ชีทสาขาวิชา (Majors)
  var mSheet = getOrCreateSheet(ss, "สาขาวิชา");
  mSheet.clear();
  createSheetHeader(mSheet, ["ลำดับ", "ชื่อสาขาวิชา"]);
  if (state.majors && state.majors.length > 0) {
    state.majors.forEach(function(m, idx) {
      mSheet.appendRow([idx + 1, m]);
    });
  }

  // 10. ชีทสรุปผู้เข้าร่วมแต่ละกิจกรรม
  updateActivitySummarySheet(ss);
}

/**
 * ==============================================================================
 * ฟังก์ชันสร้างและอัปเดตชีทสรุปผู้เข้าร่วมแต่ละกิจกรรม (อ่านจากชีท "ข้อมูลการลงทะเบียนกิจกรรม")
 * ==============================================================================
 */
function updateActivitySummarySheet(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();

  var summarySheetName = "สรุปผู้เข้าร่วมกิจกรรม";
  var summarySheet = getOrCreateSheet(ss, summarySheetName);
  summarySheet.clear();

  var headers = [
    "ชื่อกิจกรรม", "สาขาวิชา", "ชั้นปี 1 (คน)", "ชั้นปี 2 (คน)", "ชั้นปี 3 (คน)", "ชั้นปี 4 (คน)", "อื่นๆ (คน)", "รวมตามสาขา (คน)", "ยอดรวมทั้งกิจกรรม (คน)"
  ];
  createSheetHeader(summarySheet, headers);

  var regSheet = ss.getSheetByName("ข้อมูลการลงทะเบียนกิจกรรม");
  if (!regSheet) return;

  var lastRow = regSheet.getLastRow();
  if (lastRow <= 1) return;

  var data = regSheet.getRange(2, 1, lastRow - 1, 9).getValues();

  var actMap = {};
  var actTotals = {};

  data.forEach(function(row) {
    var actTitle = (row[0] || 'กิจกรรมทั่วไป').toString().trim();
    var major = (row[4] || 'ไม่ระบุสาขา').toString().trim();
    var yearStr = (row[5] || '').toString().trim();

    if (!actMap[actTitle]) actMap[actTitle] = {};
    if (!actTotals[actTitle]) actTotals[actTitle] = 0;

    if (!actMap[actTitle][major]) {
      actMap[actTitle][major] = { y1: 0, y2: 0, y3: 0, y4: 0, other: 0, total: 0 };
    }

    if (yearStr.indexOf('1') !== -1) {
      actMap[actTitle][major].y1++;
    } else if (yearStr.indexOf('2') !== -1) {
      actMap[actTitle][major].y2++;
    } else if (yearStr.indexOf('3') !== -1) {
      actMap[actTitle][major].y3++;
    } else if (yearStr.indexOf('4') !== -1) {
      actMap[actTitle][major].y4++;
    } else {
      actMap[actTitle][major].other++;
    }
    actMap[actTitle][major].total++;
    actTotals[actTitle]++;
  });

  var actTitles = Object.keys(actMap);
  actTitles.forEach(function(aTitle) {
    var majorMap = actMap[aTitle];
    var majorList = Object.keys(majorMap);
    var grandTotal = actTotals[aTitle];

    majorList.forEach(function(mKey, idx) {
      var stat = majorMap[mKey];
      var isFirstMajorRow = (idx === 0);

      summarySheet.appendRow([
        isFirstMajorRow ? aTitle : '',
        mKey,
        stat.y1,
        stat.y2,
        stat.y3,
        stat.y4,
        stat.other,
        stat.total,
        isFirstMajorRow ? grandTotal : ''
      ]);
    });

    summarySheet.appendRow(['', '', '', '', '', '', '', '', '']);
  });

  try {
    var lRow = summarySheet.getLastRow();
    if (lRow > 1) {
      summarySheet.getRange(2, 3, lRow - 1, 7).setHorizontalAlignment("center");
      summarySheet.getRange(2, 2, lRow - 1, 1).setFontWeight("bold");
    }
  } catch(e) {}
}

/**
 * ==============================================================================
 * ฟังก์ชันสำรองข้อมูลลง Google Drive (Google Drive Complete Backup)
 * ==============================================================================
 */
function getOrCreateBackupFolder() {
  var folderName = "KKU_FIS_StudentUnion_Backups";
  var folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return DriveApp.createFolder(folderName);
  }
}

function saveStateToProperties(state) {
  if (!state) return;
  var scriptProperties = PropertiesService.getScriptProperties();
  try {
    var jsonStr = JSON.stringify(state);
    var chunkSize = 8000;
    var totalChunks = Math.ceil(jsonStr.length / chunkSize);
    scriptProperties.setProperty('KKU_STATE_TOTAL_CHUNKS', totalChunks.toString());

    for (var i = 0; i < totalChunks; i++) {
      var chunk = jsonStr.substring(i * chunkSize, (i + 1) * chunkSize);
      scriptProperties.setProperty('KKU_STATE_CHUNK_' + i, chunk);
    }

    var oldChunks = scriptProperties.getProperty('KKU_STATE_PREV_CHUNKS');
    if (oldChunks) {
      var prevCount = parseInt(oldChunks, 10);
      for (var k = totalChunks; k < prevCount; k++) {
        scriptProperties.deleteProperty('KKU_STATE_CHUNK_' + k);
      }
    }
    scriptProperties.setProperty('KKU_STATE_PREV_CHUNKS', totalChunks.toString());
  } catch (err) {
    Logger.log("Error saving state to properties: " + err.toString());
  }
}

function getStateFromProperties() {
  var scriptProperties = PropertiesService.getScriptProperties();
  try {
    var totalChunksStr = scriptProperties.getProperty('KKU_STATE_TOTAL_CHUNKS');
    if (!totalChunksStr) {
      var raw = scriptProperties.getProperty('KKU_APP_STATE');
      if (raw) return JSON.parse(raw);
      return null;
    }

    var totalChunks = parseInt(totalChunksStr, 10);
    var fullJson = '';
    for (var i = 0; i < totalChunks; i++) {
      var chunk = scriptProperties.getProperty('KKU_STATE_CHUNK_' + i);
      if (chunk) {
        fullJson += chunk;
      }
    }

    if (fullJson) {
      return JSON.parse(fullJson);
    }
  } catch (err) {
    Logger.log("Error reading state from properties: " + err.toString());
  }
  return null;
}

function getStateFromDriveBackup() {
  try {
    var folder = getOrCreateBackupFolder();
    var files = folder.getFilesByName("latest_system_state.json");
    if (files.hasNext()) {
      var file = files.next();
      var content = file.getContent();
      if (content && content.trim() !== "") {
        return JSON.parse(content);
      }
    }
  } catch (err) {
    Logger.log("ไม่สามารถอ่านไฟล์สำรองจาก Google Drive ได้: " + err.toString());
  }
  return null;
}

function getStateFromSheets(ss) {
  if (!ss) ss = getSpreadsheet();
  if (!ss) return null;

  try {
    var state = {
      activities: [],
      newsList: [],
      calendarEvents: [],
      adminUsers: [],
      organizations: [],
      equipmentList: [],
      borrowRequests: [],
      majors: []
    };

    // 1. อ่านรายการกิจกรรม
    var actSheet = ss.getSheetByName("รายการกิจกรรม");
    if (actSheet && actSheet.getLastRow() > 1) {
      var actNumCols = actSheet.getLastColumn();
      var actVals = actSheet.getRange(2, 1, actSheet.getLastRow() - 1, Math.max(actNumCols, 12)).getValues();

      // อ่านชีทโครงสร้างฟอร์มเพิ่มเติม (ถ้ามี)
      var customSecMap = {};
      var csSheet = ss.getSheetByName("โครงสร้างฟอร์มเพิ่มเติม");
      if (csSheet && csSheet.getLastRow() > 1) {
        var csVals = csSheet.getRange(2, 1, csSheet.getLastRow() - 1, 7).getValues();
        csVals.forEach(function(csRow) {
          var actIdKey = csRow[0] ? csRow[0].toString() : '';
          var actTitleKey = csRow[1] ? csRow[1].toString().trim() : '';
          var secTitle = csRow[2] ? csRow[2].toString().trim() : 'ส่วนเพิ่มเติม';
          var qLabel = csRow[3] ? csRow[3].toString().trim() : '';
          var qType = csRow[4] ? csRow[4].toString().trim() : 'text';
          var qOptsRaw = csRow[5] ? csRow[5].toString() : '';
          var qReq = (csRow[6] === 'ใช่' || csRow[6] === true || csRow[6] === 'true');

          if (qLabel && (actIdKey || actTitleKey)) {
            var key = actIdKey || actTitleKey;
            if (!customSecMap[key]) customSecMap[key] = {};
            if (!customSecMap[key][secTitle]) customSecMap[key][secTitle] = [];

            var optsArr = qOptsRaw ? qOptsRaw.split(',').map(function(s){ return s.trim(); }) : [];
            customSecMap[key][secTitle].push({
              id: 'q_' + Date.now() + Math.floor(Math.random() * 1000),
              label: qLabel,
              type: qType,
              options: optsArr,
              required: qReq
            });
          }
        });
      }

      actVals.forEach(function(row) {
        if (row[1]) {
          var actId = row[0] || Date.now();
          var actTitle = row[1];
          var actImg = '';
          var actCat = 'กิจกรรม';
          var actDate = '';
          var actLoc = '';
          var actDeadline = '';
          var actStatus = 'เปิดรับสมัคร';
          var actApplicants = 0;
          var actCollectMed = false;
          var actDesc = '';
          var parsedCustomSections = [];

          if (actNumCols >= 12) {
            // โครงสร้างใหม่ 12 คอลัมน์
            actCat = row[2] || 'กิจกรรม';
            actImg = row[3] || '';
            actDate = row[4] || '';
            actLoc = row[5] || '';
            actDeadline = row[6] || '';
            actStatus = row[7] || 'เปิดรับสมัคร';
            actApplicants = parseInt(row[8]) || 0;
            actCollectMed = (row[9] === 'เก็บ' || row[9] === true);
            actDesc = row[10] || '';
            if (row[11]) {
              try {
                parsedCustomSections = JSON.parse(row[11]);
              } catch(e) {}
            }
          } else {
            // โครงสร้างเดิม 9 คอลัมน์ (Backward compatibility)
            actCat = row[2] || 'กิจกรรม';
            actDate = row[3] || '';
            actLoc = row[4] || '';
            actDeadline = row[5] || '';
            actStatus = row[6] || 'เปิดรับสมัคร';
            actApplicants = parseInt(row[7]) || 0;
            actCollectMed = (row[8] === 'เก็บ' || row[8] === true);
          }

          // ถ้าไม่มี customSections จาก JSON แต่มีข้อมูลในชีทโครงสร้างฟอร์มเพิ่มเติม ให้ประกอบสร้างกลับมา
          if ((!parsedCustomSections || parsedCustomSections.length === 0) && (customSecMap[actId.toString()] || customSecMap[actTitle])) {
            var secObj = customSecMap[actId.toString()] || customSecMap[actTitle];
            parsedCustomSections = [];
            Object.keys(secObj).forEach(function(sTitle) {
              parsedCustomSections.push({
                id: 'sec_' + Date.now() + Math.floor(Math.random() * 1000),
                title: sTitle,
                questions: secObj[sTitle]
              });
            });
          }

          state.activities.push({
            id: actId,
            title: actTitle,
            category: actCat,
            image: actImg,
            date: actDate,
            location: actLoc,
            deadline: actDeadline,
            status: actStatus,
            applicants: actApplicants,
            collectMedical: actCollectMed,
            description: actDesc,
            customSections: parsedCustomSections
          });
        }
      });
    }

    // 2. อ่านข่าวประชาสัมพันธ์
    var nSheet = ss.getSheetByName("ข่าวประชาสัมพันธ์");
    if (nSheet && nSheet.getLastRow() > 1) {
      var nVals = nSheet.getRange(2, 1, nSheet.getLastRow() - 1, 6).getValues();
      nVals.forEach(function(row) {
        if (row[1]) {
          var imgStr = (row[5] || '').toString();
          var imgs = imgStr ? imgStr.split(',').map(function(s){ return s.trim(); }) : [];
          state.newsList.push({
            id: row[0] || Date.now(),
            title: row[1],
            category: row[2] || 'ข่าวสาร',
            summary: row[3] || '',
            content: row[4] || '',
            image: imgs[0] || '',
            images: imgs
          });
        }
      });
    }

    // 3. อ่านกำหนดการปฏิทิน
    var cSheet = ss.getSheetByName("กำหนดการปฏิทิน");
    if (cSheet && cSheet.getLastRow() > 1) {
      var cVals = cSheet.getRange(2, 1, cSheet.getLastRow() - 1, 7).getValues();
      cVals.forEach(function(row) {
        if (row[4]) {
          state.calendarEvents.push({
            id: row[0] || Date.now(),
            date: row[1] || '',
            month: row[2] || '',
            monthName: row[3] || '',
            title: row[4],
            time: row[5] || '',
            location: row[6] || ''
          });
        }
      });
    }

    // 4. อ่านรายการอุปกรณ์ในคลัง
    var eqSheet = ss.getSheetByName("รายการอุปกรณ์ในคลัง");
    if (eqSheet && eqSheet.getLastRow() > 1) {
      var eqVals = eqSheet.getRange(2, 1, eqSheet.getLastRow() - 1, 6).getValues();
      eqVals.forEach(function(row) {
        if (row[1]) {
          state.equipmentList.push({
            id: row[0] || Date.now(),
            name: row[1],
            category: row[2] || 'ทั่วไป',
            totalQty: parseInt(row[3]) || 1,
            image: row[4] || '',
            description: row[5] || ''
          });
        }
      });
    }

    // 5. อ่านคำร้องขอยืมอุปกรณ์
    var bSheet = ss.getSheetByName("คำร้องขอยืมอุปกรณ์");
    if (bSheet && bSheet.getLastRow() > 1) {
      var bVals = bSheet.getRange(2, 1, bSheet.getLastRow() - 1, 14).getValues();
      bVals.forEach(function(row) {
        if (row[0]) {
          var yStr = (row[5] || '').toString().replace(/[^0-9]/g, '') || '1';
          state.borrowRequests.push({
            id: row[0],
            createdAt: row[1] || '',
            studentId: row[2] || '',
            studentName: row[3] || '',
            major: row[4] || '',
            year: yStr,
            phone: row[6] || '',
            equipmentName: row[7] || '',
            qty: parseInt(row[8]) || 1,
            borrowDate: row[9] || '',
            returnDate: row[10] || '',
            orgName: row[11] || '',
            purpose: row[12] || '',
            status: row[13] || 'รออนุมัติ'
          });
        }
      });
    }

    // 6. อ่านสาขาวิชา
    var mSheet = ss.getSheetByName("สาขาวิชา");
    if (mSheet && mSheet.getLastRow() > 1) {
      var mVals = mSheet.getRange(2, 1, mSheet.getLastRow() - 1, 2).getValues();
      mVals.forEach(function(row) {
        if (row[1]) {
          state.majors.push(row[1].toString().trim());
        }
      });
    }

    // 7. อ่านผู้ดูแลระบบ
    var uSheet = ss.getSheetByName("ผู้ดูแลระบบ");
    if (uSheet && uSheet.getLastRow() > 1) {
      var numCols = uSheet.getLastColumn();
      var uVals = uSheet.getRange(2, 1, uSheet.getLastRow() - 1, Math.max(numCols, 5)).getValues();
      uVals.forEach(function(row) {
        var uName = (row[2] || '').toString().trim();
        if (uName) {
          var uPass = 'kku123';
          var uRole = 'Admin';

          if (numCols >= 5 && row[3] && row[4]) {
            // โครงสร้างใหม่ 5 คอลัมน์ (ID, Name, Username, Password, Role)
            uPass = row[3].toString().trim();
            uRole = row[4].toString().trim();
          } else {
            // โครงสร้างเก่า 4 คอลัมน์ (ID, Name, Username, Role) -> ให้รหัสผ่านเริ่มต้นเป็น kku123
            uRole = (row[3] || 'Admin').toString().trim();
            uPass = 'kku123';
          }

          state.adminUsers.push({
            id: row[0] || Date.now(),
            fullname: row[1] || '',
            username: uName,
            password: uPass,
            role: uRole
          });
        }
      });
    }

    // 8. อ่านโครงสร้างองค์กรและสมาชิก
    var oSheet = ss.getSheetByName("โครงสร้างองค์กร");
    if (oSheet && oSheet.getLastRow() > 1) {
      var oVals = oSheet.getRange(2, 1, oSheet.getLastRow() - 1, 6).getValues();
      var orgMap = {};
      oVals.forEach(function(row) {
        var orgName = (row[0] || '').toString().trim();
        if (orgName) {
          if (!orgMap[orgName]) {
            orgMap[orgName] = { id: Date.now() + Math.floor(Math.random() * 1000), name: orgName, members: [] };
          }
          var memberName = (row[1] || '').toString().trim();
          if (memberName && memberName !== 'ยังไม่มีสมาชิก') {
            orgMap[orgName].members.push({
              id: Date.now() + Math.floor(Math.random() * 1000),
              fullname: memberName,
              nickname: row[2] || '',
              position: row[3] || 'สมาชิก',
              contact: row[4] || '',
              image: row[5] || ''
            });
          }
        }
      });
      Object.keys(orgMap).forEach(function(k) {
        state.organizations.push(orgMap[k]);
      });
    }

    if (!state.adminUsers || state.adminUsers.length === 0) {
      state.adminUsers = [{ id: 1, fullname: "ผู้ดูแลระบบสโมสรฯ", username: "admin", password: "kku123", role: "Super Admin" }];
    }

    if (state.activities.length > 0 || state.newsList.length > 0 || state.equipmentList.length > 0) {
      return state;
    }
  } catch (err) {
    Logger.log("Error reading state from sheets: " + err.toString());
  }
  return null;
}

function saveDriveBackup(state) {
  try {
    var folder = getOrCreateBackupFolder();
    var jsonContent = JSON.stringify(state, null, 2);
    var timestampStr = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd_HHmmss");

    // 1. อัปเดตไฟล์โครงสร้างระบบล่าสุด (latest_system_state.json)
    var latestFiles = folder.getFilesByName("latest_system_state.json");
    if (latestFiles.hasNext()) {
      var latestFile = latestFiles.next();
      latestFile.setContent(jsonContent);
    } else {
      folder.createFile("latest_system_state.json", jsonContent, MimeType.PLAIN_TEXT);
    }

    // 2. สร้างไฟล์สำรองประวัติตามวันเวลา (backup_state_YYYY-MM-DD_HHmmss.json)
    folder.createFile("backup_state_" + timestampStr + ".json", jsonContent, MimeType.PLAIN_TEXT);

  } catch (err) {
    Logger.log("ไม่สามารถบันทึกไฟล์ลง Google Drive ได้: " + err.toString());
  }
}

function saveRegistrationToDrive(activityTitle, regInfo, timestamp) {
  try {
    var folder = getOrCreateBackupFolder();
    var fileName = "all_activity_registrations.json";

    var existingData = [];
    var files = folder.getFilesByName(fileName);
    var file;

    if (files.hasNext()) {
      file = files.next();
      try {
        existingData = JSON.parse(file.getContent());
      } catch (e) {
        existingData = [];
      }
    } else {
      file = folder.createFile(fileName, "[]", MimeType.PLAIN_TEXT);
    }

    var newRecord = {
      activityTitle: activityTitle,
      timestamp: timestamp,
      studentId: regInfo.studentId || regInfo.student_id || '-',
      name: regInfo.name || '-',
      major: regInfo.major || '-',
      year: regInfo.year ? 'ปี ' + regInfo.year : '-',
      phone: regInfo.phone || '-',
      medical: regInfo.medical || '-',
      customAnswers: regInfo.customAnswers || '-'
    };

    existingData.push(newRecord);
    file.setContent(JSON.stringify(existingData, null, 2));

  } catch (err) {
    Logger.log("ไม่สามารถสำรองข้อมูลผู้ลงทะเบียนลง Google Drive ได้: " + err.toString());
  }
}

function saveBorrowRequestToDrive(req) {
  try {
    var folder = getOrCreateBackupFolder();
    var fileName = "all_borrow_requests.json";

    var existingData = [];
    var files = folder.getFilesByName(fileName);
    var file;

    if (files.hasNext()) {
      file = files.next();
      try {
        existingData = JSON.parse(file.getContent());
      } catch (e) {
        existingData = [];
      }
    } else {
      file = folder.createFile(fileName, "[]", MimeType.PLAIN_TEXT);
    }

    existingData.push(req);
    file.setContent(JSON.stringify(existingData, null, 2));

  } catch (err) {
    Logger.log("ไม่สามารถสำรองคำร้องขอยืมอุปกรณ์ลง Google Drive ได้: " + err.toString());
  }
}

/**
 * ==============================================================================
 * ฟังก์ชันผู้ช่วยสร้าง/ดึงแผ่นชีท และตกแต่ง Header สวยงาม
 * ==============================================================================
 */
function getSpreadsheet() {
  var ss = null;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {}

  if (!ss) {
    var targetId = "13PkrBncnM9jSBSifhp4Q0lyvULqPxGr0kxmumHWUKdA";
    try {
      ss = SpreadsheetApp.openById(targetId);
    } catch (e) {
      Logger.log("ไม่สามารถเปิด Google Sheet ด้วย ID ได้: " + e.toString());
    }
  }
  return ss;
}

function getOrCreateSheet(ss, name) {
  if (!ss) ss = getSpreadsheet();
  if (!ss) return null;
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function createSheetHeader(sheet, headers) {
  var headerRow = sheet.getRange(1, 1, 1, headers.length);
  headerRow.setValues([headers]);
  headerRow.setFontWeight("bold");
  headerRow.setBackground("#8A1B13");
  headerRow.setFontColor("#FFFFFF");
  sheet.setFrozenRows(1);
}

function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * ==============================================================================
 * ฟังก์ชันสร้างและอัปเดตชีท "ฐานข้อมูลนักศึกษา" (Student Directory Sheet)
 * ==============================================================================
 */
function updateOrAddStudentProfile(ss, profile) {
  if (!profile || !profile.studentId) return;
  var sId = (profile.studentId || '').toString().trim();
  var cleanSearchId = sId.replace(/[^0-9]/g, '');
  if (!cleanSearchId) return;

  var dbSheet = getOrCreateSheet(ss, "ฐานข้อมูลนักศึกษา");
  if (dbSheet.getLastRow() === 0) {
    var headers = ["รหัสนักศึกษา", "ชื่อ-นามสกุล", "สาขาวิชา", "ชั้นปี", "เบอร์โทรศัพท์", "โรคประจำตัว/ข้อจำกัดอาหาร", "อัปเดตล่าสุด"];
    createSheetHeader(dbSheet, headers);
  }

  var lastRow = dbSheet.getLastRow();
  var timestamp = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss");
  var yearFormatted = profile.year ? (profile.year.toString().indexOf('ปี') !== -1 ? profile.year : 'ปี ' + profile.year) : '-';
  var newRowData = [
    sId,
    profile.name || profile.studentName || '-',
    profile.major || '-',
    yearFormatted,
    profile.phone || '-',
    profile.medical || '-',
    timestamp
  ];

  var foundRowIndex = -1;
  if (lastRow > 1) {
    var values = dbSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < values.length; i++) {
      var rowId = (values[i][0] || '').toString().replace(/[^0-9]/g, '');
      if (rowId === cleanSearchId) {
        foundRowIndex = i + 2;
        break;
      }
    }
  }

  if (foundRowIndex > 1) {
    dbSheet.getRange(foundRowIndex, 1, 1, newRowData.length).setValues([newRowData]);
  } else {
    dbSheet.appendRow(newRowData);
  }
}
