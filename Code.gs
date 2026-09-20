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

    // 1.2 หน้าเว็บขอข้อมูลงบกลาง/สถานะระบบ (getState)
    if (action === 'getState') {
      var scriptProperties = PropertiesService.getScriptProperties();
      var rawState = scriptProperties.getProperty('KKU_APP_STATE');
      var state = null;
      if (rawState) {
        try {
          state = JSON.parse(rawState);
        } catch (e) {}
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
      var scriptProperties = PropertiesService.getScriptProperties();
      try {
        var jsonStr = JSON.stringify(state);
        if (jsonStr.length < 9000) {
          scriptProperties.setProperty('KKU_APP_STATE', jsonStr);
        } else {
          scriptProperties.deleteProperty('KKU_APP_STATE');
        }
      } catch (e) {
        try { scriptProperties.deleteProperty('KKU_APP_STATE'); } catch (err) {}
      }

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
  createSheetHeader(actSheet, ["ID", "ชื่อกิจกรรม", "หมวดหมู่", "วันที่จัดงาน", "สถานที่", "กำหนดปิดรับ", "สถานะ", "ผู้สมัคร (คน)", "เก็บข้อมูลสุขภาพ"]);
  if (state.activities && state.activities.length > 0) {
    state.activities.forEach(function(a) {
      actSheet.appendRow([a.id, a.title, a.category, a.date, a.location, a.deadline, a.status, a.applicants || 0, a.collectMedical ? 'เก็บ' : 'ไม่เก็บ']);
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
  createSheetHeader(uSheet, ["ID", "ชื่อ-นามสกุล", "Username", "บทบาท (Role)"]);
  if (state.adminUsers && state.adminUsers.length > 0) {
    state.adminUsers.forEach(function(u) {
      uSheet.appendRow([u.id, u.fullname, u.username, u.role]);
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
