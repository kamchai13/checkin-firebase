const express = require('express');
const cors = require('cors');
const ExcelJS = require('exceljs');
const path = require('path');
const admin = require('firebase-admin');
const fs = require('fs');

const app = express();
app.use(cors());
app.use(express.json());
app.use(express.static(__dirname));

// ----------------------------------------------------
// เชื่อมต่อ Firebase Firestore (รองรับทั้ง Local และ Render)
// ----------------------------------------------------
let serviceAccount = null;

try {
  if (fs.existsSync('./serviceAccountKey.json')) {
    serviceAccount = require('./serviceAccountKey.json');
  } else if (fs.existsSync('./serviceAccountKey.json.json')) {
    serviceAccount = require('./serviceAccountKey.json.json');
  } else if (process.env.FIREBASE_CONFIG) {
    const decodedConfig = Buffer.from(process.env.FIREBASE_CONFIG, 'base64').toString('utf-8');
    serviceAccount = JSON.parse(decodedConfig);
  } else {
    console.error("❌ Local Firebase key file not found!");
  }
} catch (err) {
  console.error("❌ Error loading serviceAccountKey:", err.message);
}

// ตรวจสอบการ Initialize
let db = null;
let studentsCol = null;
let attendanceCol = null;

if (serviceAccount) {
  try {
    admin.initializeApp({
      credential: admin.credential.cert(serviceAccount)
    });
    db = admin.firestore();
    studentsCol = db.collection('students');
    attendanceCol = db.collection('attendance');
    console.log("🔥 Firebase Firestore connected successfully!");
  } catch (err) {
    console.error("❌ Firebase initialization failed:", err.message);
  }
} else {
  console.error("⚠️ CRITICAL: Firebase serviceAccount is undefined. Firestore operations will fail!");
}

function calculateGrade(total) {
    if (total >= 80) return 'A';
    if (total >= 75) return 'B+';
    if (total >= 70) return 'B';
    if (total >= 65) return 'C+';
    if (total >= 60) return 'C';
    if (total >= 55) return 'D+';
    if (total >= 50) return 'D';
    return 'F';
}

// Middleware ตรวจสอบการเชื่อมต่อ Firebase ก่อนรับ API
const checkFirebaseConnection = (req, res, next) => {
  if (!db) {
    return res.status(500).json({ error: "Firebase DB is not initialized. Please check serviceAccountKey.json or FIREBASE_CONFIG." });
  }
  next();
};

// ----------------------------------------------------
// Routes หน้าเว็บ
// ----------------------------------------------------
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

app.get('/checkin.html', (req, res) => {
    res.sendFile(path.join(__dirname, 'checkin.html'));
});

// ----------------------------------------------------
// API ระบบเช็คชื่อ (ค้นหารหัสนักศึกษาแบบยืดหยุ่น)
// ----------------------------------------------------
app.post('/api/checkin', checkFirebaseConnection, async (req, res) => {
    try {
        const { courseKey, studentId, classStartTime } = req.body;
        
        if (!studentId) {
            return res.status(400).json({ message: "กรุณากรอกหรือสแกนรหัสนักศึกษา" });
        }

        const cleanStudentId = studentId.toString().trim();
        const cleanCourseKey = courseKey ? courseKey.toString().trim() : '';

        let studentData = null;

        // 1. ค้นหาจาก Document ID ก่อน
        const docRef = studentsCol.doc(cleanStudentId);
        const docSnap = await docRef.get();

        if (docSnap.exists) {
            studentData = docSnap.data();
        } else {
            // 2. ถ้าไม่เจอ ค้นจาก field student_id (String)
            let querySnap = await studentsCol.where('student_id', '==', cleanStudentId).get();
            if (!querySnap.empty) {
                studentData = querySnap.docs[0].data();
            } else {
                // 3. เผื่อใน DB เก็บรหัสเป็น Number
                const numStudentId = Number(cleanStudentId);
                if (!isNaN(numStudentId)) {
                    querySnap = await studentsCol.where('student_id', '==', numStudentId).get();
                    if (!querySnap.empty) {
                        studentData = querySnap.docs[0].data();
                    }
                }
            }
        }

        // หากยังไม่พบข้อมูลในระบบ
        if (!studentData) {
            return res.status(404).json({ message: `❌ ไม่พบรหัสนักศึกษา ${cleanStudentId} ในระบบ` });
        }

        const subjectToUse = cleanCourseKey || studentData.subject || 'GENERAL';

        // คำนวณเวลาปัจจุบันและตรวจจับการเข้าเรียนสาย (เกิน 20 นาที)
        const now = new Date();
        const checkinTimeStr = now.toLocaleTimeString('th-TH', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
        let status = 'มาเรียน';

        if (classStartTime) {
            const [startHour, startMinute] = classStartTime.split(':').map(Number);
            const lateThreshold = new Date();
            lateThreshold.setHours(startHour, startMinute + 20, 0, 0); // เกิน 20 นาที = สาย

            if (now > lateThreshold) {
                status = 'สาย';
            }
        }

        // บันทึกเวลาเข้าเรียนลง Firestore
        await attendanceCol.add({
            student_id: cleanStudentId,
            student_name: studentData.name || '',
            subject: subjectToUse,
            status: status,
            checkin_time: checkinTimeStr,
            timestamp: admin.firestore.FieldValue.serverTimestamp()
        });

        const statusText = status === 'สาย' ? '⚠️ (สาย)' : '✅ (มาเรียน)';
        res.json({ 
            success: true,
            message: `เช็คชื่อสำเร็จ! ${studentData.name || cleanStudentId} ${statusText}`,
            studentId: cleanStudentId,
            studentName: studentData.name || cleanStudentId,
            checkinTime: checkinTimeStr,
            status: status
        });
    } catch (err) {
        console.error("Checkin Error:", err);
        res.status(500).json({ message: "เกิดข้อผิดพลาดทางเซิร์ฟเวอร์" });
    }
});

// API ดึงรายชื่อนักศึกษา
app.get('/api/students/:subject', checkFirebaseConnection, async (req, res) => {
    try {
        const { subject } = req.params;

        let query = studentsCol;
        if (subject !== 'ALL') {
            query = studentsCol.where('subject', '==', subject);
        }
        const studentsSnap = await query.get();

        const attendanceSnap = await attendanceCol.get();
        const attendanceCounts = {};

        attendanceSnap.forEach(doc => {
            const data = doc.data();
            const sid = data.student_id;
            if (sid) {
                attendanceCounts[sid] = (attendanceCounts[sid] || 0) + 1;
            }
        });

        const data = [];
        studentsSnap.forEach(doc => {
            const row = doc.data();
            const attCount = attendanceCounts[row.student_id] || 0;
            const attScore = Math.min(attCount, 10);
            const total = attScore + (row.score_assignment || 0) + (row.score_midterm || 0) + (row.score_final || 0);

            data.push({
                student_id: row.student_id,
                name: row.name,
                subject: row.subject,
                score_assignment: row.score_assignment || 0,
                score_midterm: row.score_midterm || 0,
                score_final: row.score_final || 0,
                score_attendance: attScore,
                total_score: total,
                grade: calculateGrade(total)
            });
        });

        res.json(data);
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// API อัปเดตคะแนน
app.post('/api/scores/update', checkFirebaseConnection, async (req, res) => {
    try {
        const { student_id, score_assignment, score_midterm, score_final } = req.body;
        const cleanId = student_id.toString().trim();

        await studentsCol.doc(cleanId).set({
            score_assignment: Number(score_assignment) || 0,
            score_midterm: Number(score_midterm) || 0,
            score_final: Number(score_final) || 0
        }, { merge: true });

        res.json({ message: "อัปเดตคะแนนเรียบร้อยแล้ว" });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// API สแกนฝั่งอาจารย์
app.post('/api/scan', checkFirebaseConnection, async (req, res) => {
    try {
        const { student_id, subject } = req.body;
        if (!student_id) return res.status(400).json({ error: "ข้อมูลไม่ครบถ้วน" });

        const cleanId = student_id.toString().trim();

        await attendanceCol.add({
            student_id: cleanId,
            subject: subject || 'GENERAL',
            status: 'มาเรียน',
            timestamp: admin.firestore.FieldValue.serverTimestamp()
        });

        res.json({ message: `เช็คชื่อสำเร็จ! รหัส ${cleanId}` });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// API เพิ่มนักศึกษา (ตัดช่องว่างอัตโนมัติ)
app.post('/api/students', checkFirebaseConnection, async (req, res) => {
    try {
        const { student_id, name, subject } = req.body;
        if (!student_id) return res.status(400).json({ error: "ต้องระบุรหัสนักศึกษา" });

        const cleanId = student_id.toString().trim();
        const cleanName = name ? name.toString().trim() : '';
        const cleanSubject = subject ? subject.toString().trim() : 'GENERAL';

        await studentsCol.doc(cleanId).set({
            student_id: cleanId,
            name: cleanName,
            subject: cleanSubject
        }, { merge: true });

        res.json({ message: "บันทึก/แก้ไข ข้อมูลนักศึกษาสำเร็จ" });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// API ส่งออก Excel
app.get('/api/export/:subject', checkFirebaseConnection, async (req, res) => {
    try {
        const { subject } = req.params;

        let query = studentsCol;
        if (subject !== 'ALL') {
            query = studentsCol.where('subject', '==', subject);
        }
        const studentsSnap = await query.get();
        const attendanceSnap = await attendanceCol.get();

        const attendanceCounts = {};
        attendanceSnap.forEach(doc => {
            const data = doc.data();
            const sid = data.student_id;
            if (sid) {
                attendanceCounts[sid] = (attendanceCounts[sid] || 0) + 1;
            }
        });

        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet('สมุดบันทึกคะแนน');

        sheet.columns = [
            { header: 'รหัสนักศึกษา', key: 'student_id', width: 15 },
            { header: 'ชื่อ-นามสกุล', key: 'name', width: 25 },
            { header: 'เข้าเรียน (10)', key: 'att', width: 12 },
            { header: 'งาน (50)', key: 'assign', width: 12 },
            { header: 'กลางภาค (20)', key: 'mid', width: 12 },
            { header: 'ปลายภาค (20)', key: 'final', width: 12 },
            { header: 'รวม (100)', key: 'total', width: 15 },
            { header: 'เกรด', key: 'grade', width: 10 }
        ];

        studentsSnap.forEach(doc => {
            const row = doc.data();
            const attCount = attendanceCounts[row.student_id] || 0;
            const attScore = Math.min(attCount, 10);
            const total = attScore + (row.score_assignment || 0) + (row.score_midterm || 0) + (row.score_final || 0);

            sheet.addRow({
                student_id: row.student_id,
                name: row.name,
                att: attScore,
                assign: row.score_assignment || 0,
                mid: row.score_midterm || 0,
                final: row.score_final || 0,
                total: total,
                grade: calculateGrade(total)
            });
        });

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename=Score_${subject}.xlsx`);
        await workbook.xlsx.write(res);
        res.end();
    } catch (err) {
        res.status(500).send(err.message);
    }
});

// ----------------------------------------------------
// สั่งรัน Server (สำหรับ Render Port Binding)
// ----------------------------------------------------
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`🚀 ระบบพร้อมใช้งานบน Port ${PORT}`);
});