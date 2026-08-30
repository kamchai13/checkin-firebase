const express = require('express');
const cors = require('cors');
const ExcelJS = require('exceljs');
const path = require('path');
const admin = require('firebase-admin');

const app = express();
app.use(cors());
app.use(express.json());
app.use(express.static(__dirname));

// ----------------------------------------------------
// เชื่อมต่อ Firebase Firestore
// ----------------------------------------------------
let serviceAccount;

if (process.env.FIREBASE_CONFIG) {
  // สำหรับใช้งานบน Render
  const rawConfig = process.env.FIREBASE_CONFIG;
  const configString = rawConfig.replace(/\n/g, "\\n");
  serviceAccount = JSON.parse(configString);
} else {
  // สำหรับรับบนเครื่อง Local
  try {
    serviceAccount = require('./serviceAccountKey.json');
  } catch (err) {
    console.error("Local serviceAccountKey.json not found!");
  }
}

admin.initializeApp({
  credential: admin.credential.cert(serviceAccount)
});

const db = admin.firestore();
const studentsCol = db.collection('students');
const attendanceCol = db.collection('attendance');

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
// API ระบบเช็คชื่อ (Firebase Firestore)
// ----------------------------------------------------
app.post('/api/checkin', async (req, res) => {
    try {
        const { courseKey, studentId } = req.body;
        
        if (!studentId) {
            return res.status(400).json({ message: "กรุณากรอกรหัสนักศึกษา" });
        }

        const cleanStudentId = studentId.toString().trim();
        const cleanCourseKey = courseKey ? courseKey.toString().trim() : '';

        // ค้นหานักศึกษาใน Firebase
        const docRef = studentsCol.doc(cleanStudentId);
        const docSnap = await docRef.get();

        if (!docSnap.exists) {
            return res.status(404).json({ message: `ไม่พบรหัสนักศึกษา ${cleanStudentId} ในระบบ` });
        }

        const studentData = docSnap.data();
        const subjectToUse = cleanCourseKey || studentData.subject || 'GENERAL';

        // บันทึกเวลาเข้าเรียน
        await attendanceCol.add({
            student_id: cleanStudentId,
            subject: subjectToUse,
            timestamp: admin.firestore.FieldValue.serverTimestamp()
        });

        res.json({ message: `เช็คชื่อสำเร็จ! ยินดีต้อนรับ ${studentData.name || cleanStudentId}` });
    } catch (err) {
        console.error(err);
        res.status(500).json({ message: "เกิดข้อผิดพลาดทางเซิร์ฟเวอร์" });
    }
});

// API ดึงรายชื่อนักศึกษา
app.get('/api/students/:subject', async (req, res) => {
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
            attendanceCounts[sid] = (attendanceCounts[sid] || 0) + 1;
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
app.post('/api/scores/update', async (req, res) => {
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
app.post('/api/scan', async (req, res) => {
    try {
        const { student_id, subject } = req.body;
        if (!student_id) return res.status(400).json({ error: "ข้อมูลไม่ครบถ้วน" });

        const cleanId = student_id.toString().trim();

        await attendanceCol.add({
            student_id: cleanId,
            subject: subject || 'GENERAL',
            timestamp: admin.firestore.FieldValue.serverTimestamp()
        });

        res.json({ message: `เช็คชื่อสำเร็จ! รหัส ${cleanId}` });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// API เพิ่มนักศึกษา
app.post('/api/students', async (req, res) => {
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
app.get('/api/export/:subject', async (req, res) => {
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
            attendanceCounts[sid] = (attendanceCounts[sid] || 0) + 1;
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
            { header: 'รวม (100)', key: 'total', width: 100 },
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

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`🚀 ระบบพร้อมใช้งานถาวรด้วย Firebase ที่ port ${PORT}`));