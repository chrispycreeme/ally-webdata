import { initializeApp, getApps, getApp } from "https://www.gstatic.com/firebasejs/12.3.0/firebase-app.js";
import {
    getFirestore,
    collection,
    doc,
    getDoc,
    getDocs,
    query,
    where,
    orderBy,
    Timestamp
} from "https://www.gstatic.com/firebasejs/12.3.0/firebase-firestore.js";

// Reuse the existing Firebase project without reinitialising if the page already did it.
const firebaseConfig = {
    apiKey: "AIzaSyA5pyu0LlfvW06m1jdwVXVW8JlW5G7eXps",
    authDomain: "ally-database-669a4.firebaseapp.com",
    projectId: "ally-database-669a4",
    storageBucket: "ally-database-669a4.firebasestorage.app",
    messagingSenderId: "954693497275",
    appId: "1:954693497275:web:16cf70e20465170949149e",
    measurementId: "G-1RGKL7S7X7"
};

const app = getApps().length ? getApp() : initializeApp(firebaseConfig);
const db = getFirestore(app);

export const SCHOOL_NAME = "Regional Science High School";
export const FIXED_CLASS_DAYS = 30;
export const PRESENT_MINUTES_THRESHOLD = 15;

export function parseTimeString(timeStr) {
    const twelveHourMatch = timeStr.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/i);
    if (twelveHourMatch) {
        let hour = parseInt(twelveHourMatch[1], 10);
        const minute = parseInt(twelveHourMatch[2], 10);
        const period = twelveHourMatch[3].toUpperCase();
        if (period === "PM" && hour !== 12) hour += 12;
        if (period === "AM" && hour === 12) hour = 0;
        return { hour, minute };
    }

    const twentyFourMatch = timeStr.match(/^(\d{1,2}):(\d{2})$/);
    if (twentyFourMatch) {
        const hour = parseInt(twentyFourMatch[1], 10);
        const minute = parseInt(twentyFourMatch[2], 10);
        if (hour > 23 || minute > 59) return null;
        return { hour, minute };
    }

    return null;
}

export function parseClassSchedule(classHours) {
    if (!classHours || typeof classHours !== "string" || !classHours.includes("-")) {
        return null;
    }

    const [startRaw, endRaw] = classHours.split("-");
    const start = parseTimeString(startRaw.trim());
    const end = parseTimeString(endRaw.trim());

    if (!start || !end) {
        return null;
    }

    const startMinutes = start.hour * 60 + start.minute;
    const endMinutes = end.hour * 60 + end.minute;
    if (endMinutes <= startMinutes) {
        return null;
    }

    return {
        startHour: start.hour,
        startMinute: start.minute,
        endHour: end.hour,
        endMinute: end.minute
    };
}

export function buildClassWindow(scheduleTemplate, anchorDate, dayNumber) {
    if (!scheduleTemplate) return null;
    const { startHour, startMinute, endHour, endMinute } = scheduleTemplate;
    const year = anchorDate.getFullYear();
    const monthIndex = anchorDate.getMonth();

    const start = new Date(year, monthIndex, dayNumber, startHour, startMinute, 0);
    const end = new Date(year, monthIndex, dayNumber, endHour, endMinute, 0);

    if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime()) || end <= start) {
        return null;
    }

    return { start, end };
}

export function calculateMinutesInside(timeline, classStart, classEnd, initialStatus = "Outside School") {
    if (!timeline || !timeline.length) {
        return initialStatus === "Inside School"
            ? Math.round((classEnd - classStart) / 60000)
            : 0;
    }

    let currentStatus = initialStatus;
    let lastTimestamp = classStart;
    let minutesInside = 0;

    for (const entry of timeline) {
        const timestamp = entry.timestamp;
        if (timestamp <= classStart) {
            currentStatus = entry.status;
            lastTimestamp = classStart;
            continue;
        }
        if (timestamp >= classEnd) {
            if (currentStatus === "Inside School" && lastTimestamp < classEnd) {
                minutesInside += (classEnd - lastTimestamp) / 60000;
            }
            return Math.round(minutesInside);
        }

        if (currentStatus === "Inside School") {
            minutesInside += (timestamp - lastTimestamp) / 60000;
        }

        currentStatus = entry.status;
        lastTimestamp = timestamp;
    }

    if (currentStatus === "Inside School" && lastTimestamp < classEnd) {
        minutesInside += (classEnd - lastTimestamp) / 60000;
    }

    return Math.round(minutesInside);
}

export async function getPlannedAbsencesMap(studentId, year, monthNumber) {
    const absences = new Map();
    try {
        const snapshot = await getDocs(collection(db, "students", studentId, "plannedAbsences"));
        snapshot.forEach(docSnap => {
            const docId = docSnap.id;
            if (!/^\d{8}$/.test(docId)) return;
            const entryYear = parseInt(docId.slice(0, 4), 10);
            const entryMonth = parseInt(docId.slice(4, 6), 10);
            const entryDay = parseInt(docId.slice(6, 8), 10);
            if (entryYear === year && entryMonth === monthNumber) {
                absences.set(entryDay, docSnap.data() || {});
            }
        });
    } catch (err) {
        console.error(`Unable to load planned absences for ${studentId}`, err);
    }
    return absences;
}

export async function getStudentIdsForTeacher(teacherId, gradeLevel) {
    // 1) teacherAssignments/{teacherId}.studentIds
    try {
        const assignRef = doc(db, 'teacherAssignments', teacherId);
        const assignSnap = await getDoc(assignRef);
        if (assignSnap.exists()) {
            const data = assignSnap.data() || {};
            if (Array.isArray(data.studentIds) && data.studentIds.length) return data.studentIds;
        }
    } catch (err) {
        console.warn('teacherAssignments lookup failed', err);
    }

    // 2) Query students where adviserId or teacherId fields match
    try {
        const matches = new Set();
        const q1 = query(collection(db, 'students'), where('adviserId', '==', teacherId));
        const snap1 = await getDocs(q1);
        snap1.forEach(d => matches.add(d.id));

        const q2 = query(collection(db, 'students'), where('teacherId', '==', teacherId));
        const snap2 = await getDocs(q2);
        snap2.forEach(d => matches.add(d.id));

        if (matches.size) return Array.from(matches);
    } catch (err) {
        console.warn('Querying students by adviser/teacher field failed', err);
    }

    // 3) Fallback: query by gradeLevel if provided
    if (gradeLevel) {
        try {
            const snap = await getDocs(query(collection(db, 'students'), where('gradeLevel', '==', gradeLevel)));
            const ids = [];
            snap.forEach(d => ids.push(d.id));
            if (ids.length) return ids;
        } catch (err) {
            console.warn('Querying students by gradeLevel failed', err);
        }
    }

    return [];
}

export async function buildReportFromFirestore(opts = {}) {
    const { teacherId: inputTeacherId, schoolYear: inputSchoolYear, gradeLevel: inputGradeLevel, month: inputMonth } = opts;
    // Read form inputs if not provided
    const teacherId = inputTeacherId || (typeof document !== 'undefined' ? document.getElementById('teacherId')?.value : null);
    let schoolYear = inputSchoolYear || (typeof document !== 'undefined' ? document.getElementById('schoolYear')?.value : null);
    const gradeLevel = inputGradeLevel || (typeof document !== 'undefined' ? document.getElementById('gradeLevel')?.value : '') || '';
    let month = inputMonth || (typeof document !== 'undefined' ? parseInt(document.getElementById('month')?.value || (new Date().getMonth() + 1), 10) : (new Date().getMonth() + 1));

    if (!schoolYear) schoolYear = `${new Date().getFullYear()}-${new Date().getFullYear() + 1}`;
    let year = parseInt(schoolYear.split('-')[0], 10);
    if (!Number.isInteger(year) || year <= 0) year = new Date().getFullYear();
    if (!Number.isInteger(month) || month < 1 || month > 12) month = new Date().getMonth() + 1;

    if (!teacherId) throw new Error('Please enter Teacher ID');

    // Resolve student IDs
    const studentIds = await getStudentIdsForTeacher(teacherId, gradeLevel);

    const monthDays = [];
    const daysInMonth = new Date(year, month, 0).getDate();
    for (let d = 1; d <= daysInMonth; d++) {
        const dt = new Date(year, month - 1, d);
        const isWeekend = dt.getDay() === 0 || dt.getDay() === 6;
        if (isWeekend) continue;
        monthDays.push({ day: d, date: dt.toISOString().slice(0, 10), isWeekend: false, weekday: dt.toLocaleString('en-US', { weekday: 'short' }) });
    }

    const students = [];
    const classDaysCount = monthDays.length;
    const monthStart = new Date(year, month - 1, 1);
    const monthEnd = new Date(year, month - 1, daysInMonth, 23, 59, 59, 999);

    for (const sid of studentIds) {
        try {
            const sDocRef = doc(db, 'students', sid);
            const sSnap = await getDoc(sDocRef);
            if (!sSnap.exists()) continue;
            const sData = sSnap.data() || {};
            const name = sData.name || `${sData.lastName || ''}, ${sData.firstName || ''}`.trim();
            const classHoursStr = sData.classHours || sData.class_hours || null;
            const scheduleTemplate = parseClassSchedule(classHoursStr);

            const historyCol = collection(db, 'students', sid, 'history');
            const q = query(historyCol, where('timestamp', '>=', Timestamp.fromDate(monthStart)), where('timestamp', '<=', Timestamp.fromDate(monthEnd)), orderBy('timestamp', 'asc'));
            const historySnap = await getDocs(q);
            const entries = [];
            historySnap.forEach(d => {
                const data = d.data();
                if (!data || !data.timestamp) return;
                entries.push({ timestamp: data.timestamp.toDate(), type: data.type || null, status: data.status || null });
            });

            const daily = [];
            const plannedAbsences = await getPlannedAbsencesMap(sid, year, month);

            for (let i = 0; i < monthDays.length; i++) {
                const md = monthDays[i];
                const window = buildClassWindow(scheduleTemplate, monthStart, md.day);
                let minutesInside = 0;
                if (window) {
                    const dayStart = new Date(window.start.getFullYear(), window.start.getMonth(), window.start.getDate(), 0, 0, 0);
                    const dayEnd = new Date(window.end.getFullYear(), window.end.getMonth(), window.end.getDate(), 23, 59, 59, 999);
                    const dayEntries = entries.filter(e => e.timestamp >= dayStart && e.timestamp <= dayEnd).map(e => ({ timestamp: e.timestamp, status: e.status }));
                    minutesInside = calculateMinutesInside(dayEntries, window.start, window.end, 'Outside School');
                }

                const planned = plannedAbsences.get(md.day);
                const isPlanned = Boolean(planned);

                let status = '';
                let code = '';
                let remarks = '';

                if (isPlanned) {
                    status = 'Excused';
                    code = 'E';
                    remarks = planned.reason || 'Planned';
                } else if (minutesInside >= PRESENT_MINUTES_THRESHOLD) {
                    status = 'present';
                    code = '';
                } else if (minutesInside > 0 && minutesInside < PRESENT_MINUTES_THRESHOLD) {
                    status = 'tardy';
                    code = 'T';
                } else {
                    status = 'absent';
                    code = 'X';
                }

                daily.push({ status, code, minutes: minutesInside, remarks });
            }

            const totals = daily.reduce((acc, r) => {
                if (r.status === 'present') acc.present += 1;
                if (r.status === 'absent') acc.absent += 1;
                if (r.status === 'Excused' || r.status === 'excused') acc.excused += 1;
                if (r.status === 'tardy') acc.tardy += 1;
                return acc;
            }, { present: 0, absent: 0, excused: 0, tardy: 0 });

            students.push({ id: sid, name, daily, totals, remarks: '' , hasSchedule: !!scheduleTemplate});
        } catch (err) {
            console.warn('Failed to process student', sid, err);
            continue;
        }
    }

    const columnTotals = monthDays.map((md, idx) => {
        const present = students.reduce((acc, s) => acc + (s.daily[idx].status === 'present' ? 1 : 0), 0);
        const absent = students.reduce((acc, s) => acc + (s.daily[idx].status === 'absent' ? 1 : 0), 0);
        const excused = students.reduce((acc, s) => acc + (s.daily[idx].status === 'Excused' || s.daily[idx].status === 'excused' ? 1 : 0), 0);
        const tardy = students.reduce((acc, s) => acc + (s.daily[idx].status === 'tardy' ? 1 : 0), 0);
        return { present, absent, excused, tardy };
    });

    const totalsAcross = columnTotals.reduce((acc, t) => {
        acc.present += t.present || 0;
        acc.absent += t.absent || 0;
        acc.excused += t.excused || 0;
        acc.tardy += t.tardy || 0;
        return acc;
    }, { present: 0, absent: 0, excused: 0, tardy: 0 });

    const totalLearners = students.length;
    const attendanceRate = classDaysCount > 0 && totalLearners > 0
        ? (totalsAcross.present / (classDaysCount * totalLearners)) * 100
        : 0;

    const aggregates = {
        totalLearners,
        totals: totalsAcross,
        attendanceRate
    };

    return {
        teacherId,
        reportDate: new Date().toISOString().slice(0, 10),
        schoolYear,
        gradeLevel,
        month,
        monthDays,
        students,
        columnTotals,
        aggregates,
        schoolName: SCHOOL_NAME,
        noOfClassDays: FIXED_CLASS_DAYS,
        anchorDate: monthStart
    };
}
