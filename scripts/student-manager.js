import { addStudent, getAllStudents, updateStudent, deleteStudent } from "./sf2-firestore.js";

const form = document.getElementById("studentForm");
const messageEl = document.getElementById("studentMessage");
const cancelEditBtn = document.getElementById("cancelEdit");
const refreshBtn = document.getElementById("refreshStudentList");
const studentTableBody = document.getElementById("studentTableBody");
const gradeFilter = document.getElementById("gradeFilter");
const nameSearch = document.getElementById("nameSearch");

let students = [];
let editingStudentId = null;

function init() {
    if (form) {
        form.addEventListener("submit", handleAddOrUpdateStudent);
    }
    if (cancelEditBtn) {
        cancelEditBtn.addEventListener("click", cancelEdit);
    }
    if (refreshBtn) {
        refreshBtn.addEventListener("click", loadStudents);
    }
    if (gradeFilter) {
        gradeFilter.addEventListener("change", filterStudents);
    }
    if (nameSearch) {
        nameSearch.addEventListener("input", filterStudents);
    }
}

async function handleAddOrUpdateStudent(e) {
    e.preventDefault();
    if (messageEl) {
        messageEl.textContent = editingStudentId ? "Updating student..." : "Adding student...";
        messageEl.style.color = "blue";
    }

    const data = {
        lrn: document.getElementById("studentLrn").value.trim(),
        firstName: document.getElementById("studentFirstName").value.trim(),
        middleName: document.getElementById("studentMiddleName").value.trim(),
        lastName: document.getElementById("studentLastName").value.trim(),
        gradeLevel: document.getElementById("studentGradeLevel").value.trim(),
        teacherId: document.getElementById("studentTeacherId").value.trim(),
        classHours: document.getElementById("studentClassHours").value.trim()
    };

    try {
        if (editingStudentId) {
            // Update existing student
            await updateStudent(editingStudentId, data);
            if (messageEl) {
                messageEl.textContent = `Student ${data.firstName} ${data.lastName} updated successfully!`;
                messageEl.style.color = "green";
            }
        } else {
            // Add new student
            await addStudent(data);
            if (messageEl) {
                messageEl.textContent = `Student ${data.firstName} ${data.lastName} added successfully!`;
                messageEl.style.color = "green";
            }
        }
        form.reset();
        cancelEdit();
        await loadStudents();
    } catch (err) {
        if (messageEl) {
            messageEl.textContent = `Error: ${err.message}`;
            messageEl.style.color = "red";
        }
        console.error(err);
    }
}

function cancelEdit() {
    editingStudentId = null;
    form.reset();
    const submitBtn = form.querySelector('button[type="submit"]');
    if (submitBtn) {
        submitBtn.textContent = "Add Student";
    }
    cancelEditBtn.style.display = "none";
    document.getElementById("studentLrn").disabled = false;
}

async function loadStudents() {
    try {
        students = await getAllStudents();
        updateGradeFilter();
        renderStudents(students);
    } catch (err) {
        console.error("Error loading students:", err);
        studentTableBody.innerHTML = '<tr><td colspan="6" style="color: red;">Error loading students</td></tr>';
    }
}

function updateGradeFilter() {
    const grades = [...new Set(students.map(s => s.gradeLevel).filter(g => g))].sort();
    gradeFilter.innerHTML = '<option value="">All Grades</option>';
    grades.forEach(grade => {
        const option = document.createElement("option");
        option.value = grade;
        option.textContent = grade;
        gradeFilter.appendChild(option);
    });
}

function filterStudents() {
    const gradeValue = gradeFilter.value;
    const searchValue = nameSearch.value.toLowerCase().trim();
    
    let filtered = students;
    
    if (gradeValue) {
        filtered = filtered.filter(s => s.gradeLevel === gradeValue);
    }
    
    if (searchValue) {
        filtered = filtered.filter(s => 
            s.name.toLowerCase().includes(searchValue) ||
            s.firstName.toLowerCase().includes(searchValue) ||
            s.lastName.toLowerCase().includes(searchValue)
        );
    }
    
    renderStudents(filtered);
}

function renderStudents(studentList) {
    if (!studentTableBody) return;
    
    if (studentList.length === 0) {
        studentTableBody.innerHTML = '<tr><td colspan="6" style="text-align: center;">No students found</td></tr>';
        return;
    }
    
    studentTableBody.innerHTML = studentList.map(student => `
        <tr>
            <td style="border: 1px solid #d1d5db; padding: 8px;">${student.id}</td>
            <td style="border: 1px solid #d1d5db; padding: 8px;">${student.name}</td>
            <td style="border: 1px solid #d1d5db; padding: 8px;">${student.gradeLevel || 'N/A'}</td>
            <td style="border: 1px solid #d1d5db; padding: 8px;">${student.teacherId || 'N/A'}</td>
            <td style="border: 1px solid #d1d5db; padding: 8px;">${student.classHours || 'N/A'}</td>
            <td style="border: 1px solid #d1d5db; padding: 8px; text-align: center;">
                <button onclick="editStudent('${student.id}')" style="padding: 4px 8px; margin-right: 4px; background-color: #3B82F6; color: white; border: none; border-radius: 3px; cursor: pointer;">Edit</button>
                <button onclick="deleteStudent('${student.id}')" style="padding: 4px 8px; background-color: #EF4444; color: white; border: none; border-radius: 3px; cursor: pointer;">Delete</button>
            </td>
        </tr>
    `).join('');
}

window.editStudent = function(lrn) {
    const student = students.find(s => s.id === lrn);
    if (!student) return;
    
    editingStudentId = lrn;
    
    document.getElementById("studentLrn").value = student.id;
    document.getElementById("studentLrn").disabled = true;
    document.getElementById("studentFirstName").value = student.firstName || '';
    document.getElementById("studentMiddleName").value = student.middleName || '';
    document.getElementById("studentLastName").value = student.lastName || '';
    document.getElementById("studentGradeLevel").value = student.gradeLevel || '';
    document.getElementById("studentTeacherId").value = student.teacherId || '';
    document.getElementById("studentClassHours").value = student.classHours || '';
    
    const submitBtn = form.querySelector('button[type="submit"]');
    if (submitBtn) {
        submitBtn.textContent = "Update Student";
    }
    cancelEditBtn.style.display = "inline-block";
    
    // Scroll to form
    form.scrollIntoView({ behavior: 'smooth' });
};

window.deleteStudent = async function(lrn) {
    if (!confirm(`Are you sure you want to delete student with LRN ${lrn}?`)) {
        return;
    }
    
    try {
        await deleteStudent(lrn);
        messageEl.textContent = "Student deleted successfully!";
        messageEl.style.color = "green";
        await loadStudents();
    } catch (err) {
        messageEl.textContent = `Error: ${err.message}`;
        messageEl.style.color = "red";
        console.error(err);
    }
};

// Initialize when DOM is ready
if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
} else {
    init();
}

// Load students on page load
loadStudents();
