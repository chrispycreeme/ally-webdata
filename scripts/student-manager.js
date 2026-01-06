import { addStudent } from "./sf2-firestore.js";

const form = document.getElementById("studentForm");
const messageEl = document.getElementById("studentMessage");

function init() {
    if (form) {
        form.addEventListener("submit", handleAddStudent);
    }
}

async function handleAddStudent(e) {
    e.preventDefault();
    if (messageEl) {
        messageEl.textContent = "Adding student...";
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
        await addStudent(data);
        if (messageEl) {
            messageEl.textContent = `Student ${data.firstName} ${data.lastName} added successfully!`;
            messageEl.style.color = "green";
        }
        form.reset();
    } catch (err) {
        if (messageEl) {
            messageEl.textContent = `Error: ${err.message}`;
            messageEl.style.color = "red";
        }
        console.error(err);
    }
}

// Initialize when DOM is ready
if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
} else {
    init();
}
