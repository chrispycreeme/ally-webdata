import { buildReportFromFirestore, SCHOOL_NAME } from "./sf2-firestore.js";

let excelJsModulePromise = null;

async function getExcelJsModule() {
	if (window.ExcelJS) {
		return window.ExcelJS;
	}
	if (excelJsModulePromise) {
		try {
			return await excelJsModulePromise;
		} catch (err) {
			excelJsModulePromise = null;
			throw err;
		}
	}
	excelJsModulePromise = (async () => {
		const sources = [
			"https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js",
			"https://cdn.jsdelivr.net/npm/exceljs@4/dist/exceljs.min.js"
		];
		for (const url of sources) {
			try {
				await loadExternalScript(url);
				if (window.ExcelJS) {
					return window.ExcelJS;
				}
			} catch (err) {
				console.warn(`Failed loading ExcelJS from ${url}`, err);
			}
		}
		throw new Error("Unable to load ExcelJS library. Please check your internet connection and try again later.");
	})();
	try {
		return await excelJsModulePromise;
	} catch (err) {
		excelJsModulePromise = null;
		throw err;
	}
}

function loadExternalScript(url) {
	return new Promise((resolve, reject) => {
		const script = document.createElement("script");
		script.src = url;
		script.async = true;
		script.onload = () => resolve();
		script.onerror = () => reject(new Error(`Failed to load script: ${url}`));
		document.head.appendChild(script);
	});
}

const MONTH_NAMES = [
	"January",
	"February",
	"March",
	"April",
	"May",
	"June",
	"July",
	"August",
	"September",
	"October",
	"November",
	"December"
];

const state = {
	reportData: null,
	templateBuffer: null,
	templateName: ""
};

const ui = {
	teacherId: document.getElementById("teacherId"),
	reportDate: document.getElementById("reportDate"),
	gradeLevel: document.getElementById("gradeLevel"),
	schoolYear: document.getElementById("schoolYear"),
	month: document.getElementById("month"),
	generateButton: document.getElementById("generateReport"),
	exportExcelButton: document.getElementById("exportExcel"),
	exportCsvButton: document.getElementById("exportCSV"),
	templateInput: document.getElementById("templateFile"),
	templateLabel: document.getElementById("templateFileName"),
	reportContainer: document.getElementById("reportContainer"),
	reportHeader: document.getElementById("reportHeader"),
	summaryStats: document.getElementById("summaryStats"),
	reportTable: document.getElementById("reportTable"),
	loading: document.getElementById("loading"),
	error: document.getElementById("error")
};

function init() {
	if (ui.reportDate) {
		const today = new Date();
		const iso = new Date(today.getFullYear(), today.getMonth(), today.getDate()).toISOString().slice(0, 10);
		ui.reportDate.value = iso;
	}

	if (ui.month) {
		const today = new Date();
		ui.month.value = MONTH_NAMES[today.getMonth()];
	}

	if (ui.generateButton) {
		ui.generateButton.addEventListener("click", onGenerateReport);
	}

	if (ui.exportExcelButton) {
		ui.exportExcelButton.addEventListener("click", onExportExcel);
	}

	if (ui.exportCsvButton) {
		ui.exportCsvButton.addEventListener("click", onExportCsv);
	}

	if (ui.templateInput) {
		ui.templateInput.addEventListener("change", onTemplateSelected);
	}

	if (ui.reportContainer) {
		ui.reportContainer.hidden = true;
	}

	updateExportButtons();
}

function parseMonthInput(value, fallbackDate) {
	if (value) {
		const trimmed = value.trim();
		const numeric = parseInt(trimmed, 10);
		if (!Number.isNaN(numeric) && numeric >= 1 && numeric <= 12) {
			return { number: numeric, label: MONTH_NAMES[numeric - 1] };
		}

		const index = MONTH_NAMES.findIndex(name => name.toLowerCase() === trimmed.toLowerCase());
		if (index !== -1) {
			return { number: index + 1, label: MONTH_NAMES[index] };
		}

		const parsed = Date.parse(`${trimmed} 1, 2000`);
		if (!Number.isNaN(parsed)) {
			const date = new Date(parsed);
			return { number: date.getMonth() + 1, label: MONTH_NAMES[date.getMonth()] };
		}
	}

	const fallback = fallbackDate && !Number.isNaN(fallbackDate.getTime()) ? fallbackDate : new Date();
	return { number: fallback.getMonth() + 1, label: MONTH_NAMES[fallback.getMonth()] };
}

function parseGradeSection(value) {
	if (!value) {
		return { grade: "", section: "" };
	}
	const trimmed = value.trim();
	if (!trimmed) {
		return { grade: "", section: "" };
	}
	const parts = trimmed.split(/[-/|]+/).map(part => part.trim()).filter(Boolean);
	if (parts.length >= 2) {
		return { grade: parts[0], section: parts.slice(1).join(" ") };
	}
	return { grade: trimmed, section: "" };
}

function setLoading(isLoading) {
	if (ui.loading) {
		ui.loading.hidden = !isLoading;
	}
	if (ui.generateButton) {
		ui.generateButton.disabled = isLoading;
	}
}

function showError(message) {
	if (!ui.error) return;
	ui.error.textContent = message || "";
	ui.error.hidden = !message;
}

function clearError() {
	showError("");
}

function updateExportButtons() {
	const hasReport = Boolean(state.reportData);
	if (ui.exportCsvButton) {
		ui.exportCsvButton.disabled = !hasReport;
	}
	if (ui.exportExcelButton) {
		ui.exportExcelButton.disabled = !(hasReport && state.templateBuffer);
	}
}

async function onTemplateSelected(event) {
	const input = event.target;
	const file = input && input.files && input.files[0] ? input.files[0] : null;

	state.templateBuffer = null;
	state.templateName = "";
	if (ui.templateLabel) {
		ui.templateLabel.textContent = "";
	}

	if (!file) {
		updateExportButtons();
		return;
	}

	if (!file.name.toLowerCase().endsWith(".xlsx")) {
		showError("Please choose an .xlsx template file.");
		input.value = "";
		updateExportButtons();
		return;
	}

	try {
		state.templateBuffer = await file.arrayBuffer();
		state.templateName = file.name;
		if (ui.templateLabel) {
			ui.templateLabel.textContent = file.name;
		}
		clearError();
		// Kick off the ExcelJS loader early so export waits less when triggered later.
		getExcelJsModule().catch(() => {});
	} catch (err) {
		console.error("Template read failed", err);
		showError("Unable to read the uploaded template.");
		input.value = "";
	}

	updateExportButtons();
}

async function onGenerateReport() {
	if (!ui.teacherId || !ui.schoolYear) {
		return;
	}

	const teacherId = (ui.teacherId.value || "").trim();
	const schoolYear = (ui.schoolYear.value || "").trim();
	const gradeInput = ui.gradeLevel ? (ui.gradeLevel.value || "").trim() : "";

	let referenceDate = null;
	if (ui.reportDate && ui.reportDate.value) {
		const maybeDate = new Date(ui.reportDate.value);
		if (!Number.isNaN(maybeDate.getTime())) {
			referenceDate = maybeDate;
		}
	}

	const { number: monthNumber, label: monthLabel } = parseMonthInput(ui.month ? ui.month.value : "", referenceDate || undefined);

	if (!teacherId) {
		showError("Please enter Teacher ID before generating the report.");
		return;
	}
	if (!schoolYear) {
		showError("Please enter the School Year before generating the report.");
		return;
	}

	setLoading(true);
	clearError();

	try {
		const report = await buildReportFromFirestore({
			teacherId,
			schoolYear,
			gradeLevel: gradeInput,
			month: monthNumber
		});

		const { grade, section } = parseGradeSection(gradeInput);
		report.gradeLevel = grade || report.gradeLevel || gradeInput;
		report.section = section;
		report.monthLabel = monthLabel;
		if (!report.schoolName) {
			report.schoolName = SCHOOL_NAME;
		}

		state.reportData = report;
		renderReport(report);
		updateExportButtons();
	} catch (err) {
		console.error("SF2 generation failed", err);
		showError(err && err.message ? err.message : "Unable to build the report. Please try again.");
		state.reportData = null;
		if (ui.reportContainer) {
			ui.reportContainer.hidden = true;
		}
		updateExportButtons();
	} finally {
		setLoading(false);
	}
}

function renderReport(report) {
	if (!report) {
		if (ui.reportContainer) ui.reportContainer.hidden = true;
		return;
	}

	if (ui.reportContainer) {
		ui.reportContainer.hidden = false;
	}

	renderHeader(report);
	renderSummary(report);
	renderTable(report);
}

function renderHeader(report) {
	if (!ui.reportHeader) return;
	const rows = [
		`<div><strong>School:</strong> ${escapeHtml(report.schoolName || SCHOOL_NAME)}</div>`,
		`<div><strong>School Year:</strong> ${escapeHtml(report.schoolYear || "")}</div>`,
		`<div><strong>Month:</strong> ${escapeHtml(buildMonthLabel(report))}</div>`,
		`<div><strong>Grade Level:</strong> ${escapeHtml(report.gradeLevel || "")}</div>`
	];
	if (report.section) {
		rows.push(`<div><strong>Section:</strong> ${escapeHtml(report.section)}</div>`);
	}
	ui.reportHeader.innerHTML = rows.join("");
}

function renderSummary(report) {
	if (!ui.summaryStats) return;
	const totals = (report.aggregates && report.aggregates.totals) || { present: 0, absent: 0, excused: 0, tardy: 0 };
	const totalLearners = report.aggregates && typeof report.aggregates.totalLearners === "number" ? report.aggregates.totalLearners : 0;
	const attendanceRate = report.aggregates && typeof report.aggregates.attendanceRate === "number"
		? report.aggregates.attendanceRate.toFixed(1)
		: "0.0";

	ui.summaryStats.innerHTML = `
		<div><strong>Total Learners:</strong> ${totalLearners}</div>
		<div><strong>Present:</strong> ${totals.present || 0}</div>
		<div><strong>Absent:</strong> ${totals.absent || 0}</div>
		<div><strong>Excused:</strong> ${totals.excused || 0}</div>
		<div><strong>Tardy:</strong> ${totals.tardy || 0}</div>
		<div><strong>Attendance Rate:</strong> ${attendanceRate}%</div>
	`;
}

function renderTable(report) {
	if (!ui.reportTable) return;

	const container = ui.reportTable;
	container.innerHTML = "";

	const students = Array.isArray(report.students) ? report.students : [];
	const monthDays = Array.isArray(report.monthDays) ? report.monthDays : [];

	if (!students.length) {
		const message = document.createElement("p");
		message.textContent = "No learner records found for the selected filters.";
		container.appendChild(message);
		return;
	}

	const table = document.createElement("table");
	table.className = "sf2-table";

	const thead = document.createElement("thead");
	const headRow = document.createElement("tr");
	["#", "LRN", "Learner Name"].forEach(label => {
		const th = document.createElement("th");
		th.textContent = label;
		headRow.appendChild(th);
	});

	monthDays.forEach(day => {
		const th = document.createElement("th");
		th.textContent = String(day.day);
		th.title = day.weekday || "";
		if (day.isWeekend) {
			th.classList.add("is-weekend");
		}
		headRow.appendChild(th);
	});

	["Present", "Absent", "Excused", "Tardy", "Remarks"].forEach(label => {
		const th = document.createElement("th");
		th.textContent = label;
		headRow.appendChild(th);
	});

	thead.appendChild(headRow);
	table.appendChild(thead);

	const tbody = document.createElement("tbody");

	students.forEach((student, index) => {
		const row = document.createElement("tr");

		const seq = document.createElement("td");
		seq.textContent = String(index + 1);
		row.appendChild(seq);

		const lrnCell = document.createElement("td");
		lrnCell.textContent = student.id || "";
		row.appendChild(lrnCell);

		const nameCell = document.createElement("td");
		nameCell.textContent = student.name || "";
		row.appendChild(nameCell);

		const daily = Array.isArray(student.daily) ? student.daily : [];
		monthDays.forEach((day, dayIndex) => {
			const td = document.createElement("td");
			const entry = daily[dayIndex] || {};
			td.textContent = day.isWeekend ? "" : (entry.code || "");
			if (day.isWeekend) {
				td.classList.add("is-weekend");
			}
			if (entry.status) {
				td.dataset.status = entry.status;
			}
			row.appendChild(td);
		});

		const totals = student.totals || {};
		const presentCell = document.createElement("td");
		presentCell.textContent = totals.present != null ? String(totals.present) : "";
		row.appendChild(presentCell);

		const absentCell = document.createElement("td");
		absentCell.textContent = totals.absent != null ? String(totals.absent) : "";
		row.appendChild(absentCell);

		const excusedCell = document.createElement("td");
		excusedCell.textContent = totals.excused != null ? String(totals.excused) : "";
		row.appendChild(excusedCell);

		const tardyCell = document.createElement("td");
		tardyCell.textContent = totals.tardy != null ? String(totals.tardy) : "";
		row.appendChild(tardyCell);

		const remarksCell = document.createElement("td");
		remarksCell.textContent = student.remarks || "";
		row.appendChild(remarksCell);

		tbody.appendChild(row);
	});

	table.appendChild(tbody);

	if (Array.isArray(report.columnTotals) && report.columnTotals.length) {
		const tfoot = document.createElement("tfoot");
		const totalsRow = document.createElement("tr");

		const labelCell = document.createElement("td");
		labelCell.colSpan = 3;
		labelCell.textContent = "Daily Present";
		totalsRow.appendChild(labelCell);

		report.columnTotals.forEach((colTotals, index) => {
			const td = document.createElement("td");
			const present = colTotals && colTotals.present ? String(colTotals.present) : "";
			td.textContent = present;
			if (monthDays[index] && monthDays[index].isWeekend) {
				td.classList.add("is-weekend");
			}
			totalsRow.appendChild(td);
		});

		const aggTotals = (report.aggregates && report.aggregates.totals) || {};
		[aggTotals.present, aggTotals.absent, aggTotals.excused, aggTotals.tardy].forEach(value => {
			const td = document.createElement("td");
			td.textContent = value != null ? String(value) : "";
			totalsRow.appendChild(td);
		});

		const remarksTd = document.createElement("td");
		remarksTd.textContent = "";
		totalsRow.appendChild(remarksTd);

		tfoot.appendChild(totalsRow);
		table.appendChild(tfoot);
	}

	container.appendChild(table);
}

function escapeHtml(value) {
	return String(value ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}

function buildMonthLabel(report) {
	if (report.monthLabel && report.anchorDate) {
		const anchor = normaliseDate(report.anchorDate);
		if (anchor) {
			return `${report.monthLabel} ${anchor.getFullYear()}`;
		}
	}
	if (report.monthLabel) {
		return report.monthLabel;
	}
	const anchor = normaliseDate(report.anchorDate);
	if (anchor) {
		return anchor.toLocaleString("en-US", { month: "long", year: "numeric" });
	}
	if (typeof report.month === "number" && report.month >= 1 && report.month <= 12) {
		const today = new Date();
		return `${MONTH_NAMES[report.month - 1]} ${today.getFullYear()}`;
	}
	return "";
}

function normaliseDate(value) {
	if (!value) return null;
	if (value instanceof Date) {
		return Number.isNaN(value.getTime()) ? null : value;
	}
	const parsed = new Date(value);
	return Number.isNaN(parsed.getTime()) ? null : parsed;
}

async function onExportExcel() {
	if (!state.reportData) {
		showError("Generate the report before exporting.");
		return;
	}
	if (!state.templateBuffer) {
		showError("Upload an SF2 template before exporting to Excel.");
		return;
	}

	clearError();

	try {
		const buffer = await buildExcelFromTemplate(state.reportData);
		const blob = new Blob([buffer], {
			type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
		});
		triggerDownload(blob, buildFileName(state.reportData, "xlsx"));
	} catch (err) {
		console.error("Excel export failed", err);
		showError(err && err.message ? err.message : "Unable to build the Excel file. Please verify the template and try again.");
	}
}

function onExportCsv() {
	if (!state.reportData) {
		showError("Generate the report before exporting.");
		return;
	}

	clearError();

	try {
		const csvContent = buildCsv(state.reportData);
		const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8" });
		triggerDownload(blob, buildFileName(state.reportData, "csv"));
	} catch (err) {
		console.error("CSV export failed", err);
		showError("Unable to build the CSV file. Please try again.");
	}
}

function buildFileName(report, extension) {
	const parts = ["SF2"];
	if (report.gradeLevel) parts.push(report.gradeLevel.replace(/\s+/g, ""));
	if (report.section) parts.push(report.section.replace(/\s+/g, ""));
	if (report.monthLabel) parts.push(report.monthLabel.replace(/\s+/g, ""));
	const base = parts.filter(Boolean).join("_") || "SF2_Report";
	return `${base}.${extension}`;
}

function triggerDownload(blob, fileName) {
	const link = document.createElement("a");
	link.href = URL.createObjectURL(blob);
	link.download = fileName;
	document.body.appendChild(link);
	link.click();
	setTimeout(() => {
		URL.revokeObjectURL(link.href);
		link.remove();
	}, 0);
}

function buildCsv(report) {
	const students = Array.isArray(report.students) ? report.students : [];
	const monthDays = Array.isArray(report.monthDays) ? report.monthDays : [];

	const header = ["#", "Learner Name", ...monthDays.map(day => `Day ${day.day}`), "Present", "Absent", "Excused", "Tardy", "Remarks"];
	const rows = [header];

	students.forEach(student => {
		const daily = Array.isArray(student.daily) ? student.daily : [];
		const totals = student.totals || {};
		const row = [
			escapeCsvValue(index + 1),
			escapeCsvValue(student.name || "")
		];

		monthDays.forEach((day, index) => {
			const entry = daily[index] || {};
			row.push(escapeCsvValue(day.isWeekend ? "" : (entry.code || "")));
		});

		row.push(escapeCsvValue(totals.present != null ? totals.present : ""));
		row.push(escapeCsvValue(totals.absent != null ? totals.absent : ""));
		row.push(escapeCsvValue(totals.excused != null ? totals.excused : ""));
		row.push(escapeCsvValue(totals.tardy != null ? totals.tardy : ""));
		row.push(escapeCsvValue(student.remarks || ""));

		rows.push(row);
	});

	return rows.map(columns => columns.join(",")).join("\r\n");
}

function escapeCsvValue(value) {
	const text = String(value ?? "");
	if (text.includes("\"") || text.includes(",") || /[\r\n]/.test(text)) {
		return `"${text.replace(/"/g, '""')}"`;
	}
	return text;
}

async function buildExcelFromTemplate(report) {
	if (!state.templateBuffer) {
		throw new Error("Template buffer is missing.");
	}

	const ExcelJS = await getExcelJsModule();
	const workbook = new ExcelJS.Workbook();
	await workbook.xlsx.load(state.templateBuffer.slice(0));
	const worksheet = workbook.worksheets[0] || workbook.addWorksheet("Sheet1");

	const getWorksheetBounds = () => {
		let minRow = Number.POSITIVE_INFINITY;
		let minCol = Number.POSITIVE_INFINITY;
		let maxRow = -1;
		let maxCol = -1;
		worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
			minRow = Math.min(minRow, rowNumber - 1);
			maxRow = Math.max(maxRow, rowNumber - 1);
			row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
				minCol = Math.min(minCol, colNumber - 1);
				maxCol = Math.max(maxCol, colNumber - 1);
			});
		});
		if (maxRow === -1 || maxCol === -1) {
			const fallbackRow = Math.max(worksheet.rowCount - 1, 0);
			const fallbackCol = Math.max(worksheet.columnCount - 1, 0);
			return { s: { r: 0, c: 0 }, e: { r: fallbackRow, c: fallbackCol } };
		}
		return { s: { r: minRow, c: minCol }, e: { r: maxRow, c: maxCol } };
	};

	const extractCellText = (rowIndex, colIndex) => {
		if (rowIndex < 0 || colIndex < 0) return "";
		let cell;
		try {
			cell = worksheet.getCell(rowIndex + 1, colIndex + 1);
		} catch (err) {
			return "";
		}
		if (!cell || !cell.model) return "";
		let text = "";
		try {
			text = cell.text || "";
		} catch (err) {
			text = "";
		}
		if (typeof text === "string" && text.trim()) {
			return text.trim();
		}
		const value = cell.value;
		if (value == null) return "";
		if (typeof value === "object") {
			if (Array.isArray(value.richText)) {
				return value.richText.map(part => part.text || "").join("").trim();
			}
			if (value.result != null) {
				return String(value.result).trim();
			}
			if (value.text != null) {
				return String(value.text).trim();
			}
		}
		return String(value).trim();
	};

	const originalRange = getWorksheetBounds();

	const findHeaderRow = () => {
		const maxRow = Math.min(originalRange.e.r, originalRange.s.r + 40);
		for (let r = originalRange.s.r; r <= maxRow; r++) {
			for (let c = originalRange.s.c; c <= originalRange.e.c; c++) {
				const text = extractCellText(r, c).toLowerCase();
				if (text.includes("learner") && text.includes("name")) {
					return r;
				}
			}
		}
		return 12;
	};

	const headerRow = findHeaderRow();

	const findColumn = (predicate, rows) => {
		const rowsToCheck = Array.isArray(rows) && rows.length ? rows : [headerRow];
		for (const row of rowsToCheck) {
			if (row < originalRange.s.r || row > originalRange.e.r) continue;
			for (let col = originalRange.s.c; col <= originalRange.e.c; col++) {
				const text = extractCellText(row, col).toLowerCase();
				if (!text) continue;
				if (predicate(text)) {
					return col;
				}
			}
		}
		return null;
	};

	const detectedNameCol = findColumn(text => text.includes("learner") && text.includes("name"), [headerRow]);
	const detectedLrnCol = findColumn(text => text.includes("lrn"), [headerRow, headerRow + 1]);
	const detectedNumberCol = findColumn(text => text.startsWith("no"), [headerRow, headerRow + 1]);
	const detectedSexCol = findColumn(text => text.includes("sex") || text.includes("gender"), [headerRow, headerRow + 1]);
	const detectedRemarksCol = findColumn(text => text.includes("remark"), [headerRow, headerRow + 1, headerRow + 2]);
	const detectedPresentCol = findColumn(text => text.includes("present") && !text.includes("percentage"), [headerRow, headerRow + 1, headerRow + 2]);
	const detectedAbsentCol = findColumn(text => text.includes("absent"), [headerRow, headerRow + 1, headerRow + 2]);
	const detectedExcusedCol = findColumn(text => text.includes("excused"), [headerRow, headerRow + 1, headerRow + 2]);
	const detectedTardyCol = findColumn(text => text.includes("tardy"), [headerRow, headerRow + 1, headerRow + 2]);

	let numberCol = Number.isInteger(detectedNumberCol) ? detectedNumberCol : null;
	let nameCol = Number.isInteger(detectedNameCol) ? detectedNameCol : null;

	if (numberCol == null && nameCol == null) {
		numberCol = 0;
		nameCol = 1;
	} else if (numberCol == null && nameCol != null) {
		numberCol = Math.max(nameCol - 1, 0);
	} else if (numberCol != null && nameCol == null) {
		nameCol = numberCol + 1;
	}

	if (nameCol <= numberCol) {
		nameCol = numberCol + 1;
	}

	const lrnCol = Number.isInteger(detectedLrnCol) && detectedLrnCol !== numberCol && detectedLrnCol !== nameCol
		? detectedLrnCol
		: nameCol + 1;
	const sexCol = Number.isInteger(detectedSexCol) ? detectedSexCol : null;

	const students = Array.isArray(report.students) ? report.students : [];
	const monthDays = Array.isArray(report.monthDays) ? report.monthDays : [];

	const dayColumnMap = new Map();
	const searchEndCol = originalRange.e.c;
	for (let r = headerRow; r <= Math.min(headerRow + 6, originalRange.e.r); r++) {
		for (let c = nameCol + 1; c <= searchEndCol; c++) {
			const text = extractCellText(r, c);
			const dayNumber = parseInt(text, 10);
			if (!Number.isNaN(dayNumber) && dayNumber >= 1 && dayNumber <= 31 && !dayColumnMap.has(dayNumber)) {
				dayColumnMap.set(dayNumber, c);
			}
		}
	}
	if (!dayColumnMap.size && monthDays.length) {
		let startCol = Math.max(nameCol + 1, 3);
		monthDays.forEach((day, idx) => {
			dayColumnMap.set(day.day, startCol + idx);
		});
	}

	const sortedDayColumns = Array.from(new Set(dayColumnMap.values())).sort((a, b) => a - b);
	const firstDayColumnIndex = sortedDayColumns.length ? sortedDayColumns[0] : Math.max(nameCol + 1, 2);
	const lastDayColumnIndex = sortedDayColumns.length
		? sortedDayColumns[sortedDayColumns.length - 1]
		: firstDayColumnIndex + Math.max(monthDays.length || 0, 1) - 1;

	const remarksCol = Number.isInteger(detectedRemarksCol)
		? detectedRemarksCol
		: lastDayColumnIndex + 3;
	const presentCol = Number.isInteger(detectedPresentCol) ? detectedPresentCol : null;
	const absentCol = Number.isInteger(detectedAbsentCol) ? detectedAbsentCol : lastDayColumnIndex + 1;
	const excusedCol = Number.isInteger(detectedExcusedCol) ? detectedExcusedCol : null;
	const tardyCol = Number.isInteger(detectedTardyCol) ? detectedTardyCol : lastDayColumnIndex + 2;
	const dataStartRow = Math.max(14, headerRow + 2);
	const firstDataIndex = dataStartRow - 1;

	const columnsToClear = new Set([numberCol, lrnCol, nameCol]);
	if (presentCol != null) columnsToClear.add(presentCol);
	if (absentCol != null) columnsToClear.add(absentCol);
	if (excusedCol != null) columnsToClear.add(excusedCol);
	if (tardyCol != null) columnsToClear.add(tardyCol);
	if (remarksCol != null) columnsToClear.add(remarksCol);
	if (sexCol != null) columnsToClear.add(sexCol);
	dayColumnMap.forEach(col => columnsToClear.add(col));

	columnsToClear.forEach(col => {
		if (col == null) return;
		for (let row = firstDataIndex; row <= firstDataIndex + 150; row++) {
			const cell = worksheet.getCell(row + 1, col + 1);
			if (cell && !cell.formula) {
				cell.value = null;
			}
		}
	});

	const setCellValue = (rowIndex, colIndex, value, type = "s") => {
		if (colIndex == null || colIndex < 0) return;
		const cell = worksheet.getCell(rowIndex + 1, colIndex + 1);
		if (cell && cell.formula) {
			return;
		}
		if (value === null || value === undefined || value === "") {
			if (!cell.formula) {
				cell.value = null;
			}
			return;
		}
		if (type === "n") {
			const numeric = Number(value);
			if (Number.isFinite(numeric)) {
				cell.value = numeric;
				return;
			}
		}
		cell.value = value;
	};

	const writeAddress = (address, value) => {
		if (!address) return;
		const cell = worksheet.getCell(address);
		if (value === null || value === undefined || value === "") {
			if (!cell.formula) {
				cell.value = null;
			}
			return;
		}
		cell.value = value;
	};

	students.forEach((student, index) => {
		const rowIndex = firstDataIndex + index;
		setCellValue(rowIndex, numberCol, index + 1, "n");
		setCellValue(rowIndex, lrnCol, "");
		setCellValue(rowIndex, nameCol, student.name || "");
		if (firstDayColumnIndex != null) {
			for (let extraCol = numberCol + 1; extraCol < firstDayColumnIndex; extraCol++) {
				if (extraCol === nameCol || extraCol === lrnCol) continue;
				if (extraCol === sexCol) {
					setCellValue(rowIndex, extraCol, "");
					continue;
				}
				setCellValue(rowIndex, extraCol, "");
			}
		}
		if (sexCol != null) {
			setCellValue(rowIndex, sexCol, "");
		}

		const daily = Array.isArray(student.daily) ? student.daily : [];
		monthDays.forEach((day, dayIndex) => {
			const col = dayColumnMap.get(day.day);
			if (col == null) return;
			const entry = daily[dayIndex] || {};
			setCellValue(rowIndex, col, day.isWeekend ? "" : (entry.code || ""));
		});

		const totals = student.totals || {};
		if (presentCol != null) setCellValue(rowIndex, presentCol, totals.present != null ? totals.present : "", "n");
		if (absentCol != null) setCellValue(rowIndex, absentCol, totals.absent != null ? totals.absent : "", "n");
		if (excusedCol != null) setCellValue(rowIndex, excusedCol, totals.excused != null ? totals.excused : "", "n");
		if (tardyCol != null) setCellValue(rowIndex, tardyCol, totals.tardy != null ? totals.tardy : "", "n");
		if (remarksCol != null) setCellValue(rowIndex, remarksCol, student.remarks || "");
	});

	writeAddress("K6", report.schoolYear || "");
	writeAddress("X6", buildMonthLabel(report));
	writeAddress("C8", report.schoolName || SCHOOL_NAME);
	writeAddress("X8", report.gradeLevel || "");
	writeAddress("AC8", report.section || "");

	return workbook.xlsx.writeBuffer();
}

init();
