document.addEventListener('DOMContentLoaded', () => {
    // DOM Elements
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');
    const fileInfoCard = document.getElementById('file-info-card');
    const fileNameEl = document.getElementById('file-name');
    const fileDetailsEl = document.getElementById('file-details');
    const btnRemoveFile = document.getElementById('btn-remove-file');
    const btnProcess = document.getElementById('btn-process');
    const existingFileAlert = document.getElementById('existing-file-alert');
    const existingFileNameEl = document.getElementById('existing-file-name');
    const btnUseExisting = document.getElementById('btn-use-existing');
    const btnDownload = document.getElementById('btn-download');

    // Status & Metrics Elements
    const statusPill = document.getElementById('status-pill');
    const progressContainer = document.getElementById('progress-container');
    const progressBar = document.getElementById('progress-bar');
    const progressMessage = document.getElementById('progress-message');
    const progressPercent = document.getElementById('progress-percent');
    const processingAlert = document.getElementById('processing-alert');
    const processingAlertText = document.getElementById('processing-alert-text');

    const mTotalSheets = document.getElementById('metric-total-sheets');
    const mHistSheets = document.getElementById('metric-hist-sheets');
    const mHlSheets = document.getElementById('metric-hl-sheets');
    const mProcSheets = document.getElementById('metric-proc-sheets');
    const mSkipSheets = document.getElementById('metric-skip-sheets');
    const mTotalRecords = document.getElementById('metric-total-records');

    // Table & Pagination
    const recordsTableBody = document.getElementById('records-table-body');
    const recordsTableHeader = document.getElementById('records-table-header');
    const tabRecordCount = document.getElementById('tab-record-count');
    const tabLogCount = document.getElementById('tab-log-count');
    const auditTableBody = document.getElementById('audit-table-body');
    const paginationInfo = document.getElementById('pagination-info');
    const paginationControls = document.getElementById('pagination-controls');

    // Filters
    const filterUid = document.getElementById('filter-uid');
    const filterMeter = document.getElementById('filter-meter');
    const filterDate = document.getElementById('filter-date');
    const btnResetFilters = document.getElementById('btn-reset-filters');
    const filterIndicator = document.getElementById('filter-indicator');

    // State
    let currentUploadedFile = null;
    let currentWorkbookPath = null;
    let processedData = null;
    let allRecords = [];
    let filteredRecords = [];
    let currentPage = 1;
    const pageSize = 50;

    // Check for default workspace file on load
    checkDefaultFile();

    async function checkDefaultFile() {
        try {
            const res = await fetch('/api/default-file');
            const data = await res.json();
            if (data.available) {
                existingFileNameEl.textContent = data.filename;
                existingFileAlert.classList.remove('d-none');
                btnUseExisting.onclick = () => selectExistingFile(data.filepath, data.filename);
            }
        } catch (e) {
            console.warn('Could not check default file:', e);
        }
    }

    function selectExistingFile(filepath, filename) {
        currentWorkbookPath = filepath;
        currentUploadedFile = null;
        fileNameEl.textContent = filename;
        fileDetailsEl.textContent = 'Workspace file selected & validated';
        fileInfoCard.classList.remove('d-none');
        btnProcess.disabled = false;
        dropZone.classList.add('d-none');
        setStatus('Ready to Process', 'bg-primary-subtle text-primary');
    }

    // Drag and Drop Events
    ['dragenter', 'dragover'].forEach(eventName => {
        dropZone.addEventListener(eventName, (e) => {
            e.preventDefault();
            dropZone.classList.add('dragover');
        });
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, (e) => {
            e.preventDefault();
            dropZone.classList.remove('dragover');
        });
    });

    dropZone.addEventListener('drop', (e) => {
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            handleFileUpload(files[0]);
        }
    });

    fileInput.addEventListener('change', () => {
        if (fileInput.files.length > 0) {
            handleFileUpload(fileInput.files[0]);
        }
    });

    btnRemoveFile.addEventListener('click', () => {
        currentUploadedFile = null;
        currentWorkbookPath = null;
        fileInput.value = '';
        fileInfoCard.classList.add('d-none');
        dropZone.classList.remove('d-none');
        btnProcess.disabled = true;
        setStatus('Awaiting Workbook', 'bg-secondary-subtle text-secondary');
    });

    async function handleFileUpload(file) {
        const ext = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
        if (ext !== '.xlsx' && ext !== '.xlsm') {
            showError('Invalid file format. Please upload an Excel file (.xlsx or .xlsm).');
            return;
        }

        currentUploadedFile = file;
        fileNameEl.textContent = file.name;
        fileDetailsEl.textContent = `Uploading and validating (${(file.size / (1024 * 1024)).toFixed(2)} MB)...`;
        fileInfoCard.classList.remove('d-none');
        dropZone.classList.add('d-none');
        hideError();

        // Upload and validate with backend
        const formData = new FormData();
        formData.append('file', file);

        try {
            setStatus('Validating...', 'bg-warning-subtle text-warning');
            const res = await fetch('/api/upload', {
                method: 'POST',
                body: formData
            });

            const result = await res.json();
            if (!res.ok) {
                throw new Error(result.detail || 'Validation failed');
            }

            currentWorkbookPath = result.filepath;
            fileDetailsEl.textContent = `Master Sheet: '${result.master_sheet_name}', Total sheets: ${result.total_sheets}`;
            btnProcess.disabled = false;
            setStatus('Ready to Process', 'bg-primary-subtle text-primary');
        } catch (err) {
            showError(err.message);
            fileDetailsEl.textContent = 'Validation error occurred';
            setStatus('Validation Error', 'bg-danger-subtle text-danger');
            btnProcess.disabled = true;
        }
    }

    // Process Consolidation
    btnProcess.addEventListener('click', async () => {
        if (!currentWorkbookPath) return;

        btnProcess.disabled = true;
        hideError();
        showProgress(10, 'Initializing consolidation engine...');
        setStatus('Processing...', 'bg-primary text-white');

        try {
            showProgress(35, 'Extracting historical worksheets (23.02.2026 - 17.03.2026)...');

            const res = await fetch('/api/process', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ filepath: currentWorkbookPath })
            });

            showProgress(85, 'Resolving hyperlinks and assembling consolidated dataset...');

            const data = await res.json();
            if (!res.ok) {
                throw new Error(data.detail || 'Processing failed');
            }

            processedData = data;
            allRecords = data.records || [];
            filteredRecords = [...allRecords];

            // Update Metrics
            const summary = data.summary;
            mTotalSheets.textContent = summary.total_sheets_in_workbook;
            mHistSheets.textContent = summary.historical_sheets_count;
            mHlSheets.textContent = summary.hyperlink_sheets_count;
            mProcSheets.textContent = summary.worksheets_processed;
            mSkipSheets.textContent = summary.worksheets_skipped;
            mTotalRecords.textContent = summary.total_records_extracted.toLocaleString();

            // Update Tabs
            tabRecordCount.textContent = allRecords.length.toLocaleString();
            tabLogCount.textContent = (data.processing_log || []).length;

            // Configure Download
            btnDownload.href = `/api/download?filepath=${encodeURIComponent(data.output_file)}`;
            btnDownload.classList.remove('disabled');

            // Render Table & Audit Log
            renderHeaders(data.columns);
            currentPage = 1;
            applyFilters();
            renderAuditLog(data.processing_log);

            showProgress(100, 'Consolidation complete!');
            setTimeout(() => {
                progressContainer.classList.add('d-none');
            }, 1200);

            setStatus('Completed Successfully', 'bg-success text-white');
            btnProcess.disabled = false;
        } catch (err) {
            showError(err.message);
            setStatus('Error Occurred', 'bg-danger text-white');
            progressContainer.classList.add('d-none');
            btnProcess.disabled = false;
        }
    });

    // Rendering Headers
    function renderHeaders(columns) {
        recordsTableHeader.innerHTML = '';
        const exportCols = columns.filter(c => !c.startsWith('_'));
        exportCols.forEach(col => {
            const th = document.createElement('th');
            th.textContent = col;
            recordsTableHeader.appendChild(th);
        });
    }

    // Filter Logic
    function applyFilters() {
        const uidQuery = filterUid.value.trim().toLowerCase();
        const meterQuery = filterMeter.value.trim().toLowerCase();
        const dateQuery = filterDate.value.trim().toLowerCase();

        filteredRecords = allRecords.filter(r => {
            const matchUid = !uidQuery || (r['Panel UID'] && String(r['Panel UID']).toLowerCase().includes(uidQuery));
            const matchMeter = !meterQuery || (r['Meter Serial Number'] && String(r['Meter Serial Number']).toLowerCase().includes(meterQuery));
            const matchDate = !dateQuery || (r['Date'] && String(r['Date']).toLowerCase().includes(dateQuery));
            return matchUid && matchMeter && matchDate;
        });

        currentPage = 1;
        renderTable();
        renderPagination();

        if (uidQuery || meterQuery || dateQuery) {
            filterIndicator.textContent = `Filtered: ${filteredRecords.length} of ${allRecords.length}`;
        } else {
            filterIndicator.textContent = '';
        }
    }

    filterUid.addEventListener('input', applyFilters);
    filterMeter.addEventListener('input', applyFilters);
    filterDate.addEventListener('input', applyFilters);

    btnResetFilters.addEventListener('click', () => {
        filterUid.value = '';
        filterMeter.value = '';
        filterDate.value = '';
        applyFilters();
    });

    // Render Table Rows
    function renderTable() {
        recordsTableBody.innerHTML = '';

        if (filteredRecords.length === 0) {
            const tr = document.createElement('tr');
            tr.innerHTML = `<td colspan="15" class="text-center py-4 text-muted">No records match the current filter criteria.</td>`;
            recordsTableBody.appendChild(tr);
            paginationInfo.textContent = 'Showing 0 - 0 of 0 records';
            return;
        }

        const startIdx = (currentPage - 1) * pageSize;
        const endIdx = Math.min(startIdx + pageSize, filteredRecords.length);
        const pageRecords = filteredRecords.slice(startIdx, endIdx);

        const exportCols = (processedData.columns || []).filter(c => !c.startsWith('_'));

        pageRecords.forEach(record => {
            const tr = document.createElement('tr');
            exportCols.forEach(col => {
                const td = document.createElement('td');
                const val = record[col];
                if (col === 'Panel UID' || col === 'Meter Serial Number') {
                    td.className = 'font-monospace text-primary fw-medium';
                } else if (col === 'Date' || col === 'S No' || col === 'Phase') {
                    td.className = 'text-center';
                } else if (col === 'Location') {
                    td.innerHTML = `<span class="badge bg-primary-subtle text-primary">${val || 'EFFILUME'}</span>`;
                    tr.appendChild(td);
                    return;
                }
                td.textContent = val !== null && val !== undefined ? val : '';
                tr.appendChild(td);
            });
            recordsTableBody.appendChild(tr);
        });

        paginationInfo.textContent = `Showing ${(startIdx + 1).toLocaleString()} - ${endIdx.toLocaleString()} of ${filteredRecords.length.toLocaleString()} records`;
    }

    // Render Pagination Controls
    function renderPagination() {
        paginationControls.innerHTML = '';
        const totalPages = Math.ceil(filteredRecords.length / pageSize);
        if (totalPages <= 1) return;

        // Prev Button
        const prevLi = document.createElement('li');
        prevLi.className = `page-item ${currentPage === 1 ? 'disabled' : ''}`;
        prevLi.innerHTML = `<a class="page-link" href="#">&laquo;</a>`;
        prevLi.onclick = (e) => {
            e.preventDefault();
            if (currentPage > 1) {
                currentPage--;
                renderTable();
                renderPagination();
            }
        };
        paginationControls.appendChild(prevLi);

        // Page Numbers (sliding window of 5 pages)
        const startPage = Math.max(1, currentPage - 2);
        const endPage = Math.min(totalPages, startPage + 4);

        for (let p = startPage; p <= endPage; p++) {
            const li = document.createElement('li');
            li.className = `page-item ${p === currentPage ? 'active' : ''}`;
            li.innerHTML = `<a class="page-link" href="#">${p}</a>`;
            const pageNum = p;
            li.onclick = (e) => {
                e.preventDefault();
                currentPage = pageNum;
                renderTable();
                renderPagination();
            };
            paginationControls.appendChild(li);
        }

        // Next Button
        const nextLi = document.createElement('li');
        nextLi.className = `page-item ${currentPage === totalPages ? 'disabled' : ''}`;
        nextLi.innerHTML = `<a class="page-link" href="#">&raquo;</a>`;
        nextLi.onclick = (e) => {
            e.preventDefault();
            if (currentPage < totalPages) {
                currentPage++;
                renderTable();
                renderPagination();
            }
        };
        paginationControls.appendChild(nextLi);
    }

    // Render Audit Log
    function renderAuditLog(logEntries) {
        auditTableBody.innerHTML = '';
        if (!logEntries || logEntries.length === 0) {
            auditTableBody.innerHTML = `<tr><td colspan="7" class="text-center py-4 text-muted">No audit logs available.</td></tr>`;
            return;
        }

        logEntries.forEach(entry => {
            const tr = document.createElement('tr');
            let badgeClass = 'bg-secondary';
            if (entry.status === 'Processed') badgeClass = 'bg-success';
            else if (entry.status.includes('Error')) badgeClass = 'bg-danger';
            else if (entry.status.includes('Skipped')) badgeClass = 'bg-warning text-dark';

            tr.innerHTML = `
                <td class="font-monospace text-center">${entry.date_processed || 'N/A'}</td>
                <td class="fw-semibold">${entry.source_sheet || 'N/A'}</td>
                <td class="extra-small text-muted">${entry.source_type || 'Worksheet'}</td>
                <td class="text-center"><span class="badge bg-light text-dark border">${entry.location || 'EFFILUME'}</span></td>
                <td class="text-center fw-bold">${(entry.records_extracted || 0).toLocaleString()}</td>
                <td class="text-center"><span class="badge ${badgeClass}">${entry.status}</span></td>
                <td class="extra-small text-muted">${entry.reason || ''}</td>
            `;
            auditTableBody.appendChild(tr);
        });
    }

    // Helper UI functions
    function setStatus(text, badgeClass) {
        statusPill.textContent = text;
        statusPill.className = `badge ${badgeClass} small`;
    }

    function showProgress(pct, msg) {
        progressContainer.classList.remove('d-none');
        progressBar.style.width = `${pct}%`;
        progressPercent.textContent = `${pct}%`;
        progressMessage.textContent = msg;
    }

    function showError(msg) {
        processingAlertText.textContent = msg;
        processingAlert.classList.remove('d-none');
    }

    function hideError() {
        processingAlert.classList.add('d-none');
    }
});

