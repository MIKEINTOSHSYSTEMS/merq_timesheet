// Timesheet Management JavaScript

class TimesheetManager {
    constructor() {
        this.currentYear = null;
        this.currentMonth = null;
        this.projects = [];
        this.calendarData = [];
        this.timesheetData = {};
        this.monthDays = 0;
        this.initializeEventListeners();
    }

    initializeEventListeners() {

        // Test API method
        $('#testApiBtn').click(() => this.testApi());

        // Month/year selection
        $('#loadTimesheet').click(() => this.loadTimesheet());

        // Project management
        $('#addProjectBtn').click(() => this.showAddProjectModal());
        $('#saveProjectBtn').click(() => this.saveProject());

        // Action buttons
        $('#prefillBtn').click(() => this.prefillDefaultHours());
        $('#previewBtn').click(() => this.previewTimesheet());
        $('#exportBtn').click(() => this.exportTimesheet());
        $('#submitBtn').click(() => this.submitTimesheet());
        $('#clearBtn').click(() => this.clearAll());

        // Preview modal actions
        $('#exportFromPreview').click(() => this.exportTimesheet());
        $('#submitFromPreview').click(() => this.submitTimesheet());

        // Enter key in project modal
        $('#projectForm').on('keypress', (e) => {
            if (e.which === 13) {
                this.saveProject();
            }
        });
    }

    async loadTimesheet() {
        const year = $('#yearSelect').val();
        const month = $('#monthSelect').val();

        if (!year || !month) {
            this.showAlert('Please select both year and month', 'warning');
            return;
        }

        this.currentYear = parseInt(year);
        this.currentMonth = parseInt(month);

        this.showLoading('Loading timesheet...');

        try {
            const response = await fetch('/api/timesheet/data', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({ year: this.currentYear, month: this.currentMonth })
            });

            const data = await response.json();

            if (response.ok) {
                this.calendarData = data.calendar;
                this.timesheetData = data.timesheet_data;
                this.projects = data.projects;
                this.monthDays = data.month_days;

                this.renderProjectsList();
                this.renderTimesheet();
                this.enableActionButtons();

                this.showAlert('Timesheet loaded successfully', 'success');
            } else {
                this.showAlert(data.error || 'Failed to load timesheet', 'error');
            }
        } catch (error) {
            console.error('Error loading timesheet:', error);
            this.showAlert('Network error occurred', 'error');
        } finally {
            this.hideLoading();
        }
    }

    renderTimesheet() {
        this.renderHeader();
        this.renderBody();
        this.showTimesheetTable();
    }

    renderHeader() {
        const header = $('#timesheetHeader');
        header.empty();

        // Create header rows
        const row1 = $('<tr>').addClass('table-primary');
        const row2 = $('<tr>').addClass('table-primary');

        // First three columns (fixed)
        row1.append('<th rowspan="2">Employee/Project</th>');
        row1.append('<th rowspan="2">Type</th>');
        row1.append('<th rowspan="2">Allocated Hours</th>');
        row1.append('<th colspan="' + this.calendarData.length + '">Day Number</th>');
        row1.append('<th rowspan="2">Total Hours</th>');

        //row2.append('<th>Date</th>');

        // Day headers
        this.calendarData.forEach(day => {
            const th = $('<th>').addClass('day-cell text-center small');
            if (day.is_weekend) {
                th.addClass('weekend');
            }
            th.text(day.day);
            row2.append(th);
        });

        //row2.append('<th>Hours</th>');

        header.append(row1, row2);
    }

    renderBody() {
        const body = $('#timesheetBody');
        body.empty();

        // Employee name row
        const employeeRow = $('<tr>').addClass('table-info');
        employeeRow.append('<td colspan="3"><strong>Employee Name</strong></td>');
        employeeRow.append('<td colspan="' + (this.calendarData.length + 1) + '">' +
            (window.userData ? window.userData.full_name : 'Current User') + '</td>');
        body.append(employeeRow);

        // Date row
        const dateRow = $('<tr>');
        dateRow.append('<td colspan="3" class="table-light"><strong>Date</strong></td>');
        this.calendarData.forEach(day => {
            const td = $('<td>').addClass('text-center small');
            if (day.is_weekend) {
                td.addClass('weekend');
            }
            td.text(day.date);
            dateRow.append(td);
        });
        dateRow.append('<td class="table-light"></td>');
        body.append(dateRow);

        // Weekday row
        const weekdayRow = $('<tr>');
        weekdayRow.append('<td colspan="3" class="table-light"><strong>Weekday</strong></td>');
        this.calendarData.forEach(day => {
            const td = $('<td>').addClass('text-center small fw-bold');
            if (day.is_weekend) {
                td.addClass('weekend');
            }
            td.text(day.weekday_amharic);
            weekdayRow.append(td);
        });
        weekdayRow.append('<td class="table-light"></td>');
        body.append(weekdayRow);

        // Project rows
        this.projects.forEach((project, index) => {
            this.renderProjectRow(project, index);
        });

        // Leave rows
        this.renderLeaveRows();

        // Total rows
        this.renderTotalRows();
    }

    renderProjectRow(project, projectIndex) {
        const row = $('<tr>').addClass('project-row');

        // Project name
        row.append($('<td>').addClass('fw-bold').text(project.name));

        // Type label
        row.append($('<td>').addClass('small').text('Direct Work'));

        // Allocated hours
        row.append($('<td>').addClass('text-center fw-bold').text(project.allocated_hours));

        // Hour inputs
        let projectTotal = 0;
        this.calendarData.forEach(day => {
            const td = $('<td>').addClass('text-center');
            if (day.is_weekend) {
                td.addClass('weekend');
            }

            const currentValue = project.hours && project.hours[day.day] ? project.hours[day.day] : '';
            projectTotal += parseFloat(currentValue) || 0;

            const input = $('<input>')
                .attr({
                    type: 'number',
                    min: '0',
                    max: '24',
                    step: '0.5',
                    class: 'form-control form-control-sm hours-input'
                })
                .data('project-id', project.id)
                .data('day', day.day)
                .val(currentValue);

            input.on('input', (e) => this.handleHoursInput(e.target, 'project', project.id, day.day));

            td.append(input);
            row.append(td);
        });

        // Total cell
        row.append($('<td>').addClass('text-center fw-bold total-cell')
            .data('type', 'project-total')
            .data('project-id', project.id)
            .text(projectTotal.toFixed(1)));

        $('#timesheetBody').append(row);
    }

    renderLeaveRows() {
        const leaveTypes = [
            { key: 'vacation', name: 'Vacation' },
            { key: 'sick_leave', name: 'Sick Leave' },
            { key: 'holiday', name: 'Holiday' },
            { key: 'personal_leave', name: 'Personal Leave' },
            { key: 'bereavement', name: 'Bereavement' },
            { key: 'other', name: 'Other' }
        ];

        leaveTypes.forEach(leaveType => {
            const row = $('<tr>').addClass('leave-row');

            row.append($('<td>').addClass('fw-bold').text('Leave'));
            row.append($('<td>').addClass('small').text(leaveType.name));
            row.append($('<td>').addClass('text-center').text('-'));

            let leaveTotal = 0;
            this.calendarData.forEach(day => {
                const td = $('<td>').addClass('text-center');
                if (day.is_weekend) {
                    td.addClass('weekend');
                }

                const currentValue = this.timesheetData.leave_entries[leaveType.key][day.day] || 0;
                leaveTotal += parseFloat(currentValue) || 0;

                const input = $('<input>')
                    .attr({
                        type: 'number',
                        min: '0',
                        max: '24',
                        step: '0.5',
                        class: 'form-control form-control-sm hours-input'
                    })
                    .data('leave-type', leaveType.key)
                    .data('day', day.day)
                    .val(currentValue);

                input.on('input', (e) => this.handleHoursInput(e.target, 'leave', leaveType.key, day.day));

                td.append(input);
                row.append(td);
            });

            // Total cell
            row.append($('<td>').addClass('text-center fw-bold total-cell')
                .data('type', 'leave-total')
                .data('leave-type', leaveType.key)
                .text(leaveTotal.toFixed(1)));

            $('#timesheetBody').append(row);
        });
    }

    renderTotalRows() {
        // Direct work total
        const directTotalRow = $('<tr>').addClass('total-row');
        directTotalRow.append('<td colspan="3" class="fw-bold">Total Direct Work</td>');

        let directGrandTotal = 0;
        this.calendarData.forEach(day => {
            const dailyTotal = this.timesheetData.daily_totals[day.day] || 0;
            directGrandTotal += dailyTotal;

            const td = $('<td>').addClass('text-center fw-bold')
                .data('day', day.day)
                .data('type', 'daily-total')
                .text(dailyTotal.toFixed(1));

            if (day.is_weekend) {
                td.addClass('weekend');
            }
            directTotalRow.append(td);
        });

        directTotalRow.append($('<td>').addClass('text-center fw-bold')
            .data('type', 'direct-grand-total')
            .text(directGrandTotal.toFixed(1)));

        $('#timesheetBody').append(directTotalRow);

        // Leave total
        const leaveTotalRow = $('<tr>').addClass('total-row');
        leaveTotalRow.append('<td colspan="3" class="fw-bold">Total Leave</td>');

        let leaveGrandTotal = 0;
        this.calendarData.forEach(day => {
            const dailyTotal = this.timesheetData.leave_totals[day.day] || 0;
            leaveGrandTotal += dailyTotal;

            const td = $('<td>').addClass('text-center fw-bold')
                .data('day', day.day)
                .data('type', 'leave-total')
                .text(dailyTotal.toFixed(1));

            if (day.is_weekend) {
                td.addClass('weekend');
            }
            leaveTotalRow.append(td);
        });

        leaveTotalRow.append($('<td>').addClass('text-center fw-bold')
            .data('type', 'leave-grand-total')
            .text(leaveGrandTotal.toFixed(1)));

        $('#timesheetBody').append(leaveTotalRow);

        // Grand total
        const grandTotalRow = $('<tr>').addClass('total-row bg-dark');
        grandTotalRow.append('<td colspan="3" class="fw-bold">Grand Total</td>');

        let grandTotal = 0;
        this.calendarData.forEach(day => {
            const dailyTotal = this.timesheetData.grand_totals[day.day] || 0;
            grandTotal += dailyTotal;

            const td = $('<td>').addClass('text-center fw-bold')
                .data('day', day.day)
                .data('type', 'grand-total')
                .text(dailyTotal.toFixed(1));

            if (day.is_weekend) {
                td.addClass('weekend');
            }
            grandTotalRow.append(td);
        });

        grandTotalRow.append($('<td>').addClass('text-center fw-bold')
            .data('type', 'grand-grand-total')
            .text(grandTotal.toFixed(1)));

        $('#timesheetBody').append(grandTotalRow);
    }

    async handleHoursInput(input, type, identifier, day) {
        const value = parseFloat(input.value) || 0;

        // Validate hours (0-24)
        if (value < 0 || value > 24) {
            this.showAlert('Hours must be between 0 and 24', 'warning');
            input.value = '';
            return;
        }

        // Prepare data for saving
        const saveData = {
            year: this.currentYear,
            month: this.currentMonth
        };

        if (type === 'project') {
            saveData.project_hours = {
                [identifier]: {
                    [day]: value
                }
            };
        } else {
            saveData.leave_hours = {
                [identifier]: {
                    [day]: value
                }
            };
        }

        try {
            const response = await fetch('/api/timesheet/save', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify(saveData)
            });

            const result = await response.json();

            if (response.ok) {
                // Update local data
                if (type === 'project') {
                    const project = this.projects.find(p => p.id == identifier);
                    if (project) {
                        if (!project.hours) project.hours = {};
                        project.hours[day] = value;
                    }
                } else {
                    this.timesheetData.leave_entries[identifier][day] = value;
                }

                // Update totals from server response
                if (result.totals) {
                    this.timesheetData.daily_totals = result.totals.daily_totals;
                    this.timesheetData.leave_totals = result.totals.leave_totals;
                    this.timesheetData.grand_totals = result.totals.grand_totals;
                    this.updateTotalsDisplay();
                }
            } else {
                this.showAlert('Failed to save hours', 'error');
            }
        } catch (error) {
            console.error('Error saving hours:', error);
            this.showAlert('Network error occurred', 'error');
        }
    }

    updateTotalsDisplay() {
        // Update project totals
        this.projects.forEach(project => {
            const projectTotal = Object.values(project.hours || {}).reduce((sum, hours) => sum + (parseFloat(hours) || 0), 0);
            $(`.total-cell[data-type="project-total"][data-project-id="${project.id}"]`).text(projectTotal.toFixed(1));
        });

        // Update leave totals
        const leaveTypes = ['vacation', 'sick_leave', 'holiday', 'personal_leave', 'bereavement', 'other'];
        leaveTypes.forEach(leaveType => {
            const leaveTotal = Object.values(this.timesheetData.leave_entries[leaveType] || {}).reduce((sum, hours) => sum + (parseFloat(hours) || 0), 0);
            $(`.total-cell[data-type="leave-total"][data-leave-type="${leaveType}"]`).text(leaveTotal.toFixed(1));
        });

        // Update daily totals
        this.calendarData.forEach(day => {
            $(`.total-cell[data-type="daily-total"][data-day="${day.day}"]`).text((this.timesheetData.daily_totals[day.day] || 0).toFixed(1));
            $(`.total-cell[data-type="leave-total"][data-day="${day.day}"]`).text((this.timesheetData.leave_totals[day.day] || 0).toFixed(1));
            $(`.total-cell[data-type="grand-total"][data-day="${day.day}"]`).text((this.timesheetData.grand_totals[day.day] || 0).toFixed(1));
        });

        // Update grand totals
        const directGrandTotal = Object.values(this.timesheetData.daily_totals).reduce((sum, hours) => sum + (parseFloat(hours) || 0), 0);
        const leaveGrandTotal = Object.values(this.timesheetData.leave_totals).reduce((sum, hours) => sum + (parseFloat(hours) || 0), 0);
        const grandTotal = directGrandTotal + leaveGrandTotal;

        $('.total-cell[data-type="direct-grand-total"]').text(directGrandTotal.toFixed(1));
        $('.total-cell[data-type="leave-grand-total"]').text(leaveGrandTotal.toFixed(1));
        $('.total-cell[data-type="grand-grand-total"]').text(grandTotal.toFixed(1));
    }

    async loadProjects() {
        try {
            const response = await fetch(`/api/projects?year=${this.currentYear}&month=${this.currentMonth}`);
            const data = await response.json();

            if (response.ok) {
                this.projects = data.projects;
                this.renderProjectsList();
            }
        } catch (error) {
            console.error('Error loading projects:', error);
        }
    }

    renderProjectsList() {
        const container = $('#projectsContainer');
        container.empty();

        if (this.projects.length === 0) {
            container.html(`
                <div class="alert alert-info">
                    <i class="bi bi-info-circle me-2"></i>
                    No projects added yet. Click "Add Project" to get started.
                </div>
            `);
            return;
        }

        this.projects.forEach((project, index) => {
            const progressPercent = project.allocated_hours > 0 ? (project.total_hours / project.allocated_hours * 100) : 0;
            const progressColor = progressPercent > 100 ? 'bg-danger' : progressPercent >= 80 ? 'bg-warning' : 'bg-success';

            const projectCard = `
                <div class="card mb-2">
                    <div class="card-body py-2">
                        <div class="d-flex justify-content-between align-items-center">
                            <div>
                                <h6 class="mb-1">${project.name}</h6>
                                <small class="text-muted">Allocated: ${project.allocated_hours} hours</small>
                                <div class="progress mt-1" style="height: 5px; width: 150px;">
                                    <div class="progress-bar ${progressColor}" style="width: ${Math.min(progressPercent, 100)}%"></div>
                                </div>
                            </div>
                            <div class="d-flex gap-2 align-items-center">
                                <span class="badge bg-success">${project.total_hours || 0} hours worked</span>
                                ${project.id !== 1 ? `
                                <button class="btn btn-sm btn-outline-danger delete-project" data-project-id="${project.id}">
                                    <i class="bi bi-trash"></i>
                                </button>
                                ` : ''}
                            </div>
                        </div>
                    </div>
                </div>
            `;
            container.append(projectCard);
        });

        // Add delete event listeners
        $('.delete-project').click((e) => {
            const projectId = $(e.target).closest('.delete-project').data('project-id');
            this.deleteProject(projectId);
        });
    }

    showAddProjectModal() {
        $('#projectName').val('');
        $('#allocatedHours').val('0.0');
        $('#addProjectModal').modal('show');
    }

    async saveProject() {
        const name = $('#projectName').val().trim();
        const allocatedHours = parseFloat($('#allocatedHours').val());

        if (!name) {
            this.showAlert('Please enter project name', 'warning');
            return;
        }

        if (isNaN(allocatedHours) || allocatedHours < 0) {
            this.showAlert('Please enter valid allocated hours', 'warning');
            return;
        }

        try {
            const response = await fetch('/api/projects', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    name: name,
                    allocated_hours: allocatedHours,
                    year: this.currentYear,
                    month: this.currentMonth
                })
            });

            const data = await response.json();

            if (response.ok) {
                this.projects.push(data.project);
                this.renderProjectsList();
                this.renderTimesheet();
                $('#addProjectModal').modal('hide');
                this.showAlert('Project added successfully', 'success');
            } else {
                this.showAlert(data.error || 'Failed to add project', 'error');
            }
        } catch (error) {
            console.error('Error saving project:', error);
            this.showAlert('Network error occurred', 'error');
        }
    }

    async deleteProject(projectId) {
        if (!confirm('Are you sure you want to delete this project?')) {
            return;
        }

        try {
            const response = await fetch(`/api/projects?project_id=${projectId}&year=${this.currentYear}&month=${this.currentMonth}`, {
                method: 'DELETE'
            });

            if (response.ok) {
                this.projects = this.projects.filter(p => p.id !== projectId);
                this.renderProjectsList();
                this.renderTimesheet();
                this.showAlert('Project deleted successfully', 'success');
            } else {
                this.showAlert('Failed to delete project', 'error');
            }
        } catch (error) {
            console.error('Error deleting project:', error);
            this.showAlert('Network error occurred', 'error');
        }
    }

    async prefillDefaultHours() {
        if (!confirm('This will prefill default hours based on weekdays. Continue?')) {
            return;
        }

        this.showLoading('Prefilling default hours...');

        try {
            const response = await fetch('/api/timesheet/prefill', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    year: this.currentYear,
                    month: this.currentMonth
                })
            });

            const data = await response.json();

            if (response.ok) {
                // Reload the timesheet to reflect changes
                await this.loadTimesheet();
                this.showAlert('Default hours prefilled successfully', 'success');
            } else {
                this.showAlert(data.error || 'Failed to prefill hours', 'error');
            }
        } catch (error) {
            console.error('Error prefilling hours:', error);
            this.showAlert('Network error occurred', 'error');
        } finally {
            this.hideLoading();
        }
    }

    async previewTimesheet() {
        console.log('Starting preview generation...');

        if (!this.currentYear || !this.currentMonth) {
            this.showAlert('Please load a timesheet first', 'warning');
            return;
        }

        this.showLoading('Generating preview...');

        try {
            console.log('Sending preview request:', {
                year: this.currentYear,
                month: this.currentMonth
            });

            const response = await fetch('/api/timesheet/preview', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    year: this.currentYear,
                    month: this.currentMonth
                })
            });

            console.log('Response status:', response.status);
            console.log('Response ok:', response.ok);

            if (!response.ok) {
                const errorText = await response.text();
                console.error('Server response error:', errorText);
                throw new Error(`Server returned ${response.status}: ${errorText}`);
            }

            const data = await response.json();
            console.log('Preview data received:', data);

            if (data.success) {
                this.showPreview(data.preview_data);
            } else {
                this.showAlert(data.error || 'Failed to generate preview', 'error');
            }
        } catch (error) {
            console.error('Error generating preview:', error);
            this.showAlert(`Preview failed: ${error.message}`, 'error');
        } finally {
            this.hideLoading();
        }
    }

    async testApi() {
        try {
            console.log('Testing API connection...');

            const response = await fetch('/api/debug');
            const data = await response.json();

            console.log('API test response:', data);
            this.showAlert('API connection successful! Check console for details.', 'success');
        } catch (error) {
            console.error('API test failed:', error);
            this.showAlert(`API test failed: ${error.message}`, 'error');
        }
    }

    showPreview(previewData) {
        const previewContent = $('#previewContent');

        let html = `
            <div class="preview-content">
                <h4 class="text-center mb-4">MERQ CONSULTANCY</h4>
                <p class="text-center"><strong>ወርሃዊ የስራ ሰዓት መከታተያ / Monthly Timesheet Tracker</strong></p>
                <hr>
                
                <div class="row mb-4">
                    <div class="col-md-6">
                        <p><strong>ሰራተኛ/አማካሪ ስም / Employee/Consultant Name:</strong> ${window.userData.full_name}</p>
                    </div>
                    <div class="col-md-6">
                        <p><strong>ወር / Month:</strong> ${previewData.month_name} ${previewData.year}</p>
                    </div>
                </div>

                <h5>ፕሮጀክቶች / Projects Summary:</h5>
                <div class="table-responsive">
                    <table class="table table-bordered">
                        <thead class="table-primary">
                            <tr>
                                <th>Project Name</th>
                                <th>Total Hours</th>
                                <th>Allocated Hours</th>
                                <th>Equivalent Days</th>
                                <th>% of Direct</th>
                                <th>% of Total</th>
                            </tr>
                        </thead>
                        <tbody>
        `;

        previewData.project_totals.forEach(project => {
            html += `
                <tr>
                    <td>${project.name}</td>
                    <td>${project.total_hours.toFixed(1)}</td>
                    <td>${project.allocated_hours.toFixed(1)}</td>
                    <td>${project.equiv_days.toFixed(1)}</td>
                    <td>${project.percent_direct.toFixed(1)}%</td>
                    <td>${project.percent_total.toFixed(1)}%</td>
                </tr>
            `;
        });

        html += `
                        </tbody>
                    </table>
                </div>

                <h5 class="mt-4">ጠቅላላ ሰዓቶች ማጠቃለያ / Total Hours Summary:</h5>
                <div class="row">
                    <div class="col-md-6">
                        <table class="table table-bordered">
                            <tr>
                                <th>Total Work Hours:</th>
                                <td class="fw-bold">${previewData.total_work_hours.toFixed(1)}</td>
                            </tr>
                            <tr>
                                <th>Total Leave Hours:</th>
                                <td class="fw-bold">${previewData.total_leave_hours.toFixed(1)}</td>
                            </tr>
                            <tr>
                                <th>Grand Total:</th>
                                <td class="fw-bold text-success">${previewData.grand_total.toFixed(1)}</td>
                            </tr>
                        </table>
                    </div>
                </div>

                <div class="alert alert-info mt-4">
                    <h6>Declaration:</h6>
                    <p class="mb-2">
                        እኔ፣ ከዚህ በላይ ያለው መረጃ እውነት መሆኑን፣ ከእውነታው በኋላ የሚወሰነው እና በእኔ በተከናወነው ትክክለኛ ስራ ላይ የተመሰረተ መሆኑን እገልጻለሁ።
                    </p>
                    <p class="mb-0">
                        I, hereby declare that the foregoing information is true, is determined after the fact and is based on actual work performed by me.
                    </p>
                </div>
            </div>
        `;

        previewContent.html(html);
        $('#previewModal').modal('show');
    }

    exportTimesheet() {
        window.location.href = `/timesheet/export?year=${this.currentYear}&month=${this.currentMonth}`;
    }

    async submitTimesheet() {
        if (!confirm('Are you sure you want to submit this timesheet to HR? This action cannot be undone.')) {
            return;
        }

        this.showLoading('Submitting timesheet...');

        try {
            const response = await fetch('/timesheet/submit', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    year: this.currentYear,
                    month: this.currentMonth
                })
            });

            const data = await response.json();

            if (response.ok) {
                this.showAlert('Timesheet submitted to HR successfully', 'success');
                $('#previewModal').modal('hide');
            } else {
                this.showAlert(data.error || 'Failed to submit timesheet', 'error');
            }
        } catch (error) {
            console.error('Error submitting timesheet:', error);
            this.showAlert('Network error occurred', 'error');
        } finally {
            this.hideLoading();
        }
    }

    async clearAll() {
        if (!confirm('Are you sure you want to clear all data? This cannot be undone.')) {
            return;
        }

        this.showLoading('Clearing timesheet...');

        try {
            const response = await fetch('/timesheet/clear', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    year: this.currentYear,
                    month: this.currentMonth
                })
            });

            const data = await response.json();

            if (response.ok) {
                // Reload the timesheet to reflect changes
                await this.loadTimesheet();
                this.showAlert('All data cleared successfully', 'success');
            } else {
                this.showAlert(data.error || 'Failed to clear timesheet', 'error');
            }
        } catch (error) {
            console.error('Error clearing timesheet:', error);
            this.showAlert('Network error occurred', 'error');
        } finally {
            this.hideLoading();
        }
    }

    showTimesheetTable() {
        $('#timesheetEmpty').hide();
        $('#timesheetTableContainer').show();
    }

    enableActionButtons() {
        $('#prefillBtn').prop('disabled', false);
        $('#previewBtn').prop('disabled', false);
        $('#exportBtn').prop('disabled', false);
        $('#submitBtn').prop('disabled', false);
        $('#clearBtn').prop('disabled', false);
    }

    showAlert(message, type) {
        // Remove existing alerts
        $('.alert-dismissible').remove();

        const alertClass = {
            'success': 'alert-success',
            'error': 'alert-danger',
            'warning': 'alert-warning',
            'info': 'alert-info'
        }[type] || 'alert-info';

        const alertHtml = `
            <div class="alert ${alertClass} alert-dismissible fade show" role="alert">
                ${message}
                <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
            </div>
        `;

        $('.container.mt-3').prepend(alertHtml);
    }

    showLoading(message) {
        $('#timesheetLoading').show();
        $('#loadTimesheet').prop('disabled', true).html('<span class="loading-spinner me-2"></span>' + message);
    }

    hideLoading() {
        $('#timesheetLoading').hide();
        $('#loadTimesheet').prop('disabled', false).html('<i class="bi bi-arrow-clockwise me-2"></i>Load Timesheet');
    }
}

// Initialize when document is ready
$(document).ready(function () {
    // Set user data from Flask template
    try {
        window.userData = JSON.parse('{{ user.user_data | tojson | safe }}');
    } catch (e) {
        console.error('Error parsing user data:', e);
    }

    window.timesheetManager = new TimesheetManager();
});


// Global error handler
window.addEventListener('error', function (e) {
    console.error('Global error:', e.error);
    console.error('Error details:', {
        message: e.message,
        filename: e.filename,
        lineno: e.lineno,
        colno: e.colno
    });
});

// Promise rejection handler
window.addEventListener('unhandledrejection', function (e) {
    console.error('Unhandled promise rejection:', e.reason);
});