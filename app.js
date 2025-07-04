class SeatingApp {
    constructor() {
        this.students = [];
        this.seats = [];
        this.rows = 8;
        this.cols = 6;
        this.selectedStudent = null;
        this.selectedSeat = null;
        this.history = [];
        this.historyIndex = -1;
        this.constraints = [];
        this.showCoordinates = true;
        
        this.init();
    }

    init() {
        this.loadData();
        this.setupEventListeners();
        this.renderClassroom();
        this.renderStudentList();
        this.updateStats();
        this.applyCurrentFilter();
        this.updateHistoryButtons();
    }

    loadData() {
        const savedData = localStorage.getItem('seatingData');
        if (savedData) {
            const data = JSON.parse(savedData);
            this.students = data.students || [];
            this.rows = data.rows || 8;
            this.cols = data.cols || 6;
            this.constraints = data.constraints || [];
            this.showCoordinates = data.showCoordinates !== undefined ? data.showCoordinates : true;
            // 恢复历史记录
            this.history = data.history || [];
            this.historyIndex = data.historyIndex !== undefined ? data.historyIndex : -1;
        }
        this.initializeSeats();
        
        // 在加载完数据后，如果没有历史记录，添加初始状态
        if (this.history.length === 0) {
            this.addToHistory('seatArrangement', { seats: this.seats });
            this.historyIndex = -1; // 重置为初始状态，因为这是起始点
        }
    }

    saveData() {
        const data = {
            students: this.students,
            seats: this.seats,
            rows: this.rows,
            cols: this.cols,
            constraints: this.constraints,
            showCoordinates: this.showCoordinates,
            history: this.history,
            historyIndex: this.historyIndex
        };
        localStorage.setItem('seatingData', JSON.stringify(data));
    }

    initializeSeats() {
        this.seats = [];
        // 教室坐标系统说明:
        // - row: 0为第一排(靠近讲台), 数值增加表示远离讲台
        // - col: 0为第一列(最左边), 数值增加表示向右移动
        // - 这符合传统教室布局: 第一排第一列在左前方
        for (let row = 0; row < this.rows; row++) {
            for (let col = 0; col < this.cols; col++) {
                this.seats.push({
                    id: `${row}-${col}`,
                    row: row,    // 内部行索引: 0-based, 0=第一排(靠近讲台)
                    col: col,    // 内部列索引: 0-based, 0=第一列(最左边)
                    student: null,
                    position: row * this.cols + col + 1  // 线性位置编号
                });
            }
        }
    }

    setupEventListeners() {
        document.getElementById('addStudent').addEventListener('click', () => {
            this.showStudentModal();
        });

        document.getElementById('randomSeat').addEventListener('click', () => {
            this.randomSeatArrangement();
        });

        document.getElementById('clearSeats').addEventListener('click', () => {
            this.clearAllSeats();
        });

        document.getElementById('saveLayout').addEventListener('click', () => {
            this.saveCurrentLayout();
        });

        document.getElementById('exportLayout').addEventListener('click', () => {
            this.exportLayout();
        });

        // 撤销按钮在主页工具栏中，直接绑定
        document.getElementById('undoBtn').addEventListener('click', () => {
            this.undo();
        });

        // 使用事件委托处理模态框中的按钮
        document.addEventListener('click', (e) => {
            switch (e.target.id) {
                case 'applyLayout':
                    this.applyNewLayout();
                    break;
                case 'addConstraint':
                    this.addConstraint();
                    break;
            }
        });

        document.getElementById('saveStudent').addEventListener('click', () => {
            this.saveStudent();
        });

        document.getElementById('cancelStudent').addEventListener('click', () => {
            this.hideStudentModal();
        });

        document.querySelector('.modal-close').addEventListener('click', () => {
            this.hideStudentModal();
        });

        document.getElementById('studentModal').addEventListener('click', (e) => {
            if (e.target.id === 'studentModal') {
                this.hideStudentModal();
            }
        });

        document.getElementById('searchStudent').addEventListener('input', (e) => {
            this.filterStudents(e.target.value);
        });

        document.getElementById('filterStudents').addEventListener('change', (e) => {
            this.filterStudentsByStatus(e.target.value);
        });

        document.getElementById('clearAllStudents').addEventListener('click', () => {
            this.clearAllStudents();
        });

        document.getElementById('seatingSettingsBtn').addEventListener('click', () => {
            this.showSeatingSettingsModal();
        });

        document.getElementById('closeSeatingSettings').addEventListener('click', () => {
            this.hideSeatingSettingsModal();
        });

        document.getElementById('showCoordinatesMain').addEventListener('change', (e) => {
            this.toggleCoordinatesDisplay(e.target.checked);
        });

        // 添加约束输入框的回车键支持
        document.getElementById('constraintInput').addEventListener('keypress', (e) => {
            if (e.key === 'Enter') {
                this.addConstraint();
            }
        });

        document.getElementById('constraintInput').addEventListener('keypress', (e) => {
            if (e.key === 'Enter') {
                this.addConstraint();
            }
        });

        document.getElementById('seatingSettingsModal').addEventListener('click', (e) => {
            if (e.target.id === 'seatingSettingsModal') {
                this.hideSeatingSettingsModal();
            }
        });

        this.setupExcelEventListeners();
    }

    setupExcelEventListeners() {
        document.getElementById('importExcel').addEventListener('click', () => {
            this.importExcelFile();
        });

        document.getElementById('downloadTemplate').addEventListener('click', () => {
            this.downloadExcelTemplate();
        });

        document.getElementById('excelFileInput').addEventListener('change', (e) => {
            this.handleExcelFileSelect(e);
        });

        document.getElementById('cancelImport').addEventListener('click', () => {
            this.hideExcelPreviewModal();
        });

        document.getElementById('confirmImport').addEventListener('click', () => {
            this.confirmExcelImport();
        });

        document.getElementById('excelPreviewModal').addEventListener('click', (e) => {
            if (e.target.id === 'excelPreviewModal') {
                this.hideExcelPreviewModal();
            }
        });

        document.querySelectorAll('.tab-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                this.switchTab(e.target.dataset.tab);
            });
        });
    }

    addToHistory(action, data) {
        this.history = this.history.slice(0, this.historyIndex + 1);
        this.history.push({
            action: action,
            data: JSON.parse(JSON.stringify(data)),
            timestamp: Date.now()
        });
        this.historyIndex++;
        console.log(`历史记录已添加: ${action}, 当前索引: ${this.historyIndex}, 总数: ${this.history.length}`);
        this.updateHistoryButtons();
    }

    undo() {
        if (this.historyIndex >= 0) {
            const historyItem = this.history[this.historyIndex];
            console.log(`撤销操作: 索引 ${this.historyIndex}, 操作类型: ${historyItem.action}`);
            this.restoreFromHistory(historyItem);
            this.historyIndex--;
        } else {
            console.log('无法撤销: 已经是最早的状态');
        }
    }

    redo() {
        if (this.historyIndex < this.history.length - 1) {
            this.historyIndex++;
            const historyItem = this.history[this.historyIndex];
            console.log(`重做操作: 索引 ${this.historyIndex}, 操作类型: ${historyItem.action}`);
            this.restoreFromHistory(historyItem);
        } else {
            console.log('无法重做: 已经是最新的状态');
        }
    }

    restoreFromHistory(historyItem) {
        switch (historyItem.action) {
            case 'seatArrangement':
                this.seats = JSON.parse(JSON.stringify(historyItem.data.seats));
                this.renderClassroom();
                this.renderStudentList();
                this.updateStats();
                break;
        }
        this.updateHistoryButtons();
    }

    updateHistoryButtons() {
        const undoBtn = document.getElementById('undoBtn');
        
        const canUndo = this.historyIndex >= 0;
        
        if (undoBtn) {
            undoBtn.disabled = !canUndo;
            console.log(`撤销按钮状态: ${canUndo ? '启用' : '禁用'}`);
        }
        
        console.log(`历史状态: 索引 ${this.historyIndex}, 总数 ${this.history.length}`);
    }

    showStudentModal(student = null) {
        const modal = document.getElementById('studentModal');
        const form = document.getElementById('studentForm');
        const title = document.getElementById('modalTitle');
        
        if (student) {
            title.textContent = '编辑学生';
            document.getElementById('studentName').value = student.name;
            document.getElementById('studentId').value = student.id || '';
            document.getElementById('studentHeight').value = student.height || '';
            document.getElementById('studentGender').value = student.gender || '';
            document.getElementById('studentVision').checked = student.needsFrontSeat || false;
            document.getElementById('studentNotes').value = student.notes || '';
            form.dataset.editId = student.uuid;
        } else {
            title.textContent = '添加学生';
            form.reset();
            delete form.dataset.editId;
        }
        
        modal.style.display = 'flex';
    }

    hideStudentModal() {
        document.getElementById('studentModal').style.display = 'none';
    }

    saveStudent() {
        const form = document.getElementById('studentForm');
        const name = document.getElementById('studentName').value.trim();
        
        if (!name) {
            alert('请输入学生姓名');
            return;
        }

        const student = {
            uuid: form.dataset.editId || this.generateUUID(),
            name: name,
            id: document.getElementById('studentId').value.trim(),
            height: parseInt(document.getElementById('studentHeight').value) || null,
            gender: document.getElementById('studentGender').value,
            needsFrontSeat: document.getElementById('studentVision').checked,
            notes: document.getElementById('studentNotes').value.trim(),
            seatId: null
        };

        if (form.dataset.editId) {
            const index = this.students.findIndex(s => s.uuid === form.dataset.editId);
            if (index !== -1) {
                this.students[index] = student;
            }
        } else {
            this.students.push(student);
        }

        this.saveData();
        this.renderStudentList();
        this.updateStats();
        this.applyCurrentFilter();
        this.hideStudentModal();
    }

    deleteStudent(uuid) {
        if (confirm('确定要删除这个学生吗？')) {
            const seat = this.seats.find(s => s.student && s.student.uuid === uuid);
            if (seat) {
                seat.student = null;
            }
            
            this.students = this.students.filter(s => s.uuid !== uuid);
            this.saveData();
            this.renderStudentList();
            this.renderClassroom();
            this.updateStats();
            this.applyCurrentFilter();
        }
    }

    generateUUID() {
        return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
            const r = Math.random() * 16 | 0;
            const v = c == 'x' ? r : (r & 0x3 | 0x8);
            return v.toString(16);
        });
    }

    renderStudentList() {
        const container = document.getElementById('studentList');
        container.innerHTML = '';

        this.students.forEach(student => {
            const item = document.createElement('div');
            item.className = 'student-item';
            item.draggable = true;
            
            const isSeated = this.seats.some(seat => seat.student && seat.student.uuid === student.uuid);
            if (isSeated) {
                item.classList.add('seated');
            }

            // 构建学生详情信息
            let details = [];
            if (student.gender) {
                const genderText = student.gender === 'male' ? '男' : student.gender === 'female' ? '女' : student.gender;
                details.push(`性别: ${genderText}`);
            }
            if (student.needsFrontSeat) details.push('需前排');
            
            item.innerHTML = `
                <div class="student-info">
                    <div class="student-name">${student.name}</div>
                    <div class="student-details">
                        ${details.join(' | ')}
                    </div>
                    ${student.notes ? `<div class="student-notes">备注: ${student.notes}</div>` : ''}
                </div>
                <div class="student-actions">
                    <button class="btn btn-small btn-edit" onclick="app.showStudentModal(app.students.find(s => s.uuid === '${student.uuid}'))">编辑</button>
                    <button class="btn btn-small btn-secondary" onclick="app.deleteStudent('${student.uuid}')">删除</button>
                </div>
            `;

            item.addEventListener('dragstart', (e) => {
                e.dataTransfer.setData('text/plain', student.uuid);
                item.classList.add('dragging');
            });

            item.addEventListener('dragend', () => {
                item.classList.remove('dragging');
            });

            container.appendChild(item);
        });
    }

    renderClassroom() {
        const container = document.getElementById('classroomGrid');
        container.innerHTML = '';
        container.style.gridTemplateColumns = `repeat(${this.cols}, 1fr)`;
        container.style.gridTemplateRows = `repeat(${this.rows}, 1fr) auto`;

        this.seats.forEach(seat => {
            const seatElement = document.createElement('div');
            seatElement.className = 'seat';
            seatElement.dataset.seatId = seat.id;
            
            if (seat.student) {
                seatElement.classList.add('seat-occupied');
                if (seat.student.gender === 'male') {
                    seatElement.classList.add('male');
                } else if (seat.student.gender === 'female') {
                    seatElement.classList.add('female');
                }
                
                // 根据姓名长度确定字体大小类名
                let nameLengthClass = '';
                const nameLength = seat.student.name.length;
                if (nameLength === 4) {
                    nameLengthClass = ' name-4';
                } else if (nameLength === 5) {
                    nameLengthClass = ' name-5';
                } else if (nameLength === 6) {
                    nameLengthClass = ' name-6';
                } else if (nameLength === 7) {
                    nameLengthClass = ' name-7';
                } else if (nameLength >= 8) {
                    nameLengthClass = ' name-8-plus';
                }
                
                // 座位编号显示: 转换内部坐标为教室视角编号
                // this.rows - seat.row: 将内部行号转换为教室行号(底部=1, 顶部=最大)
                // seat.col + 1: 将内部列号转换为教室列号(左=1, 右=最大)
                seatElement.innerHTML = `
                    <div class="seat-number">${this.rows - seat.row}-${seat.col + 1}</div>
                    <div class="student-name-display${nameLengthClass}" draggable="true" data-student-uuid="${seat.student.uuid}" data-source-seat-id="${seat.id}">${seat.student.name}</div>
                    <div class="seat-remove-btn" data-seat-id="${seat.id}" title="移除学生">×</div>
                `;
            } else {
                seatElement.classList.add('seat-empty');
                // 空座位也显示教室视角的座位编号
                seatElement.innerHTML = `<div class="seat-number">${this.rows - seat.row}-${seat.col + 1}</div>`;
            }

            seatElement.addEventListener('click', () => {
                this.selectSeat(seat);
            });

            seatElement.addEventListener('dragover', (e) => {
                e.preventDefault();
                seatElement.classList.add('drag-over');
            });

            seatElement.addEventListener('dragleave', () => {
                seatElement.classList.remove('drag-over');
            });

            seatElement.addEventListener('drop', (e) => {
                e.preventDefault();
                seatElement.classList.remove('drag-over');
                const dragDataString = e.dataTransfer.getData('text/plain');
                
                let studentUuid, sourceSeatId;
                try {
                    // 尝试解析JSON格式的拖拽数据（座位间拖拽）
                    const dragData = JSON.parse(dragDataString);
                    studentUuid = dragData.studentUuid;
                    sourceSeatId = dragData.sourceSeatId;
                } catch {
                    // 如果解析失败，说明是旧格式（从学生列表拖拽）
                    studentUuid = dragDataString;
                    sourceSeatId = null;
                }
                
                this.assignStudentToSeat(studentUuid, seat.id, sourceSeatId);
            });

            container.appendChild(seatElement);
        });

        // Add podium below the seats
        const podiumElement = document.createElement('div');
        podiumElement.className = 'podium-in-grid';
        podiumElement.style.gridColumn = `1 / -1`;
        podiumElement.style.gridRow = `${this.rows + 1}`;
        podiumElement.innerHTML = `
            <div class="podium-shape">
                <span class="podium-text">讲台</span>
            </div>
        `;
        container.appendChild(podiumElement);

        // 为移除按钮添加事件监听器
        this.setupSeatRemoveListeners();
        
        // 为座位内的学生姓名添加拖拽事件监听器
        this.setupSeatDragListeners();
        
        this.updateClassroomInfo();
        
        // 应用坐标显示设置
        if (this.showCoordinates) {
            container.classList.remove('hide-coordinates');
        } else {
            container.classList.add('hide-coordinates');
        }
        
        // 同步主界面复选框状态
        document.getElementById('showCoordinatesMain').checked = this.showCoordinates;
    }

    setupSeatRemoveListeners() {
        document.querySelectorAll('.seat-remove-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.stopPropagation(); // 阻止事件冒泡到座位点击事件
                const seatId = e.target.dataset.seatId;
                this.removeStudentFromSeat(seatId);
            });
        });
    }

    setupSeatDragListeners() {
        document.querySelectorAll('.student-name-display[draggable="true"]').forEach(nameElement => {
            nameElement.addEventListener('dragstart', (e) => {
                e.stopPropagation(); // 阻止事件冒泡到座位点击事件
                const studentUuid = e.target.dataset.studentUuid;
                const sourceSeatId = e.target.dataset.sourceSeatId;
                
                // 传递拖拽数据，包含学生UUID和来源信息
                const dragData = JSON.stringify({
                    studentUuid: studentUuid,
                    source: 'seat',
                    sourceSeatId: sourceSeatId
                });
                e.dataTransfer.setData('text/plain', dragData);
                
                // 添加拖拽视觉反馈
                e.target.classList.add('dragging');
                setTimeout(() => {
                    e.target.closest('.seat').classList.add('seat-dragging');
                }, 0);
            });

            nameElement.addEventListener('dragend', (e) => {
                // 移除拖拽视觉反馈
                e.target.classList.remove('dragging');
                e.target.closest('.seat').classList.remove('seat-dragging');
            });
        });
    }

    removeStudentFromSeat(seatId) {
        const seat = this.seats.find(s => s.id === seatId);
        if (seat && seat.student) {
            // 添加到历史记录
            this.addToHistory('seatArrangement', { seats: this.seats });
            
            // 移除学生
            seat.student = null;
            
            // 更新界面
            this.saveData();
            this.renderClassroom();
            this.renderStudentList();
            this.updateStats();
            this.applyCurrentFilter();
        }
    }

    selectSeat(seat) {
        document.querySelectorAll('.seat').forEach(s => s.classList.remove('seat-selected'));
        
        if (this.selectedSeat === seat) {
            this.selectedSeat = null;
        } else {
            this.selectedSeat = seat;
            document.querySelector(`[data-seat-id="${seat.id}"]`).classList.add('seat-selected');
        }
    }

    assignStudentToSeat(studentUuid, seatId, sourceSeatId = null) {
        const student = this.students.find(s => s.uuid === studentUuid);
        const targetSeat = this.seats.find(s => s.id === seatId);
        
        if (!student || !targetSeat) return;

        this.addToHistory('seatArrangement', { seats: this.seats });

        // 如果提供了源座位ID，优先使用它（座位间拖拽）
        const currentSeat = sourceSeatId 
            ? this.seats.find(s => s.id === sourceSeatId)
            : this.seats.find(s => s.student && s.student.uuid === studentUuid);

        // 如果目标座位已被占用，需要处理位置交换
        if (targetSeat.student) {
            const displacedStudent = targetSeat.student;
            
            // 座位间拖拽：直接交换两个学生的位置
            if (currentSeat && sourceSeatId) {
                currentSeat.student = displacedStudent;
            } 
            // 从学生列表拖拽：将被替换的学生移回未安排状态
            else if (currentSeat) {
                currentSeat.student = displacedStudent;
            }
            // 如果被拖拽的学生之前没有座位，被替换的学生将变为未安排状态
        }

        // 清空学生的原座位
        if (currentSeat) {
            // 如果不是交换操作（即目标座位为空），才清空原座位
            if (!targetSeat.student || !sourceSeatId) {
                currentSeat.student = null;
            }
        }

        // 将学生分配到目标座位
        targetSeat.student = student;
        
        this.saveData();
        this.renderClassroom();
        this.renderStudentList();
        this.updateStats();
        this.applyCurrentFilter();
    }

    randomSeatArrangement() {
        if (this.students.length === 0) {
            alert('请先添加学生');
            return;
        }

        this.addToHistory('seatArrangement', { seats: this.seats });

        this.seats.forEach(seat => seat.student = null);

        const availableSeats = [...this.seats];
        const studentsToSeat = [...this.students];

        // 视力不佳的学生优先安排到前排
        // 注意: seat.row < Math.ceil(this.rows / 3) 表示前1/3排
        // 因为内部坐标系中 row=0 为第一排(靠近讲台), row值越小越靠前
        const frontRowSeats = availableSeats.filter(seat => seat.row < Math.ceil(this.rows / 3));
        const visionImpairedStudents = studentsToSeat.filter(student => student.needsFrontSeat);

        visionImpairedStudents.forEach(student => {
            if (frontRowSeats.length > 0) {
                const randomIndex = Math.floor(Math.random() * frontRowSeats.length);
                const seat = frontRowSeats.splice(randomIndex, 1)[0];
                const seatIndex = availableSeats.indexOf(seat);
                availableSeats.splice(seatIndex, 1);
                
                seat.student = student;
                studentsToSeat.splice(studentsToSeat.indexOf(student), 1);
            }
        });

        // 为剩余学生随机分配座位
        // availableSeats包含所有剩余座位(各排各列), 无特殊位置偏好
        while (studentsToSeat.length > 0 && availableSeats.length > 0) {
            const randomStudentIndex = Math.floor(Math.random() * studentsToSeat.length);
            const randomSeatIndex = Math.floor(Math.random() * availableSeats.length);
            
            const student = studentsToSeat.splice(randomStudentIndex, 1)[0];
            const seat = availableSeats.splice(randomSeatIndex, 1)[0];
            
            seat.student = student;
        }

        this.saveData();
        this.renderClassroom();
        this.renderStudentList();
        this.updateStats();
        this.applyCurrentFilter();
    }

    clearAllSeats() {
        if (confirm('确定要清空所有座位吗？')) {
            this.addToHistory('seatArrangement', { seats: this.seats });
            
            this.seats.forEach(seat => seat.student = null);
            this.saveData();
            this.renderClassroom();
            this.renderStudentList();
            this.updateStats();
        }
    }

    applyNewLayout() {
        const newRows = parseInt(document.getElementById('rowCount').value);
        const newCols = parseInt(document.getElementById('colCount').value);
        
        if (newRows < 1 || newRows > 15 || newCols < 1 || newCols > 12) {
            alert('行数范围: 1-15，列数范围: 1-12');
            return;
        }

        if (confirm('改变布局将清空现有座位安排，确定继续吗？')) {
            // 记录布局改变前的状态
            this.addToHistory('seatArrangement', { seats: this.seats });
            
            this.rows = newRows;
            this.cols = newCols;
            this.initializeSeats();
            this.saveData();
            this.renderClassroom();
            this.renderStudentList();
            this.updateStats();
        }
    }

    saveCurrentLayout() {
        const layoutName = prompt('请输入方案名称:');
        if (layoutName) {
            const layouts = JSON.parse(localStorage.getItem('savedLayouts') || '{}');
            layouts[layoutName] = {
                seats: this.seats,
                students: this.students,
                rows: this.rows,
                cols: this.cols,
                timestamp: Date.now()
            };
            localStorage.setItem('savedLayouts', JSON.stringify(layouts));
            alert('方案保存成功！');
        }
    }

    exportLayout() {
        // 检查是否有html2canvas库
        if (typeof html2canvas === 'undefined') {
            console.error('html2canvas库未加载');
            alert('导出功能需要html2canvas库，请检查网络连接后刷新页面重试');
            return;
        }

        const classroomGrid = document.getElementById('classroomGrid');
        if (!classroomGrid) {
            alert('无法找到座位图元素');
            return;
        }

        // 显示加载提示
        const originalText = document.getElementById('exportLayout').textContent;
        document.getElementById('exportLayout').textContent = '导出中...';
        document.getElementById('exportLayout').disabled = true;

        // 使用html2canvas捕获classroom-grid元素，导出包含座位和讲台的完整布局
        html2canvas(classroomGrid, {
            backgroundColor: null, // 透明背景，无背景色
            scale: 2, // 提高分辨率
            useCORS: true,
            allowTaint: true,
            logging: true,
            onclone: (clonedDoc) => {
                // 移除背景和边框，保持与网页显示一致
                const clonedGrid = clonedDoc.getElementById('classroomGrid');
                if (clonedGrid) {
                    clonedGrid.style.background = 'linear-gradient(135deg, #f8fafc 0%, #e2e8f0 100%)';
                    clonedGrid.style.border = '2px solid #94a3b8';
                    clonedGrid.style.borderRadius = '8px';
                    clonedGrid.style.padding = '4% 4% 0.5% 4%';
                    clonedGrid.style.margin = '0';
                }
                
                // 确保座位样式与网页显示完全一致
                const clonedSeats = clonedDoc.querySelectorAll('.seat');
                clonedSeats.forEach(seat => {
                    seat.style.width = '100px';
                    seat.style.height = '50px';
                    seat.style.borderRadius = '6px';
                    seat.style.display = 'flex';
                    seat.style.alignItems = 'center';
                    seat.style.justifyContent = 'center';
                    seat.style.position = 'relative';
                    seat.style.transition = 'all 0.2s ease';
                    seat.style.border = '2px solid transparent';
                    seat.style.userSelect = 'none';
                    
                    if (seat.classList.contains('seat-occupied')) {
                        seat.style.backgroundColor = '#f1f5f9';
                        seat.style.borderColor = '#cbd5e1';
                        seat.style.fontSize = '1.7rem';
                        seat.style.fontWeight = '600';
                        
                        // 根据性别设置不同颜色
                        if (seat.classList.contains('male')) {
                            seat.style.color = '#2563eb'; // 蓝色 - 男生
                        } else if (seat.classList.contains('female')) {
                            seat.style.color = '#ec4899'; // 粉色 - 女生
                        } else {
                            seat.style.color = '#1e293b'; // 默认深色
                        }
                    } else {
                        seat.style.backgroundColor = '#f1f5f9';
                        seat.style.borderColor = '#cbd5e1';
                        seat.style.color = '#94a3b8';
                        seat.style.fontSize = '0.7rem';
                        seat.style.fontWeight = '500';
                    }
                });
                
                // 确保讲台样式与网页显示完全一致
                const clonedPodiums = clonedDoc.querySelectorAll('.podium-in-grid');
                clonedPodiums.forEach(podium => {
                    podium.style.display = 'flex';
                    podium.style.justifyContent = 'center';
                    podium.style.alignItems = 'center';
                    podium.style.padding = '2%';
                    podium.style.marginTop = '1%';
                });
                
                const clonedPodiumShapes = clonedDoc.querySelectorAll('.podium-shape');
                clonedPodiumShapes.forEach(shape => {
                    shape.style.position = 'relative';
                    shape.style.width = '200px';
                    shape.style.height = '50px';
                    shape.style.background = 'linear-gradient(135deg, #8b5a3c 0%, #a0692e 100%)';
                    shape.style.borderRadius = '8px';
                    shape.style.display = 'flex';
                    shape.style.alignItems = 'center';
                    shape.style.justifyContent = 'center';
                    shape.style.boxShadow = '0 4px 8px rgba(0,0,0,0.1), inset 0 1px 2px rgba(255,255,255,0.3)';
                    
                    // 移除3D效果和梯形，使用简单矩形
                    shape.style.transform = 'none';
                    shape.style.clipPath = 'none';
                });
                
                const clonedPodiumTexts = clonedDoc.querySelectorAll('.podium-text');
                clonedPodiumTexts.forEach(text => {
                    text.style.color = 'white';
                    text.style.fontSize = '1.1rem';
                    text.style.fontWeight = '600';
                    text.style.textShadow = '1px 1px 2px rgba(0,0,0,0.3)';
                    text.style.letterSpacing = '0.5px';
                    
                    // 移除3D文字效果
                    text.style.transform = 'none';
                });
                
                // 确保座位号码显示正确
                const clonedNumbers = clonedDoc.querySelectorAll('.seat-number');
                clonedNumbers.forEach(num => {
                    num.style.position = 'absolute';
                    num.style.top = '2px';
                    num.style.left = '2px';
                    num.style.fontSize = '0.6rem';
                    num.style.fontWeight = '700';
                    num.style.color = '#64748b';
                    num.style.backgroundColor = 'rgba(255, 255, 255, 0.9)';
                    num.style.padding = '1px 3px';
                    num.style.borderRadius = '2px';
                    num.style.zIndex = '1';
                    num.style.lineHeight = '1';
                });
                
                // 确保学生姓名显示正确，支持动态字体大小
                const clonedNames = clonedDoc.querySelectorAll('.student-name-display');
                clonedNames.forEach(name => {
                    name.style.position = 'absolute';
                    name.style.top = '50%';
                    name.style.left = '50%';
                    name.style.transform = 'translate(-50%, -50%)';
                    name.style.fontWeight = '600';
                    name.style.zIndex = '2';
                    name.style.textAlign = 'center';
                    name.style.width = '95%';
                    name.style.overflow = 'hidden';
                    name.style.textOverflow = 'ellipsis';
                    name.style.whiteSpace = 'nowrap';
                    
                    // 根据名字长度设置字体大小
                    if (name.classList.contains('name-4')) {
                        name.style.fontSize = '1.1rem';
                        name.style.width = '98%';
                    } else if (name.classList.contains('name-5')) {
                        name.style.fontSize = '0.9rem';
                        name.style.width = '98%';
                    } else if (name.classList.contains('name-6')) {
                        name.style.fontSize = '0.75rem';
                        name.style.width = '98%';
                    } else if (name.classList.contains('name-7')) {
                        name.style.fontSize = '0.65rem';
                        name.style.width = '98%';
                    } else if (name.classList.contains('name-8-plus')) {
                        name.style.fontSize = '0.55rem';
                        name.style.width = '98%';
                    } else {
                        name.style.fontSize = '1.7rem';
                    }
                });
            },
            ignoreElements: (element) => {
                // 只忽略移除按钮，保留讲台
                return element.classList.contains('seat-remove-btn');
            }
        }).then(canvas => {
            console.log('html2canvas成功，画布尺寸:', canvas.width, 'x', canvas.height);
            
            // 下载图片
            const link = document.createElement('a');
            link.download = `座位表_${new Date().toLocaleDateString().replace(/\//g, '-')}.png`;
            link.href = canvas.toDataURL('image/png');
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            
            console.log('座位图导出成功');
            
        }).catch(error => {
            console.error('html2canvas导出失败:', error);
            alert('使用html2canvas导出失败，将使用备用方法');
            this.exportLayoutCanvas();
        }).finally(() => {
            // 恢复按钮状态
            document.getElementById('exportLayout').textContent = originalText;
            document.getElementById('exportLayout').disabled = false;
        });
    }

    // 备用canvas导出方法 - 导出完整教室布局（座位+讲台），匹配网页显示
    exportLayoutCanvas() {
        console.log('使用备用canvas导出方法');
        
        const canvas = document.createElement('canvas');
        const ctx = canvas.getContext('2d');
        
        // 使用与网页相同的尺寸和间距
        const seatWidth = 100;
        const seatHeight = 50;
        const gridPadding = 40; // 网格周围的padding
        const rowGap = 5; // 行间距
        const colGap = 2; // 列间距
        const podiumHeight = 60; // 讲台高度
        const podiumMargin = 10; // 讲台上边距
        
        // 计算画布尺寸 - 包含座位区域和讲台
        const gridWidth = this.cols * seatWidth + (this.cols - 1) * colGap;
        const gridHeight = this.rows * seatHeight + (this.rows - 1) * rowGap;
        canvas.width = gridWidth + gridPadding * 2;
        canvas.height = gridHeight + podiumHeight + podiumMargin + gridPadding * 2;
        
        // 绘制教室背景
        const gradient = ctx.createLinearGradient(0, 0, canvas.width, canvas.height);
        gradient.addColorStop(0, '#f8fafc');
        gradient.addColorStop(1, '#e2e8f0');
        ctx.fillStyle = gradient;
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        
        // 绘制教室边框
        ctx.strokeStyle = '#94a3b8';
        ctx.lineWidth = 4;
        ctx.strokeRect(2, 2, canvas.width - 4, canvas.height - 4);
        
        // 绘制座位
        this.seats.forEach(seat => {
            const x = gridPadding + seat.col * (seatWidth + colGap);
            const y = gridPadding + seat.row * (seatHeight + rowGap);
            
            // 绘制座位圆角矩形
            this.drawRoundedRect(ctx, x, y, seatWidth, seatHeight, 6);
            
            if (seat.student) {
                // 已占用座位
                ctx.fillStyle = '#f1f5f9';
                ctx.fill();
                
                // 座位边框
                ctx.strokeStyle = '#cbd5e1';
                ctx.lineWidth = 2;
                ctx.stroke();
                
                // 根据性别设置姓名颜色
                let nameColor = '#1e293b'; // 默认颜色
                if (seat.student.gender === 'male') {
                    nameColor = '#2563eb'; // 蓝色 - 男生
                } else if (seat.student.gender === 'female') {
                    nameColor = '#ec4899'; // 粉色 - 女生
                }
                
                // 学生姓名 - 根据名字长度调整字体大小
                ctx.fillStyle = nameColor;
                const nameLength = seat.student.name.length;
                let fontSize = 27; // 默认 1.7rem 对应约 27px
                if (nameLength >= 8) fontSize = 9;  // 0.55rem
                else if (nameLength === 7) fontSize = 10; // 0.65rem
                else if (nameLength === 6) fontSize = 12; // 0.75rem
                else if (nameLength === 5) fontSize = 14; // 0.9rem
                else if (nameLength === 4) fontSize = 18; // 1.1rem
                
                ctx.font = `600 ${fontSize}px Arial, sans-serif`;
                ctx.textAlign = 'center';
                ctx.textBaseline = 'middle';
                ctx.fillText(seat.student.name, x + seatWidth / 2, y + seatHeight / 2);
                
                // 座位号码背景
                const seatNumber = `${this.rows - seat.row}-${seat.col + 1}`;
                ctx.fillStyle = 'rgba(255, 255, 255, 0.9)';
                ctx.fillRect(x + 2, y + 2, 20, 12);
                
                // 座位号码文字
                ctx.fillStyle = '#64748b';
                ctx.font = '700 10px Arial, sans-serif';
                ctx.textAlign = 'left';
                ctx.textBaseline = 'top';
                ctx.fillText(seatNumber, x + 4, y + 4);
            } else {
                // 空座位
                ctx.fillStyle = '#f1f5f9';
                ctx.fill();
                
                // 座位边框
                ctx.strokeStyle = '#cbd5e1';
                ctx.lineWidth = 2;
                ctx.stroke();
                
                // 座位号码（居中）
                ctx.fillStyle = '#94a3b8';
                ctx.font = '500 11px Arial, sans-serif';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'middle';
                ctx.fillText(`${this.rows - seat.row}-${seat.col + 1}`, x + seatWidth / 2, y + seatHeight / 2);
            }
        });
        
        // 绘制讲台
        const podiumY = gridPadding + gridHeight + podiumMargin;
        const podiumX = (canvas.width - 200) / 2; // 居中
        
        // 绘制讲台梯形形状
        this.drawPodium(ctx, podiumX, podiumY, 200, 50);
        
        console.log('备用canvas导出完成，画布尺寸:', canvas.width, 'x', canvas.height);
        
        const link = document.createElement('a');
        link.download = `座位表_${new Date().toLocaleDateString().replace(/\//g, '-')}.png`;
        link.href = canvas.toDataURL('image/png');
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        
        console.log('备用方法导出成功');
    }

    // 绘制圆角矩形的辅助方法
    drawRoundedRect(ctx, x, y, width, height, radius) {
        ctx.beginPath();
        ctx.moveTo(x + radius, y);
        ctx.lineTo(x + width - radius, y);
        ctx.quadraticCurveTo(x + width, y, x + width, y + radius);
        ctx.lineTo(x + width, y + height - radius);
        ctx.quadraticCurveTo(x + width, y + height, x + width - radius, y + height);
        ctx.lineTo(x + radius, y + height);
        ctx.quadraticCurveTo(x, y + height, x, y + height - radius);
        ctx.lineTo(x, y + radius);
        ctx.quadraticCurveTo(x, y, x + radius, y);
        ctx.closePath();
    }

    // 绘制讲台的辅助方法
    drawPodium(ctx, x, y, width, height) {
        // 绘制讲台矩形背景
        ctx.save();
        
        // 绘制简单矩形（带圆角）
        this.drawRoundedRect(ctx, x, y, width, height, 8);
        
        // 绘制讲台渐变背景
        const gradient = ctx.createLinearGradient(x, y, x + width, y + height);
        gradient.addColorStop(0, '#8b5a3c');
        gradient.addColorStop(1, '#a0692e');
        ctx.fillStyle = gradient;
        ctx.fill();
        
        // 绘制讲台边框
        ctx.strokeStyle = '#6b4226';
        ctx.lineWidth = 1;
        ctx.stroke();
        
        // 绘制讲台阴影效果
        ctx.shadowColor = 'rgba(0, 0, 0, 0.1)';
        ctx.shadowBlur = 8;
        ctx.shadowOffsetX = 0;
        ctx.shadowOffsetY = 4;
        
        // 绘制讲台文字
        ctx.shadowColor = 'rgba(0, 0, 0, 0.3)';
        ctx.shadowBlur = 2;
        ctx.shadowOffsetX = 1;
        ctx.shadowOffsetY = 1;
        ctx.fillStyle = 'white';
        ctx.font = '600 18px Arial, sans-serif';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText('讲台', x + width / 2, y + height / 2);
        
        ctx.restore();
    }

    updateStats() {
        const totalStudents = this.students.length;
        const seatedStudents = this.seats.filter(seat => seat.student).length;
        
        document.getElementById('totalStudents').textContent = totalStudents;
        document.getElementById('seatedStudents').textContent = seatedStudents;
    }

    updateClassroomInfo() {
        document.getElementById('classroomSize').textContent = `${this.rows}行 × ${this.cols}列`;
        
        const totalSeats = this.seats.length;
        const occupiedSeats = this.seats.filter(seat => seat.student).length;
        const occupancyRate = totalSeats > 0 ? Math.round((occupiedSeats / totalSeats) * 100) : 0;
        
        document.getElementById('occupancyRate').textContent = `占用率: ${occupancyRate}%`;
    }

    filterStudents(searchTerm) {
        const items = document.querySelectorAll('.student-item');
        items.forEach(item => {
            const name = item.querySelector('.student-name').textContent.toLowerCase();
            const details = item.querySelector('.student-details').textContent.toLowerCase();
            
            if (name.includes(searchTerm.toLowerCase()) || details.includes(searchTerm.toLowerCase())) {
                item.style.display = 'flex';
            } else {
                item.style.display = 'none';
            }
        });
    }

    filterStudentsByStatus(status) {
        const items = document.querySelectorAll('.student-item');
        items.forEach(item => {
            const isSeated = item.classList.contains('seated');
            
            switch (status) {
                case 'all':
                    item.style.display = 'flex';
                    break;
                case 'seated':
                    item.style.display = isSeated ? 'flex' : 'none';
                    break;
                case 'unseated':
                    item.style.display = !isSeated ? 'flex' : 'none';
                    break;
            }
        });
    }

    applyCurrentFilter() {
        const filterSelect = document.getElementById('filterStudents');
        const searchInput = document.getElementById('searchStudent');
        
        if (filterSelect && searchInput) {
            this.filterStudentsByStatus(filterSelect.value);
            if (searchInput.value.trim()) {
                this.filterStudents(searchInput.value);
            }
        }
    }

    // Excel导入相关方法
    importExcelFile() {
        const fileInput = document.getElementById('excelFileInput');
        fileInput.click();
    }

    handleExcelFileSelect(event) {
        const file = event.target.files[0];
        if (!file) return;

        if (!file.name.match(/\.(xlsx|xls)$/)) {
            alert('请选择Excel文件 (.xlsx 或 .xls)');
            return;
        }

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                this.parseExcelData(e.target.result, file.name);
            } catch (error) {
                console.error('Excel文件读取失败:', error);
                alert('Excel文件读取失败，请检查文件格式是否正确');
            }
        };
        reader.readAsArrayBuffer(file);
    }

    parseExcelData(data, filename) {
        try {
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            
            // 转换为JSON数据
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            
            if (jsonData.length < 2) {
                alert('Excel文件中没有足够的数据');
                return;
            }

            // 获取表头和数据
            const headers = jsonData[0];
            const rows = jsonData.slice(1);
            
            // 调试信息
            console.log('Excel原始数据:', jsonData);
            console.log('表头数据:', headers);
            console.log('数据行数:', rows.length);

            // 查找列索引
            const columnMap = this.mapExcelColumns(headers);
            
            if (!columnMap.name && columnMap.name !== 0) {
                // 提供更详细的错误信息
                const headersList = headers.map((h, i) => `${i}: "${h || '(空)'}"`).join(', ');
                const headersDisplay = headers.map(h => h || '(空)').join(', ');
                console.error('未找到姓名列。当前表头:', headersList);
                alert(`Excel文件必须包含"姓名"列。\n\n当前检测到的表头: ${headersDisplay}\n\n请确保第一行包含"姓名"、"名字"或"name"列。`);
                return;
            }

            // 解析并验证数据
            const { validData, errorData } = this.validateExcelData(rows, columnMap);
            
            // 显示预览模态框
            this.showExcelPreviewModal(validData, errorData, filename);
            
        } catch (error) {
            console.error('Excel解析失败:', error);
            alert('Excel文件解析失败，请检查文件格式');
        }
    }

    mapExcelColumns(headers) {
        const columnMap = {};
        
        // 调试信息：打印headers内容
        console.log('Excel表头信息:', headers);
        
        if (!headers || !Array.isArray(headers)) {
            console.error('Invalid headers:', headers);
            return columnMap;
        }
        
        headers.forEach((header, index) => {
            if (!header && header !== 0) return; // 允许数字0作为header
            
            // 更强的字符串处理，移除所有可能的隐藏字符
            const headerStr = header.toString()
                .trim()
                .replace(/[\u200B-\u200D\uFEFF]/g, '') // 移除零宽字符
                .replace(/\s+/g, '') // 移除所有空格
                .toLowerCase();
            
            console.log(`处理表头 [${index}]: "${header}" -> "${headerStr}"`);
            
            // 姓名列的可能名称（增加更多匹配选项）
            if (headerStr.includes('姓名') || headerStr.includes('名字') || headerStr.includes('name') || 
                headerStr === '姓名' || headerStr === '名字' || headerStr === 'name') {
                columnMap.name = index;
                console.log(`找到姓名列: 索引 ${index}`);
            }
            // 学号列的可能名称
            if (headerStr.includes('学号') || headerStr.includes('编号') || headerStr.includes('id') || 
                headerStr.includes('number') || headerStr === '学号' || headerStr === '编号') {
                columnMap.id = index;
            }
            // 性别列的可能名称
            if (headerStr.includes('性别') || headerStr.includes('gender') || headerStr === '性别') {
                columnMap.gender = index;
            }
            // 视力列的可能名称
            if (headerStr.includes('视力') || headerStr.includes('近视') || headerStr.includes('眼镜') || 
                headerStr.includes('vision') || headerStr.includes('情况')) {
                columnMap.vision = index;
            }
            // 备注列的可能名称
            if (headerStr.includes('备注') || headerStr.includes('说明') || headerStr.includes('note') || 
                headerStr.includes('remark') || headerStr === '备注') {
                columnMap.notes = index;
            }
        });
        
        console.log('列映射结果:', columnMap);
        return columnMap;
    }

    validateExcelData(rows, columnMap) {
        const validData = [];
        const errorData = [];

        rows.forEach((row, rowIndex) => {
            const actualRow = rowIndex + 2; // Excel行号（从1开始，加上表头行）
            const errors = [];
            
            // 检查是否为空行
            if (!row || row.every(cell => !cell || cell.toString().trim() === '')) {
                return; // 跳过空行
            }

            const studentData = {
                name: row[columnMap.name] ? row[columnMap.name].toString().trim() : '',
                id: row[columnMap.id] ? row[columnMap.id].toString().trim() : '',
                gender: row[columnMap.gender] ? row[columnMap.gender].toString().trim() : '',
                vision: row[columnMap.vision] ? row[columnMap.vision].toString().trim() : '',
                notes: row[columnMap.notes] ? row[columnMap.notes].toString().trim() : ''
            };

            // 验证必填字段
            if (!studentData.name) {
                errors.push('姓名不能为空');
            }

            // 验证性别
            if (studentData.gender) {
                const genderLower = studentData.gender.toLowerCase();
                if (genderLower.includes('男') || genderLower.includes('male') || genderLower === 'm') {
                    studentData.gender = 'male';
                } else if (genderLower.includes('女') || genderLower.includes('female') || genderLower === 'f') {
                    studentData.gender = 'female';
                } else {
                    studentData.gender = '';
                }
            }

            // 验证视力信息
            if (studentData.vision) {
                const visionLower = studentData.vision.toLowerCase();
                if (visionLower.includes('近视') || visionLower.includes('不佳') || 
                    visionLower.includes('戴眼镜') || visionLower.includes('眼镜') ||
                    visionLower.includes('poor') || visionLower.includes('bad') ||
                    visionLower.includes('yes') || visionLower.includes('是') || visionLower.includes('需要')) {
                    studentData.needsFrontSeat = true;
                } else {
                    studentData.needsFrontSeat = false;
                }
            } else {
                studentData.needsFrontSeat = false;
            }

            if (errors.length > 0) {
                errorData.push({
                    row: actualRow,
                    data: studentData,
                    errors: errors,
                    originalData: row
                });
            } else {
                validData.push(studentData);
            }
        });

        return { validData, errorData };
    }

    showExcelPreviewModal(validData, errorData, filename) {
        this.importData = { validData, errorData, filename };
        
        // 更新统计信息
        document.getElementById('totalImportCount').textContent = validData.length + errorData.length;
        document.getElementById('validImportCount').textContent = validData.length;
        document.getElementById('errorImportCount').textContent = errorData.length;

        // 渲染有效数据表格
        this.renderValidDataTable(validData);
        
        // 渲染错误数据列表
        this.renderErrorDataList(errorData);
        
        // 显示模态框
        document.getElementById('excelPreviewModal').style.display = 'flex';
        
        // 默认显示有效数据标签页
        this.switchTab('valid');
    }

    renderValidDataTable(validData) {
        const tbody = document.getElementById('validDataBody');
        tbody.innerHTML = '';

        validData.forEach((student, index) => {
            const row = document.createElement('tr');
            row.innerHTML = `
                <td>${student.name}</td>
                <td>${student.id || '<span class="empty-cell">无</span>'}</td>
                <td>${student.gender === 'male' ? '男' : student.gender === 'female' ? '女' : '<span class="empty-cell">无</span>'}</td>
                <td>${student.needsFrontSeat ? '需要前排' : '正常'}</td>
                <td>${student.notes || '<span class="empty-cell">无</span>'}</td>
            `;
            tbody.appendChild(row);
        });
    }

    renderErrorDataList(errorData) {
        const container = document.getElementById('errorsList');
        container.innerHTML = '';

        if (errorData.length === 0) {
            container.innerHTML = '<div style="text-align: center; color: #6b7280; padding: 2rem;">没有错误数据</div>';
            return;
        }

        errorData.forEach(error => {
            const errorItem = document.createElement('div');
            errorItem.className = 'error-item';
            
            errorItem.innerHTML = `
                <div class="error-row">第 ${error.row} 行</div>
                <div class="error-message">${error.errors.join('、')}</div>
                <div class="error-data">
                    原始数据: ${JSON.stringify(error.originalData).replace(/"/g, '')}
                </div>
            `;
            
            container.appendChild(errorItem);
        });
    }

    switchTab(tabName) {
        // 更新标签按钮状态
        document.querySelectorAll('.tab-btn').forEach(btn => {
            btn.classList.remove('active');
        });
        document.querySelector(`[data-tab="${tabName}"]`).classList.add('active');

        // 更新内容显示
        document.querySelectorAll('.tab-content').forEach(content => {
            content.classList.remove('active');
        });
        document.getElementById(`${tabName}DataPreview`).classList.add('active');
    }

    hideExcelPreviewModal() {
        document.getElementById('excelPreviewModal').style.display = 'none';
        this.importData = null;
        
        // 重置文件输入
        document.getElementById('excelFileInput').value = '';
    }

    confirmExcelImport() {
        if (!this.importData || !this.importData.validData.length) {
            alert('没有有效数据可以导入');
            return;
        }

        const overwriteExisting = document.getElementById('overwriteExisting').checked;
        const skipInvalid = document.getElementById('skipInvalid').checked;
        
        let importCount = 0;
        let updateCount = 0;
        let skipCount = 0;

        this.importData.validData.forEach(studentData => {
            // 检查是否存在同名学生
            const existingStudent = this.students.find(s => s.name === studentData.name);
            
            if (existingStudent) {
                if (overwriteExisting) {
                    // 更新现有学生信息
                    existingStudent.id = studentData.id || existingStudent.id;
                    existingStudent.gender = studentData.gender || existingStudent.gender;
                    existingStudent.needsFrontSeat = studentData.needsFrontSeat;
                    existingStudent.notes = studentData.notes || existingStudent.notes;
                    updateCount++;
                } else {
                    skipCount++;
                }
            } else {
                // 添加新学生
                const newStudent = {
                    uuid: this.generateUUID(),
                    name: studentData.name,
                    id: studentData.id,
                    gender: studentData.gender,
                    needsFrontSeat: studentData.needsFrontSeat,
                    notes: studentData.notes,
                    seatId: null
                };
                this.students.push(newStudent);
                importCount++;
            }
        });

        // 保存数据并更新界面
        this.saveData();
        this.renderStudentList();
        this.updateStats();
        
        // 显示导入结果
        let message = `导入完成！\n新增学生: ${importCount} 人`;
        if (updateCount > 0) {
            message += `\n更新学生: ${updateCount} 人`;
        }
        if (skipCount > 0) {
            message += `\n跳过重复: ${skipCount} 人`;
        }
        
        alert(message);
        
        // 关闭模态框
        this.hideExcelPreviewModal();
    }

    downloadExcelTemplate() {
        try {
            // 创建示例数据，确保列名与解析逻辑一致
            const templateData = [
                ['姓名', '学号', '性别', '视力情况', '备注'],
                ['张三', '001', '男', '正常', '班长'],
                ['李四', '002', '女', '近视', '需要坐前排'],
                ['王五', '003', '男', '正常', ''],
                ['赵六', '004', '女', '戴眼镜', '']
            ];

            // 创建工作簿
            const wb = XLSX.utils.book_new();
            const ws = XLSX.utils.aoa_to_sheet(templateData);
            
            // 设置列宽
            ws['!cols'] = [
                { width: 12 }, // 姓名
                { width: 10 }, // 学号
                { width: 8 },  // 性别
                { width: 12 }, // 视力情况
                { width: 20 }  // 备注
            ];

            // 设置表头样式
            const headerRange = XLSX.utils.decode_range(ws['!ref']);
            for (let col = headerRange.s.c; col <= headerRange.e.c; col++) {
                const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
                if (!ws[cellAddress]) continue;
                ws[cellAddress].s = {
                    font: { bold: true },
                    fill: { fgColor: { rgb: "EFEFEF" } }
                };
            }

            XLSX.utils.book_append_sheet(wb, ws, '学生名单');
            
            // 设置工作簿属性，确保编码正确
            wb.Props = {
                Title: "学生名单导入模板",
                Subject: "学生信息",
                Author: "智能座位安排系统"
            };
            
            // 下载文件，指定文件格式
            XLSX.writeFile(wb, '学生名单导入模板.xlsx', {
                bookType: 'xlsx',
                bookSST: false,
                type: 'binary'
            });
            
            console.log('模板下载完成，列标题:', templateData[0]);
            
        } catch (error) {
            console.error('模板生成失败:', error);
            alert('模板生成失败，请重试');
        }
    }

    clearAllStudents() {
        const confirmMessage = `
确定要清空所有数据吗？

此操作将：
• 删除所有学生信息
• 清空所有座位安排
• 重置所有设置到初始状态
• 清空操作历史记录

⚠️ 此操作无法撤销！
        `.trim();

        if (confirm(confirmMessage)) {
            // 清空所有数据
            this.students = [];
            this.initializeSeats(); // 重新初始化座位
            this.selectedStudent = null;
            this.selectedSeat = null;
            this.history = [];
            this.historyIndex = -1;
            this.constraints = [];

            // 重置筛选器到默认状态
            document.getElementById('filterStudents').value = 'unseated';
            document.getElementById('searchStudent').value = '';

            // 保存并更新界面
            this.saveData();
            this.renderClassroom();
            this.renderStudentList();
            this.updateStats();
            this.updateHistoryButtons();
            this.applyCurrentFilter();

            // 显示成功消息
            alert('所有数据已清空，系统已重置到初始状态！');
        }
    }

    showSeatingSettingsModal() {
        document.getElementById('seatingSettingsModal').style.display = 'flex';
        // 更新约束列表显示
        this.renderConstraintList();
    }

    hideSeatingSettingsModal() {
        document.getElementById('seatingSettingsModal').style.display = 'none';
    }

    toggleCoordinatesDisplay(show) {
        this.showCoordinates = show;
        this.saveData();
        
        const classroomGrid = document.getElementById('classroomGrid');
        if (show) {
            classroomGrid.classList.remove('hide-coordinates');
        } else {
            classroomGrid.classList.add('hide-coordinates');
        }
    }

    addConstraint() {
        const constraintInput = document.getElementById('constraintInput');
        const constraintText = constraintInput.value.trim();
        
        if (!constraintText) {
            alert('请输入约束条件');
            return;
        }
        
        // 创建约束对象
        const constraint = {
            id: this.generateUUID(),
            text: constraintText,
            type: 'custom',
            active: true,
            timestamp: Date.now()
        };
        
        // 添加到约束列表
        this.constraints.push(constraint);
        
        // 保存数据
        this.saveData();
        
        // 清空输入框
        constraintInput.value = '';
        
        // 更新约束列表显示
        this.renderConstraintList();
        
        // 显示成功消息
        console.log('约束条件已添加:', constraint);
    }

    renderConstraintList() {
        const container = document.getElementById('constraintList');
        container.innerHTML = '';
        
        if (this.constraints.length === 0) {
            container.innerHTML = '<div class="no-constraints">暂无约束条件</div>';
            return;
        }
        
        this.constraints.forEach(constraint => {
            const constraintItem = document.createElement('div');
            constraintItem.className = 'constraint-item';
            constraintItem.innerHTML = `
                <div class="constraint-text">${constraint.text}</div>
                <div class="constraint-actions">
                    <button class="btn btn-small btn-secondary" onclick="app.removeConstraint('${constraint.id}')">删除</button>
                </div>
            `;
            container.appendChild(constraintItem);
        });
    }

    removeConstraint(constraintId) {
        if (confirm('确定要删除这个约束条件吗？')) {
            this.constraints = this.constraints.filter(c => c.id !== constraintId);
            this.saveData();
            this.renderConstraintList();
        }
    }

}

const app = new SeatingApp();