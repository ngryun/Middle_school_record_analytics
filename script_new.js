class ScoreAnalyzer {
    constructor() {
        this.filesData = new Map(); // 파일명 -> 분석 데이터 매핑
        this.combinedData = null; // 통합된 분석 데이터
        this.selectedFiles = null; // 사용자가 선택/드롭한 파일 목록
        this.initializeEventListeners();

        // If the page provides preloaded analysis data, render directly
        if (window.PRELOADED_DATA) {
            try {
                this.combinedData = window.PRELOADED_DATA;
                const upload = document.querySelector('.upload-section');
                if (upload) upload.style.display = 'none';
                const results = document.getElementById('results');
                if (results) results.style.display = 'block';
                this.displayResults();
                const exportBtn = document.getElementById('exportBtn');
                if (exportBtn) exportBtn.disabled = false;
            } catch (e) {
                console.error('PRELOADED_DATA 처리 중 오류:', e);
            }
        }
    }

    initializeEventListeners() {
        const fileInput = document.getElementById('excelFiles');
        const analyzeBtn = document.getElementById('analyzeBtn');
        const exportBtn = document.getElementById('exportBtn');
        const tabBtns = document.querySelectorAll('.tab-btn');
        const studentSearch = document.getElementById('studentSearch');
        const gradeSelect = document.getElementById('gradeSelect');
        const classSelect = document.getElementById('classSelect');
        const studentSelect = document.getElementById('studentSelect');
        const studentNameSearch = document.getElementById('studentNameSearch');
        const showStudentDetail = document.getElementById('showStudentDetail');
        const tableViewBtn = document.getElementById('tableViewBtn');
        const detailViewBtn = document.getElementById('detailViewBtn');
        const pdfClassBtn = document.getElementById('pdfClassBtn');
        const uploadSection = document.querySelector('.upload-section');
        const fileLabel = document.querySelector('.file-input-label');

        fileInput.addEventListener('change', (e) => {
            const files = Array.from(e.target.files);
            if (files.length > 0) {
                this.selectedFiles = files;
                this.displayFileList(files);
                analyzeBtn.disabled = false;
                this.hideError();
            }
        });

        // Drag & drop 지원
        if (uploadSection) {
            const prevent = (ev) => {
                ev.preventDefault();
                ev.stopPropagation();
            };
            const setDragState = (on) => {
                if (fileLabel) fileLabel.classList.toggle('dragover', on);
                uploadSection.classList.toggle('dragover', on);
            };

            ['dragover', 'drop'].forEach(evt => {
                window.addEventListener(evt, (ev) => {
                    prevent(ev);
                });
            });

            ['dragenter', 'dragover'].forEach(evt => {
                uploadSection.addEventListener(evt, (ev) => {
                    prevent(ev);
                    setDragState(true);
                });
            });
            ['dragleave', 'dragend'].forEach(evt => {
                uploadSection.addEventListener(evt, (ev) => {
                    prevent(ev);
                    setDragState(false);
                });
            });
            uploadSection.addEventListener('drop', (ev) => {
                prevent(ev);
                setDragState(false);
                const dropped = Array.from(ev.dataTransfer?.files || []);
                const files = dropped.filter(f => /\\.(xlsx|xls)$/i.test(f.name));
                if (files.length === 0) {
                    this.showError('XLS/XLSX 파일을 드래그하여 업로드하세요.');
                    return;
                }
                this.selectedFiles = files;
                this.displayFileList(files);
                analyzeBtn.disabled = false;
                this.hideError();
                try { if (fileInput) fileInput.files = ev.dataTransfer.files; } catch (_) {}
            });
        }

        analyzeBtn.addEventListener('click', () => {
            this.analyzeFiles();
        });

        // Tab switching
        tabBtns.forEach(btn => {
            btn.addEventListener('click', () => this.switchTab(btn.dataset.tab));
        });

        // Student search and filtering
        if (studentSearch) {
            studentSearch.addEventListener('input', (e) => {
                this.filterStudentTable(e.target.value);
            });
        }

        // Student selectors
        if (gradeSelect) {
            gradeSelect.addEventListener('change', () => {
                this.updateClassOptions();
                this.updateStudentOptions();
            });
        }

        if (classSelect) {
            classSelect.addEventListener('change', () => {
                this.updateStudentOptions();
            });
        }

        if (studentSelect) {
            studentSelect.addEventListener('change', () => {
                const selectedValue = studentSelect.value;
                showStudentDetail.disabled = !selectedValue;
            });
        }

        if (studentNameSearch) {
            studentNameSearch.addEventListener('input', (e) => {
                this.searchStudentByName(e.target.value);
            });
        }

        // View toggle buttons
        if (tableViewBtn) {
            tableViewBtn.addEventListener('click', () => {
                this.switchView('table');
            });
        }

        if (detailViewBtn) {
            detailViewBtn.addEventListener('click', () => {
                this.switchView('detail');
            });
        }

        // Show student detail
        if (showStudentDetail) {
            showStudentDetail.addEventListener('click', () => {
                const selectedStudent = studentSelect.value;
                if (selectedStudent) {
                    this.showStudentDetailView(selectedStudent);
                }
            });
        }

        // PDF export for class
        if (pdfClassBtn) {
            pdfClassBtn.addEventListener('click', () => {
                this.exportClassPDF();
            });
        }
    }

    displayFileList(files) {
        const fileList = document.getElementById('fileList');
        if (!fileList) return;

        fileList.innerHTML = files.map((file, index) => `
            <div class="file-item">
                <span class="file-name">${file.name}</span>
                <span class="file-size">(${this.formatFileSize(file.size)})</span>
            </div>
        `).join('');
        fileList.style.display = 'block';
    }

    formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }

    showLoading() {
        const loading = document.getElementById('loading');
        if (loading) loading.style.display = 'block';
    }

    hideLoading() {
        const loading = document.getElementById('loading');
        if (loading) loading.style.display = 'none';
    }

    showError(message) {
        const errorDiv = document.getElementById('error');
        if (errorDiv) {
            errorDiv.textContent = message;
            errorDiv.style.display = 'block';
        }
    }

    hideError() {
        const errorDiv = document.getElementById('error');
        if (errorDiv) errorDiv.style.display = 'none';
    }

    async analyzeFiles() {
        if (!this.selectedFiles || this.selectedFiles.length === 0) {
            this.showError('분석할 파일을 선택하세요.');
            return;
        }

        this.showLoading();
        this.hideError();

        try {
            this.filesData.clear();
            
            for (const file of this.selectedFiles) {
                const data = await this.readExcelFile(file);
                const fileData = this.parseFileData(data, file.name);
                this.filesData.set(file.name, fileData);
            }
            
            this.combineAllData();
            this.displayResults();
            this.hideLoading();

            // Enable export button after successful analysis
            const exportBtn = document.getElementById('exportBtn');
            if (exportBtn) exportBtn.disabled = false;
        } catch (error) {
            this.hideLoading();
            this.showError('파일 분석 중 오류가 발생했습니다: ' + error.message);
        }
    }

    readExcelFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                    resolve({ rows, worksheet });
                } catch (error) {
                    reject(error);
                }
            };
            reader.onerror = () => reject(new Error('파일 읽기 실패'));
            reader.readAsArrayBuffer(file);
        });
    }

    parseFileData(parsed, fileName) {
        const rows = parsed?.rows || [];
        const worksheet = parsed?.worksheet || null;
        const fileData = {
            fileName: fileName,
            data: rows,
            subjects: [],
            students: new Map(), // 학생별 데이터를 Map으로 관리
            grade: 1,
            class: 1
        };

        console.log('파일 데이터 구조:', rows);

        // 학년/반 텍스트에서 숫자 추출 유틸
        const parseGradeClass = (text) => {
            if (!text) return null;
            const s = text.toString();
            const g = s.match(/(\d+)\s*학년/);
            const c = s.match(/(\d+)\s*반/);
            if (g || c) {
                return {
                    grade: g ? parseInt(g[1]) : undefined,
                    klass: c ? parseInt(c[1]) : undefined
                };
            }
            return null;
        };

        // 1) 시트의 A3 직접 확인 (병합 영향 고려해서 A1:D6까지 스캔)
        let extracted = null;
        try {
            if (worksheet) {
                const direct = worksheet['A3']?.w || worksheet['A3']?.v;
                extracted = parseGradeClass(direct);
                if (!extracted) {
                    const cols = ['A','B','C','D','E'];
                    for (let r = 1; r <= 8 && !extracted; r++) {
                        for (const col of cols) {
                            const cell = worksheet[`${col}${r}`];
                            const val = cell?.w ?? cell?.v;
                            const got = parseGradeClass(val);
                            if (got) { extracted = got; break; }
                        }
                    }
                }
            }
        } catch (e) {
            console.warn('시트 스캔 중 오류:', e);
        }

        // 2) rows 배열의 상단 일부 스캔 (sheet_to_json이 빈칸을 생략하는 케이스 대비)
        if (!extracted) {
            for (let r = 0; r < Math.min(10, rows.length) && !extracted; r++) {
                const row = rows[r] || [];
                for (let c = 0; c < Math.min(8, row.length); c++) {
                    const got = parseGradeClass(row[c]);
                    if (got) { extracted = got; break; }
                }
            }
        }

        if (extracted) {
            if (typeof extracted.grade === 'number') fileData.grade = extracted.grade;
            if (typeof extracted.klass === 'number') fileData.class = extracted.klass;
            console.log('추출된 학년/반:', fileData.grade, fileData.class);
        } else {
            console.log('학년/반 정보를 찾지 못했습니다. 기본값 사용');
        }

        // 헤더 행 및 컬럼 인덱스 탐지
        let headerIndex = 3; // 기본 가정: 4행이 헤더
        let idx = { num: 0, name: 1, grade: 2, sem: 3, subjCat: 4, subjName: 5, score: 6, ach: 7 };
        const normalize = (v) => (v == null ? '' : v.toString().replace(/\s+/g, ''));

        // 상단 10~15행 내에서 동적으로 헤더 탐색
        for (let r = 0; r < Math.min(15, rows.length); r++) {
            const row = rows[r] || [];
            const compact = row.map(normalize);
            const hasNum = compact.some(v => v === '번호' || v === '번호');
            const hasName = compact.some(v => v === '성명' || v === '성명(한글)');
            const hasScoreOrAch = compact.some(v => v.includes('원점수') || v.includes('과목평균') || v.includes('성취도'));
            if (hasNum && hasName && hasScoreOrAch) {
                headerIndex = r;
                break;
            }
        }

        // 컬럼 위치를 가능한 한 정확히 매핑
        const header = rows[headerIndex] || [];
        const headerCompact = header.map(normalize);
        const findIdx = (keywords) => {
            for (let i = 0; i < headerCompact.length; i++) {
                for (const k of keywords) {
                    if (headerCompact[i].includes(k)) return i;
                }
            }
            return -1;
        };

        const foundIdx = {
            num: findIdx(['번호','번호']),
            name: findIdx(['성명']),
            grade: findIdx(['학년']),
            sem: findIdx(['학기']),
            subjCat: findIdx(['교과']),
            subjName: findIdx(['과목']),
            score: findIdx(['원점수','과목평균']),
            ach: findIdx(['성취도','성취도(수강자수)'])
        };

        Object.keys(foundIdx).forEach(k => {
            if (foundIdx[k] >= 0) idx[k] = foundIdx[k];
        });

        console.log('탐지된 헤더 위치:', { headerIndex, idx });

        // 데이터 파싱 시작 행: 헤더 다음 행
        const startRow = headerIndex + 1;

        // 5행부터 데이터 파싱 (동적 시작 행)
        let currentStudent = null;
        let currentGrade = fileData.grade || null;
        let currentSemester = null;

        // 상단에서 학기(1/2학기) 텍스트도 탐지하여 초기값으로 사용
        const parseSemesterText = (text) => {
            if (!text) return null;
            const s = text.toString();
            const m = s.match(/(\d+)\s*학기/);
            return m ? parseInt(m[1]) : null;
        };

        let defaultSemester = null;
        try {
            if (worksheet) {
                const directS = parseSemesterText(worksheet['A3']?.w || worksheet['A3']?.v);
                if (directS) defaultSemester = directS;
                if (!defaultSemester) {
                    const cols = ['A','B','C','D','E'];
                    for (let r = 1; r <= 8 && !defaultSemester; r++) {
                        for (const col of cols) {
                            const cell = worksheet[`${col}${r}`];
                            const val = cell?.w ?? cell?.v;
                            const got = parseSemesterText(val);
                            if (got) { defaultSemester = got; break; }
                        }
                    }
                }
            }
        } catch (e) {
            console.warn('학기 스캔 중 오류:', e);
        }
        if (!defaultSemester) {
            for (let r = 0; r < Math.min(10, rows.length) && !defaultSemester; r++) {
                const row = rows[r] || [];
                for (let c = 0; c < Math.min(8, row.length); c++) {
                    const got = parseSemesterText(row[c]);
                    if (got) { defaultSemester = got; break; }
                }
            }
        }
        if (defaultSemester) currentSemester = defaultSemester;
        
        // 중간 섹션 전환 시에도 헤더 재탐지 함수
        const detectHeaderFromRow = (row) => {
            const headerRow = (row || []).map(normalize);
            const hasNum = headerRow.some(v => v === '번호');
            const hasName = headerRow.some(v => v === '성명');
            const hasAny = headerRow.some(v => v.includes('교과') || v.includes('과목') || v.includes('성취도') || v.includes('원점수'));
            if (hasNum && hasName && hasAny) {
                const localIdx = { ...idx };
                const findLocal = (keywords) => {
                    for (let i = 0; i < headerRow.length; i++) {
                        for (const k of keywords) {
                            if (headerRow[i].includes(k)) return i;
                        }
                    }
                    return -1;
                };
                const overrides = {
                    num: findLocal(['번호']),
                    name: findLocal(['성명']),
                    grade: findLocal(['학년']),
                    sem: findLocal(['학기']),
                    subjCat: findLocal(['교과']),
                    subjName: findLocal(['과목']),
                    score: findLocal(['원점수','과목평균']),
                    ach: findLocal(['성취도'])
                };
                Object.keys(overrides).forEach(k => { if (overrides[k] >= 0) localIdx[k] = overrides[k]; });
                return localIdx;
            }
            return null;
        };

        let lastSubjectCategory = null;
        // 페이지 번호/아티팩트 행 판단 유틸
        const isPageIndicatorRow = (row) => {
            const parts = (row || []).map(v => (v == null ? '' : v.toString().trim())).filter(Boolean);
            if (parts.length === 0) return false;
            const joined = parts.join(' ').replace(/\s+/g, ' ').trim();
            return /^\d+\s*\/\s*\d+$/.test(joined);
        };
        const isNonSubjectCell = (val) => {
            if (val == null) return true;
            const s = val.toString().trim();
            // 숫자나 슬래시만 있는 경우 과목으로 간주하지 않음
            if (/^\d+$/.test(s)) return true;
            if (/^\/?$/.test(s)) return true;
            if (/^\d+\s*\/\s*\d+$/.test(s)) return true;
            return false;
        };

        for (let i = startRow; i < rows.length; i++) {
            const row = rows[i];
            if (!row || row.length === 0) continue;

            console.log(`${i+1}행 데이터:`, row); // 디버깅용

            // 섹션 전환 감지: 새로운 헤더 행이면 인덱스 재설정 후 다음 행으로
            const maybeNewIdx = detectHeaderFromRow(row);
            if (maybeNewIdx) {
                idx = maybeNewIdx;
                lastSubjectCategory = null;
                console.log('섹션 전환: 헤더 재탐지', { rowIndex: i, idx });
                continue;
            }

            // 페이지 번호 행(예: "1 / 37")은 스킵
            if (isPageIndicatorRow(row)) {
                console.log('페이지 번호 행 스킵:', row);
                continue;
            }
            
            // 학번이 있는 경우 새로운 학생 (숫자인지 확인)
            const valNum = row[idx.num];
            const valName = row[idx.name];
            if (valNum && !isNaN(parseInt(valNum)) && normalize(valNum) !== '번호') {
                const studentNumber = parseInt(valNum);
                const studentName = valName || `학생${studentNumber}`; // B열: 성명
                
                // 헤더 행이나 잘못된 데이터 필터링 (정확히 '성명' 표기만 스킵)
                const nameNoSpace = studentName.toString().replace(/\s+/g, '');
                if (nameNoSpace === '성명') {
                    console.log('헤더 행 스킵:', row);
                    continue;
                }
                
                const studentKey = `${studentNumber}_${studentName}`;
                
                if (!fileData.students.has(studentKey)) {
                    fileData.students.set(studentKey, {
                        number: studentNumber,
                        name: studentName,
                        grade: fileData.grade,
                        class: fileData.class,
                        subjects: new Map() // 과목별 데이터
                    });
                }
                currentStudent = fileData.students.get(studentKey);
            }
            
            // 학년/학기 정보 업데이트 (병합으로 비는 경우가 있어 초기값/유지값 활용)
            if (row[idx.grade]) currentGrade = parseInt(row[idx.grade]);
            if (row[idx.sem]) currentSemester = parseInt(row[idx.sem]);
            
            // 과목 데이터 추출 (E열: 교과, F열: 과목, G열: 원점수/과목평균, H열: 성취도)
            // 교과와 과목이 있고 학생이 선택되어 있으면 처리. 학년/학기는 비어있으면 초기값/기본값 사용
            const usedGrade = (idx.grade >= 0 && row[idx.grade]) ? parseInt(row[idx.grade]) : (currentGrade || fileData.grade);
            const usedSemester = (idx.sem >= 0 && row[idx.sem]) ? parseInt(row[idx.sem]) : (currentSemester || defaultSemester || 1);
            const subjectCategory = (idx.subjCat >= 0 && row[idx.subjCat]) ? row[idx.subjCat] : lastSubjectCategory;
            const subjectNameCell = row[idx.subjName];
            if (subjectCategory && subjectNameCell && !isNonSubjectCell(subjectNameCell) && currentStudent && usedGrade && usedSemester) {
                const subject = subjectCategory; // 교과
                const subjectName = subjectNameCell; // 과목
                const scoreData = (idx.score >= 0) ? row[idx.score] : undefined; // 원점수/과목평균 (없을 수 있음)
                const achievement = (idx.ach >= 0) ? row[idx.ach] : undefined; // 성취도(수강자수)
                
                console.log(`${i+1}행에서 과목 데이터 발견: 학생=${currentStudent.name}, 교과=${subject}, 과목=${subjectName}, 학년=${currentGrade}, 학기=${currentSemester}`);
                
                // 점수 파싱 (예: "79/83.8" 형태에서 79는 원점수, 83.8은 과목평균)
                let originalScore = null;
                let subjectAverage = null;
                
                console.log(`점수 데이터 (G열):`, scoreData); // 디버깅용
                
                if (scoreData) {
                    const scoreText = scoreData.toString();
                    console.log(`점수 텍스트:`, scoreText); // 디버깅용
                    const scoreMatch = scoreText.match(/(\d+(?:\.\d+)?)\/(\d+(?:\.\d+)?)/);
                    console.log(`점수 매칭 결과:`, scoreMatch); // 디버깅용
                    if (scoreMatch) {
                        originalScore = parseFloat(scoreMatch[1]);
                        subjectAverage = parseFloat(scoreMatch[2]);
                        console.log(`파싱된 점수: 원점수=${originalScore}, 과목평균=${subjectAverage}`);
                    }
                }
                
                // 성취도 파싱 (예: "C(195)" 형태에서 C는 성취도, 195는 수강자수)
                let achievementLevel = null;
                let totalStudents = null;
                
                console.log(`성취도 데이터 (H열):`, achievement); // 디버깅용
                
                if (achievement) {
                    const achievementText = achievement.toString().trim();
                    console.log(`성취도 텍스트:`, achievementText); // 디버깅용
                    let m = achievementText.match(/([A-E])\s*\((\d+)\)/i);
                    if (m) {
                        achievementLevel = m[1].toUpperCase();
                        totalStudents = parseInt(m[2]);
                    } else {
                        // 단독 등급(A~E) 혹은 P(패스)
                        m = achievementText.match(/^([A-EP])$/i);
                        if (m) {
                            achievementLevel = m[1].toUpperCase();
                        }
                    }
                    console.log(`파싱된 성취도: 등급=${achievementLevel || ''}, 수강자수=${totalStudents || ''}`);
                }
                
                // 학생의 과목 데이터에 추가 (학기별로 구분)
                const subjectKey = `${subjectName}_${usedGrade}학년_${usedSemester}학기`;
                const subjectData = {
                    subject: subject,
                    subjectName: subjectName,
                    originalScore: originalScore,
                    subjectAverage: subjectAverage,
                    achievement: achievementLevel,
                    totalStudents: totalStudents,
                    grade: usedGrade,
                    semester: usedSemester,
                    displayName: `${subjectName}(${usedGrade}-${usedSemester})`
                };
                
                currentStudent.subjects.set(subjectKey, subjectData);
                if (currentStudent.number === 1) {
                    console.log(`*** 1번 학생 ${currentStudent.name}의 ${subjectKey} 데이터 추가:`, {
                        key: subjectKey,
                        displayName: subjectData.displayName,
                        score: subjectData.originalScore,
                        achievement: subjectData.achievement,
                        행번호: i+1
                    });
                }
                
                // 전체 과목 리스트에 추가 (학기별로 중복 제거)
                const globalSubjectKey = `${subject}_${subjectName}_${usedGrade}_${usedSemester}`;
                if (!fileData.subjects.find(s => s.key === globalSubjectKey)) {
                    fileData.subjects.push({
                        key: globalSubjectKey,
                        subject: subject,
                        name: subjectName,
                        displayName: `${subjectName}(${usedGrade}-${usedSemester})`,
                        grade: usedGrade,
                        semester: usedSemester,
                        scores: [],
                        averages: []
                    });
                }
                // 마지막 교과 갱신
                if (idx.subjCat >= 0 && row[idx.subjCat]) {
                    lastSubjectCategory = row[idx.subjCat];
                }
            }
        }
        
        // 과목별 통계 계산
        fileData.subjects.forEach(subjectInfo => {
            const scores = [];
            const averages = [];
            
            fileData.students.forEach(student => {
                // 학기별 키로 데이터 검색
                const searchKey = `${subjectInfo.name}_${subjectInfo.grade}학년_${subjectInfo.semester}학기`;
                const subjectData = student.subjects.get(searchKey);
                if (subjectData) {
                    if (subjectData.originalScore !== null) {
                        scores.push(subjectData.originalScore);
                    }
                    if (subjectData.subjectAverage !== null) {
                        averages.push(subjectData.subjectAverage);
                    }
                }
            });
            
            subjectInfo.scores = scores;
            subjectInfo.averages = averages;
            subjectInfo.classAverage = averages.length > 0 ? averages.reduce((a, b) => a + b, 0) / averages.length : 0;
            subjectInfo.scoreAverage = scores.length > 0 ? scores.reduce((a, b) => a + b, 0) / scores.length : 0;
        });
        
        // Map을 Array로 변환 (나머지 코드와의 호환성을 위해)
        fileData.studentsArray = Array.from(fileData.students.values()).map(student => {
            const studentObj = {
                number: student.number,
                name: student.name,
                grade: student.grade,
                class: student.class,
                scores: {},
                achievements: {},
                averageScore: 0,
                totalSubjects: 0
            };
            
            let totalScore = 0;
            let subjectCount = 0;
            
            student.subjects.forEach((subjectData, subjectKey) => {
                // displayName을 키로 사용하여 학기 정보를 포함
                const displayKey = subjectData.displayName || subjectKey;
                studentObj.scores[displayKey] = subjectData.originalScore || 0;
                studentObj.achievements[displayKey] = subjectData.achievement || '';
            });
            
            // 점수가 있는 과목만 평균 계산에 사용
            student.subjects.forEach((subjectData, subjectKey) => {
                if (subjectData.originalScore !== null && subjectData.originalScore > 0) {
                    totalScore += subjectData.originalScore;
                    subjectCount++;
                }
            });
            
            studentObj.averageScore = subjectCount > 0 ? totalScore / subjectCount : 0;
            // 전체 과목 수는 점수나 성취도가 있는 모든 과목
            studentObj.totalSubjects = student.subjects.size;
            
            console.log(`학생 ${student.name} 최종 데이터:`, {
                name: studentObj.name,
                totalSubjects: studentObj.totalSubjects,
                averageScore: studentObj.averageScore,
                scoreKeys: Object.keys(studentObj.scores),
                achievementKeys: Object.keys(studentObj.achievements)
            }); // 디버깅용
            
            return studentObj;
        });
        
        console.log('파싱된 파일 데이터:', fileData);
        return fileData;
    }

    combineAllData() {
        if (this.filesData.size === 0) return;

        this.combinedData = {
            subjects: [],
            students: [],
            fileNames: Array.from(this.filesData.keys())
        };

        // 모든 과목을 통합 (학기별로 중복 제거)
        const subjectMap = new Map();
        this.filesData.forEach((fileData) => {
            fileData.subjects.forEach(subject => {
                const key = subject.key; // 이미 학기 정보가 포함된 키 사용
                if (!subjectMap.has(key)) {
                    subjectMap.set(key, {
                        key: subject.key,
                        name: subject.name,
                        subject: subject.subject,
                        displayName: subject.displayName,
                        grade: subject.grade,
                        semester: subject.semester,
                        scores: [],
                        averages: [],
                        classAverages: [],
                        scoreAverages: []
                    });
                }
                
                const combinedSubject = subjectMap.get(key);
                combinedSubject.scores.push(...subject.scores);
                combinedSubject.averages.push(...subject.averages);
                if (subject.classAverage) combinedSubject.classAverages.push(subject.classAverage);
                if (subject.scoreAverage) combinedSubject.scoreAverages.push(subject.scoreAverage);
            });
        });

        // 과목별 전체 평균 계산
        subjectMap.forEach(subject => {
            subject.average = subject.averages.length > 0 ? 
                subject.averages.reduce((a, b) => a + b, 0) / subject.averages.length : 0;
            subject.scoreAverage = subject.scores.length > 0 ?
                subject.scores.reduce((a, b) => a + b, 0) / subject.scores.length : 0;
        });

        this.combinedData.subjects = Array.from(subjectMap.values());
        console.log('통합된 과목 목록:', this.combinedData.subjects.map(s => s.displayName || s.name));

        // 모든 학생을 통합
        const studentMap = new Map();
        this.filesData.forEach((fileData) => {
            fileData.studentsArray.forEach(student => {
                const key = `${student.grade}-${student.class}-${student.number}`;
                if (!studentMap.has(key)) {
                    studentMap.set(key, { ...student });
                } else {
                    // 같은 학생의 데이터가 여러 파일에 있는 경우 병합
                    const existingStudent = studentMap.get(key);
                    Object.assign(existingStudent.scores, student.scores);
                    Object.assign(existingStudent.achievements, student.achievements);
                    
                    // 평균 점수 재계산
                    const scores = Object.values(existingStudent.scores).filter(score => typeof score === 'number' && score > 0);
                    existingStudent.averageScore = scores.length > 0 ? scores.reduce((a, b) => a + b, 0) / scores.length : 0;
                    existingStudent.totalSubjects = scores.length;
                }
            });
        });

        this.combinedData.students = Array.from(studentMap.values());
        
        // 학년별, 반별 정보 추출
        this.combinedData.grades = [...new Set(this.combinedData.students.map(s => s.grade))].sort();
        this.combinedData.classes = [...new Set(this.combinedData.students.map(s => `${s.grade}-${s.class}`))].sort();

        console.log('통합된 데이터:', this.combinedData);
    }

    displayResults() {
        const resultsSection = document.getElementById('results');
        const uploadSection = document.querySelector('.upload-section');
        
        if (resultsSection) resultsSection.style.display = 'block';
        if (uploadSection) uploadSection.style.display = 'none';

        this.displayGradeDistribution();
        this.displayStudentAnalysis();
    }


    scoreToGrade(score) {
        if (score >= 90) return 'A';
        if (score >= 80) return 'B';
        if (score >= 70) return 'C';
        if (score >= 60) return 'D';
        return 'E';
    }

    displayGradeDistribution() {
        // 성취도 분포 차트 구현
        this.createScatterChart();
        this.updateGradeStats();
    }

    createScatterChart() {
        const canvas = document.getElementById('scatterChart');
        if (!canvas || !this.combinedData) return;

        const ctx = canvas.getContext('2d');
        
        // 기존 차트가 있으면 제거
        if (this.scatterChart) {
            this.scatterChart.destroy();
        }

        const students = this.combinedData.students;
        // 반(클래스)별 평균점수 순위 계산
        const classGroups = new Map(); // key: grade-class, value: array of {key, avg}
        students.forEach(s => {
            const key = `${s.grade}-${s.class}`;
            if (!classGroups.has(key)) classGroups.set(key, []);
            classGroups.get(key).push({ key: `${s.grade}-${s.class}-${s.number}`, avg: s.averageScore || 0 });
        });
        const classRanks = new Map(); // key: studentKey, value: rank (1-based)
        classGroups.forEach(arr => {
            arr.sort((a,b) => (b.avg||0) - (a.avg||0));
            arr.forEach((item, idx) => {
                classRanks.set(item.key, idx + 1);
            });
        });

        // x: 평균점수(0~100), y: 겹침 방지를 위한 작은 난수 지터(거의 일렬)
        const data = students.map((student) => {
            const avg = typeof student.averageScore === 'number' ? student.averageScore : 0;
            const jitter = (Math.random() - 0.5) * 0.04; // -0.02 ~ 0.02 범위
            const sKey = `${student.grade}-${student.class}-${student.number}`;
            const rank = classRanks.get(sKey) || null;
            return { 
                x: avg, 
                y: jitter,
                grade: student.grade,
                class: student.class,
                number: student.number,
                name: student.name,
                avg: avg,
                classRank: rank
            };
        });

        this.scatterChart = new Chart(ctx, {
            type: 'scatter',
            data: {
                datasets: [{
                    label: '학생별 평균점수',
                    data: data,
                    parsing: { xAxisKey: 'x', yAxisKey: 'y' },
                    backgroundColor: 'rgba(54, 162, 235, 0.7)',
                    borderColor: 'rgba(54, 162, 235, 1)',
                    pointRadius: 5,
                    pointHoverRadius: 8,
                    pointHitRadius: 20,
                    pointBorderWidth: 1,
                    clip: 0
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                normalized: true,
                interaction: { mode: 'nearest', axis: 'x', intersect: false },
                hover: { mode: 'nearest', intersect: false },
                animation: false,
                scales: {
                    x: {
                        title: { display: true, text: '평균점수' },
                        min: 0,
                        max: 100
                    },
                    y: {
                        title: { display: false, text: '' },
                        min: -0.06,
                        max: 0.06,
                        ticks: { display: false },
                        grid: { display: false }
                    }
                },
                plugins: { 
                    legend: { display: false },
                    tooltip: {
                        callbacks: {
                            label: (ctx) => {
                                const r = ctx.raw || {};
                                const rankText = r.classRank ? `, 반내 ${r.classRank}위` : '';
                                const nameText = r.name ? ` ${r.name}` : '';
                                return `${r.grade}학년 ${r.class}반 ${r.number}번${nameText}${rankText} — ${r.avg?.toFixed ? r.avg.toFixed(1) : r.avg}점`;
                            }
                        }
                    }
                }
            }
        });
    }

    updateGradeStats() {
        if (!this.combinedData) return;

        const students = this.combinedData.students;
        const scores = students.map(s => s.averageScore || 0).filter(score => score > 0);
        
        if (scores.length === 0) return;

        const avgScore = scores.reduce((a, b) => a + b, 0) / scores.length;
        const stdDev = Math.sqrt(scores.reduce((sum, score) => sum + Math.pow(score - avgScore, 2), 0) / scores.length);
        const bestScore = Math.max(...scores);
        const worstScore = Math.min(...scores);

        // Update UI
        const overallAvg = document.getElementById('overallAverage');
        const standardDev = document.getElementById('standardDeviation');
        const bestGrade = document.getElementById('bestGrade');
        const worstGrade = document.getElementById('worstGrade');

        if (overallAvg) overallAvg.textContent = avgScore.toFixed(2) + '점';
        if (standardDev) standardDev.textContent = stdDev.toFixed(2);
        if (bestGrade) bestGrade.textContent = bestScore.toFixed(1) + '점';
        if (worstGrade) worstGrade.textContent = worstScore.toFixed(1) + '점';
    }

    displayStudentAnalysis() {
        this.populateStudentSelectors();
        this.displayStudentCards();
    }

    populateStudentSelectors() {
        if (!this.combinedData) return;

        const gradeSelect = document.getElementById('gradeSelect');
        const classSelect = document.getElementById('classSelect');

        // 학년 옵션 추가
        if (gradeSelect) {
            gradeSelect.innerHTML = '<option value="">전체</option>';
            this.combinedData.grades.forEach(grade => {
                gradeSelect.innerHTML += `<option value="${grade}">${grade}학년</option>`;
            });
        }

        this.updateClassOptions();
    }

    updateClassOptions() {
        if (!this.combinedData) return;

        const gradeSelect = document.getElementById('gradeSelect');
        const classSelect = document.getElementById('classSelect');
        
        if (!gradeSelect || !classSelect) return;

        const selectedGrade = gradeSelect.value;
        
        classSelect.innerHTML = '<option value="">전체</option>';
        
        const availableClasses = selectedGrade 
            ? this.combinedData.students.filter(s => s.grade.toString() === selectedGrade).map(s => s.class)
            : this.combinedData.students.map(s => s.class);
            
        const uniqueClasses = [...new Set(availableClasses)].sort((a, b) => a - b);
        
        uniqueClasses.forEach(cls => {
            classSelect.innerHTML += `<option value="${cls}">${cls}반</option>`;
        });
    }

    updateStudentOptions() {
        if (!this.combinedData) return;

        const gradeSelect = document.getElementById('gradeSelect');
        const classSelect = document.getElementById('classSelect');
        const studentSelect = document.getElementById('studentSelect');
        
        if (!studentSelect) return;

        const selectedGrade = gradeSelect ? gradeSelect.value : '';
        const selectedClass = classSelect ? classSelect.value : '';
        
        let filteredStudents = this.combinedData.students;
        
        if (selectedGrade) {
            filteredStudents = filteredStudents.filter(s => s.grade.toString() === selectedGrade);
        }
        if (selectedClass) {
            filteredStudents = filteredStudents.filter(s => s.class.toString() === selectedClass);
        }
        
        studentSelect.innerHTML = '<option value="">학생 선택</option>';
        filteredStudents.forEach(student => {
            const value = `${student.grade}-${student.class}-${student.number}`;
            studentSelect.innerHTML += `<option value="${value}">${student.number}번 ${student.name}</option>`;
        });
    }

    searchStudentByName(searchTerm) {
        if (!this.combinedData || !searchTerm.trim()) {
            this.updateStudentOptions();
            return;
        }

        const studentSelect = document.getElementById('studentSelect');
        if (!studentSelect) return;

        const filteredStudents = this.combinedData.students.filter(student => 
            student.name.includes(searchTerm.trim())
        );

        studentSelect.innerHTML = '<option value="">학생 선택</option>';
        filteredStudents.forEach(student => {
            const value = `${student.grade}-${student.class}-${student.number}`;
            studentSelect.innerHTML += `<option value="${value}">${student.number}번 ${student.name}</option>`;
        });
    }

    displayStudentCards() {
        const container = document.getElementById('studentTable');
        if (!container || !this.combinedData) return;

        const students = this.combinedData.students;
        
        if (students.length === 0) {
            container.innerHTML = '<p>학생 데이터가 없습니다.</p>';
            return;
        }

        // 학생별 카드 생성
        let cardsHtml = '<div class="student-cards-container">';
        
        students.forEach(student => {
            // 학기별로 성적 그룹화
            const semesterGroups = this.groupSubjectsBySemester(student);
            
            cardsHtml += `
                <div class="student-card" data-student="${student.grade}-${student.class}-${student.number}">
                    <div class="student-card-header">
                        <div class="student-info">
                            <h3>${student.grade}학년 ${student.class}반 ${student.number}번</h3>
                            <h4>${student.name}</h4>
                        </div>
                        <div class="student-stats">
                            <div class="stat-item">
                                <span class="stat-label">평균</span>
                                <span class="stat-value">${student.averageScore ? student.averageScore.toFixed(1) : '0.0'}점</span>
                            </div>
                            <div class="stat-item">
                                <span class="stat-label">응시과목</span>
                                <span class="stat-value">${student.totalSubjects}개</span>
                            </div>
                        </div>
                    </div>
                    <div class="student-card-body">
            `;
            
            // 학기별 성적 표시
            Object.keys(semesterGroups).sort().forEach(semesterKey => {
                const [grade, semester] = semesterKey.split('-');
                const subjects = semesterGroups[semesterKey];
                
                if (subjects.length > 0) {
                    cardsHtml += `
                        <div class="semester-group">
                            <h5>${grade}학년 ${semester}학기</h5>
                            <div class="subjects-grid">
                    `;
                    
                    subjects.forEach(subject => {
                        const score = subject.score > 0 ? subject.score : '';
                        const achievement = subject.achievement || '';
                        const achievementClass = achievement ? `achievement-${achievement.toLowerCase()}` : '';
                        
                        cardsHtml += `
                            <div class="subject-item ${achievementClass}">
                                <div class="subject-name">${subject.name}</div>
                                <div class="subject-score">
                                    ${score && achievement ? `${score}점 (${achievement})` : 
                                      score ? `${score}점` : 
                                      achievement ? `${achievement}` : ''}
                                </div>
                            </div>
                        `;
                    });
                    
                    cardsHtml += `
                            </div>
                        </div>
                    `;
                }
            });
            
            cardsHtml += `
                    </div>
                </div>
            `;
        });
        
        cardsHtml += '</div>';
        container.innerHTML = cardsHtml;
    }

    groupSubjectsBySemester(student) {
        const groups = {};
        
        // 학생의 모든 과목을 학기별로 그룹화
        Object.keys(student.scores).forEach(subjectKey => {
            const score = student.scores[subjectKey];
            const achievement = student.achievements[subjectKey];
            
            // displayName에서 과목명과 학기 정보 추출
            // 예: "국어(1-1)" -> 과목: "국어", 학년: "1", 학기: "1"
            const match = subjectKey.match(/^(.+)\((\d+)-(\d+)\)$/);
            if (match) {
                const [, subjectName, grade, semester] = match;
                const semesterKey = `${grade}-${semester}`;
                
                if (!groups[semesterKey]) {
                    groups[semesterKey] = [];
                }
                
                groups[semesterKey].push({
                    name: subjectName,
                    score: score,
                    achievement: achievement
                });
            }
        });
        
        return groups;
    }

    filterStudentTable(searchTerm) {
        const cards = document.querySelectorAll('.student-card');
        if (!cards.length) return;

        const term = searchTerm.toLowerCase().trim();

        cards.forEach(card => {
            const studentInfo = card.querySelector('.student-info');
            if (studentInfo) {
                const text = studentInfo.textContent.toLowerCase();
                
                if (text.includes(term)) {
                    card.style.display = '';
                } else {
                    card.style.display = 'none';
                }
            }
        });
    }

    switchTab(tabName) {
        // Remove active class from all tabs and contents
        document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
        document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));

        // Add active class to selected tab and content
        document.querySelector(`[data-tab="${tabName}"]`).classList.add('active');
        document.getElementById(`${tabName}-tab`).classList.add('active');
    }

    switchView(viewType) {
        const tableViewBtn = document.getElementById('tableViewBtn');
        const detailViewBtn = document.getElementById('detailViewBtn');
        const tableView = document.getElementById('tableView');
        const detailView = document.getElementById('detailView');

        if (viewType === 'table') {
            tableViewBtn.classList.add('active');
            detailViewBtn.classList.remove('active');
            tableView.style.display = 'block';
            detailView.style.display = 'none';
        } else {
            tableViewBtn.classList.remove('active');
            detailViewBtn.classList.add('active');
            tableView.style.display = 'none';
            detailView.style.display = 'block';
        }
    }

    showStudentDetailView(studentKey) {
        if (!this.combinedData) return;

        const student = this.combinedData.students.find(s => 
            `${s.grade}-${s.class}-${s.number}` === studentKey
        );

        if (!student) return;

        const container = document.getElementById('studentDetailContent');
        if (!container) return;

        let detailHtml = `
            <div class="student-detail-header">
                <h3>${student.grade}학년 ${student.class}반 ${student.number}번 ${student.name}</h3>
                <div class="student-summary">
                    <span>평균점수: ${student.averageScore ? student.averageScore.toFixed(1) : '0.0'}점</span>
                    <span>응시과목: ${student.totalSubjects}과목</span>
                </div>
            </div>
            <h4 style="margin-top:10px; margin-bottom:10px; color:#2c3e50;">교과별 성취 요약</h4>
            <div class="category-grid" id="categoryGrid"></div>
        `;
        container.innerHTML = detailHtml;

        // 교과별 카드 + 미니 차트 렌더링
        const stats = this.computeStudentCategoryStats(student);
        this.renderStudentCategoryGrid(stats);

        this.switchView('detail');
    }

    computeStudentCategoryStats(student) {
        const categories = new Map(); // cat -> { terms: Map(term->values[]), overall: [] }
        const termOrder = (g, s) => (parseInt(g) - 1) * 2 + parseInt(s);

        (this.combinedData.subjects || []).forEach(subj => {
            const key = subj.displayName || subj.name;
            const score = student.scores[key];
            const ach = student.achievements[key];
            const achUpper = ach ? ach.toString().toUpperCase() : '';
            const cat = subj.subject || '기타';
            // term label
            let label = '';
            const m = key.match(/\((\d+)-(\d+)\)$/);
            if (m) label = `${m[1]}-${m[2]}`;
            if (!label) return;
            if (!categories.has(cat)) categories.set(cat, { terms: new Map(), overall: [], termKeys: new Set(), items: [] });
            const c = categories.get(cat);
            // 평균 집계: P 제외 + 점수>0
            if (achUpper !== 'P' && typeof score === 'number' && score > 0) {
                if (!c.terms.has(label)) c.terms.set(label, []);
                c.terms.get(label).push(score);
                c.overall.push(score);
                c.termKeys.add(label);
            } else {
                c.termKeys.add(label); // 항목 표시는 위해 학기 레이블만 보존
            }
            // 항목 목록: 점수 또는 성취도 존재 시 포함(P 포함)
            if ((typeof score === 'number' && score >= 0) || achUpper) {
                c.items.push({
                    display: `${subj.name}(${label})`,
                    score: (typeof score === 'number' && score > 0) ? score : null,
                    achievement: achUpper || null
                });
            }
        });

        // finalize stats
        const result = [];
        categories.forEach((c, cat) => {
            const labels = Array.from(c.termKeys).sort((a, b) => {
                const [ag, as] = a.split('-');
                const [bg, bs] = b.split('-');
                return termOrder(ag, as) - termOrder(bg, bs);
            });
            const series = labels.map(lab => {
                const vals = c.terms.get(lab) || [];
                if (!vals.length) return null;
                return vals.reduce((a,b)=>a+b,0)/vals.length;
            });
            const overallAvg = c.overall.length ? c.overall.reduce((a,b)=>a+b,0)/c.overall.length : 0;
            let changePct = null;
            const first = series.find(v => typeof v === 'number');
            const last = [...series].reverse().find(v => typeof v === 'number');
            if (typeof first === 'number' && typeof last === 'number' && first > 0) {
                changePct = ((last - first) / first) * 100;
            }
            const subjectCount = c.overall.length;
            // 아이템 정렬(학기 -> 이름)
            const items = (c.items || []).sort((a, b) => {
                const am = a.display.match(/\((\d+)-(\d+)\)$/);
                const bm = b.display.match(/\((\d+)-(\d+)\)$/);
                const ao = am ? termOrder(am[1], am[2]) : 0;
                const bo = bm ? termOrder(bm[1], bm[2]) : 0;
                if (ao !== bo) return ao - bo;
                return a.display.localeCompare(b.display, 'ko');
            });
            result.push({ category: cat, labels, series, overallAvg, changePct, subjectCount, items });
        });
        return result;
    }

    renderStudentCategoryGrid(stats) {
        const grid = document.getElementById('categoryGrid');
        if (!grid) return;
        if (!stats || !stats.length) {
            grid.innerHTML = '<div class="no-grade-notice"><span>교과 요약 데이터가 없습니다.</span></div>';
            return;
        }
        grid.innerHTML = stats.map((s, i) => {
            const delta = s.changePct;
            const cls = delta == null ? 'neutral' : (delta >= 0 ? 'positive' : 'negative');
            const deltaText = delta == null ? '변화율 없음' : `${delta >= 0 ? '▲' : '▼'} ${Math.abs(delta).toFixed(1)}%`;
            const count = s.subjectCount || s.series.filter(v => typeof v === 'number').length;
            const hasChart = Array.isArray(s.series) && s.series.some(v => typeof v === 'number');
            return `
                <div class="category-card">
                    <div class="title">${s.category}</div>
                    <div class="row"><span>전체 평균</span><strong>${s.overallAvg.toFixed(1)}점</strong></div>
                    <div class="row"><span>과목 수</span><span>${count}개</span></div>
                    <div class="row"><span>변화율</span><span class="change-badge ${cls}">${deltaText}</span></div>
                    ${hasChart ? `<div class="chart"><canvas id="catChart-${i}"></canvas></div>` : ''}
                    <div class="items">
                        <table class="mini-table">
                            <thead>
                                <tr>
                                    <th>과목(학기)</th>
                                    <th style="text-align:right">점수</th>
                                    <th style="text-align:center">성취도</th>
                                </tr>
                            </thead>
                            <tbody>
                                ${s.items.map(it => `
                                    <tr>
                                        <td>${it.display}</td>
                                        <td class="score-cell">${it.score != null ? `${it.score}` : '-'}</td>
                                        <td class="ach-cell">${it.achievement ? `<span class=\"achievement ${it.achievement}\">${it.achievement}</span>` : '-'}</td>
                                    </tr>
                                `).join('')}
                            </tbody>
                        </table>
                    </div>
                </div>
            `;
        }).join('');

        // 미니 차트 렌더링
        this._categoryCharts = this._categoryCharts || [];
        // 기존 차트 정리
        this._categoryCharts.forEach(ch => { try { ch.destroy(); } catch(_) {} });
        this._categoryCharts = [];

        stats.forEach((s, i) => {
            const canvas = document.getElementById(`catChart-${i}`);
            if (!canvas) return;
            const ctx = canvas.getContext('2d');
            const chart = new Chart(ctx, {
                type: 'bar',
                data: {
                    labels: s.labels,
                    datasets: [{
                        label: `${s.category}`,
                        data: s.series,
                        backgroundColor: 'rgba(79, 172, 254, 0.7)',
                        borderColor: 'rgba(79, 172, 254, 1)',
                        borderWidth: 1,
                        borderRadius: 4,
                        barPercentage: 0.7,
                        categoryPercentage: 0.8
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: {
                        y: { min: 0, max: 100, ticks: { display: true, stepSize: 20 }, grid: { display: false } },
                        x: { ticks: { display: true, maxRotation: 0 }, grid: { display: false } }
                    },
                    plugins: { legend: { display: false }, tooltip: { enabled: true } }
                }
            });
            this._categoryCharts.push(chart);
        });
    }

    exportClassPDF() {
        // PDF 내보내기 기능 구현
        alert('PDF 내보내기 기능은 개발 중입니다.');
    }
}

// 페이지 로드 시 ScoreAnalyzer 초기화
document.addEventListener('DOMContentLoaded', () => {
    new ScoreAnalyzer();
});
