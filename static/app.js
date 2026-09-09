// Smart Excel Merger Frontend JS - Pro Edition (With Undo, Progress & Blank Template Export)
document.addEventListener('DOMContentLoaded', () => {
    // State management
    const state = {
        selectedFiles: [],
        analysisResult: null,
        customRules: loadRulesFromStorage(),
        templates: loadTemplatesFromStorage(),
        similarityThreshold: 0.75,
        ignoredClusters: [],
        historyStack: [],
        dedupCols: [],
        cleanRules: {},
        activeCleanTarget: null,
        activeCleanQualityInfo: null,
        downloadId: null,
        templateDownloadId: null,
        activeStep: 1
    };

    // Helper: LocalStorage
    function loadRulesFromStorage() {
        try {
            const saved = localStorage.getItem('smart_excel_custom_rules');
            return saved ? JSON.parse(saved) : [];
        } catch (e) {
            return [];
        }
    }

    function saveRulesToStorage() {
        try {
            localStorage.setItem('smart_excel_custom_rules', JSON.stringify(state.customRules));
        } catch (e) {
            console.error('保存自定义规则失败:', e);
        }
    }

    function loadTemplatesFromStorage() {
        try {
            const saved = localStorage.getItem('smart_excel_templates');
            return saved ? JSON.parse(saved) : {};
        } catch (e) {
            return {};
        }
    }

    function saveTemplatesToStorage() {
        try {
            localStorage.setItem('smart_excel_templates', JSON.stringify(state.templates));
        } catch (e) {
            console.error('保存模板失败:', e);
        }
    }

    function saveLastConfig(selectedCols) {
        try {
            const config = {
                customRules: state.customRules,
                similarityThreshold: state.similarityThreshold,
                selectedCols: selectedCols || []
            };
            localStorage.setItem('smart_excel_last_config', JSON.stringify(config));
        } catch (e) {
            console.error('保存最后运行配置失败:', e);
        }
    }

    function loadLastConfig() {
        try {
            const saved = localStorage.getItem('smart_excel_last_config');
            return saved ? JSON.parse(saved) : null;
        } catch (e) {
            return null;
        }
    }

    // DOM Elements
    const quickStartCard = document.getElementById('quick-start-card');
    const loadLastConfigBtn = document.getElementById('load-last-config-btn');
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');
    const browseBtn = document.getElementById('browse-btn');
    const filesSummaryBar = document.getElementById('files-summary-bar');
    const fileCount = document.getElementById('file-count');
    const filesGrid = document.getElementById('files-grid');
    const clearFilesBtn = document.getElementById('clear-files-btn');
    const gotoStep2Btn = document.getElementById('goto-step-2-btn');

    // Progress Bar DOM
    const parseProgressWrapper = document.getElementById('parse-progress-wrapper');
    const parseProgressText = document.getElementById('parse-progress-text');
    const parseProgressPercent = document.getElementById('parse-progress-percent');
    const parseProgressFill = document.getElementById('parse-progress-fill');

    // Step 2 elements
    const sidebarFileCount = document.getElementById('sidebar-file-count');
    const sidebarFilesList = document.getElementById('sidebar-files-list');
    const templateSelect = document.getElementById('template-select');
    const applyTemplateBtn = document.getElementById('apply-template-btn');
    const saveTemplateBtn = document.getElementById('save-template-btn');
    const exportTemplateStep2Btn = document.getElementById('export-template-step2-btn');
    const similaritySlider = document.getElementById('similarity-slider');
    const sliderVal = document.getElementById('slider-val');
    const coverageBadge = document.getElementById('coverage-badge');
    const reanalyzeBtn = document.getElementById('reanalyze-btn');
    const clustersContainer = document.getElementById('clusters-container');
    const addRuleBtn = document.getElementById('add-rule-btn');
    const backStep1Btn = document.getElementById('back-step-1-btn');
    const gotoStep3Btn = document.getElementById('goto-step-3-btn');
    const ignoredColsBadge = document.getElementById('ignored-cols-badge');
    const ignoredCount = document.getElementById('ignored-count');
    const undoBtn = document.getElementById('undo-btn');

    // Rules Panel DOM
    const rulesCount = document.getElementById('rules-count');
    const clearAllRulesBtn = document.getElementById('clear-all-rules-btn');
    const rulesListGrid = document.getElementById('rules-list-grid');

    // Step 3 elements
    const dedupColsGrid = document.getElementById('dedup-cols-grid');
    const dedupActiveSummary = document.getElementById('dedup-active-summary');
    const clearDedupColsBtn = document.getElementById('clear-dedup-cols-btn');
    const dedupKeepSelect = document.getElementById('dedup-keep-select');
    const columnsSelectorGrid = document.getElementById('columns-selector-grid');
    const selectAllColsBtn = document.getElementById('select-all-cols-btn');
    const invertColsBtn = document.getElementById('invert-cols-btn');
    const deselectAllColsBtn = document.getElementById('deselect-all-cols-btn');
    const colSearchInput = document.getElementById('col-search-input');
    const backStep2Btn = document.getElementById('back-step-2-btn');
    const startMergeBtn = document.getElementById('start-merge-btn');

    // Step 4 elements
    const mergeProgressPanel = document.getElementById('merge-progress-panel');
    const mergeProgressText = document.getElementById('merge-progress-text');
    const mergeProgressPercent = document.getElementById('merge-progress-percent');
    const mergeProgressFill = document.getElementById('merge-progress-fill');
    const mergeProgressDetail = document.getElementById('merge-progress-detail');
    const statsCardsRow = document.getElementById('stats-cards-row');
    const statTotalRows = document.getElementById('stat-total-rows');
    const statDedupRemoved = document.getElementById('stat-dedup-removed');
    const statTotalCols = document.getElementById('stat-total-cols');
    const statTotalFiles = document.getElementById('stat-total-files');
    const previewThead = document.getElementById('preview-thead');
    const previewTbody = document.getElementById('preview-tbody');
    const exportTemplateStep4Btn = document.getElementById('export-template-step4-btn');
    const downloadBtn = document.getElementById('download-btn');
    const downloadBtn2 = document.getElementById('download-btn-2');
    const restartBtn = document.getElementById('restart-btn');

    // Modal elements
    const ruleModal = document.getElementById('rule-modal');
    const modalTargetInput = document.getElementById('modal-target-input');
    const modalAliasesInput = document.getElementById('modal-aliases-input');
    const modalCancelBtn = document.getElementById('modal-cancel-btn');
    const modalSaveBtn = document.getElementById('modal-save-btn');

    // Visual Cleaning Canvas Modal DOM Elements
    const cleanCanvasModal = document.getElementById('clean-canvas-modal');
    const cleanTargetName = document.getElementById('clean-target-name');
    const cleanModalCloseBtn = document.getElementById('clean-modal-close-btn');
    const cleanMaskSelect = document.getElementById('clean-mask-select');
    const cleanDateSelect = document.getElementById('clean-date-select');
    const cleanTrimCb = document.getElementById('clean-trim-cb');
    const cleanLinebreaksCb = document.getElementById('clean-linebreaks-cb');
    const cleanCaseSelect = document.getElementById('clean-case-select');
    const cleanFillEmptyInput = document.getElementById('clean-fill-empty-input');
    const sampleContrastContainer = document.getElementById('sample-contrast-container');
    const cleanSaveRuleBtn = document.getElementById('clean-save-rule-btn');
    const cleanClearRuleBtn = document.getElementById('clean-clear-rule-btn');
    const cleanCancelBtn = document.getElementById('clean-cancel-btn');

    // Initialize UI
    renderRulesList();
    renderTemplateOptions();
    checkQuickStart();
    updateUndoBtnState();

    // Keydown Listener for Ctrl+Z / Cmd+Z Undo
    document.addEventListener('keydown', (e) => {
        if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') {
            if (state.activeStep === 2 && state.historyStack.length > 0) {
                e.preventDefault();
                undoLastAction();
            }
        }
    });

    function pushHistorySnapshot() {
        if (state.analysisResult && state.analysisResult.suggested_clusters) {
            const snapshot = {
                clusters: JSON.parse(JSON.stringify(state.analysisResult.suggested_clusters)),
                ignoredClusters: [...state.ignoredClusters]
            };
            state.historyStack.push(snapshot);
            updateUndoBtnState();
        }
    }

    function undoLastAction() {
        if (state.historyStack.length > 0) {
            const lastState = state.historyStack.pop();
            state.analysisResult.suggested_clusters = lastState.clusters;
            state.ignoredClusters = lastState.ignoredClusters;
            
            renderClusters(state.analysisResult.suggested_clusters);
            updateCoverageBadge(state.analysisResult);
            updateIgnoredBadge();
            updateUndoBtnState();
        }
    }

    function updateUndoBtnState() {
        if (undoBtn) {
            undoBtn.disabled = state.historyStack.length === 0;
        }
    }

    undoBtn.addEventListener('click', undoLastAction);

    function checkQuickStart() {
        const lastConfig = loadLastConfig();
        if (lastConfig && lastConfig.customRules && lastConfig.customRules.length > 0) {
            quickStartCard.style.display = 'flex';
        } else {
            quickStartCard.style.display = 'none';
        }
    }

    loadLastConfigBtn.addEventListener('click', () => {
        const lastConfig = loadLastConfig();
        if (lastConfig) {
            state.customRules = lastConfig.customRules || [];
            if (lastConfig.similarityThreshold) {
                state.similarityThreshold = lastConfig.similarityThreshold;
                similaritySlider.value = Math.round(lastConfig.similarityThreshold * 100);
                sliderVal.textContent = `${similaritySlider.value}%`;
            }
            saveRulesToStorage();
            renderRulesList();
            alert('✅ 已成功一键恢复上次任务的规则与配置！请选择或拖入要合并的文件。');
        }
    });

    // Dropdown Menu explicit click & outside click handling
    const navToolsDropdown = document.getElementById('nav-tools-dropdown');
    const navBtnTools = document.getElementById('nav-btn-tools');
    const navToolsMenu = document.getElementById('nav-tools-menu');

    if (navBtnTools && navToolsMenu) {
        // Guarantee hidden on initial script load
        navToolsMenu.style.setProperty('display', 'none', 'important');
        if (navToolsDropdown) navToolsDropdown.classList.remove('open');

        navBtnTools.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            const isHidden = navToolsMenu.style.display === 'none' || getComputedStyle(navToolsMenu).display === 'none';
            if (isHidden) {
                navToolsMenu.style.setProperty('display', 'flex', 'important');
                navToolsMenu.style.flexDirection = 'column';
                if (navToolsDropdown) navToolsDropdown.classList.add('open');
            } else {
                navToolsMenu.style.setProperty('display', 'none', 'important');
                if (navToolsDropdown) navToolsDropdown.classList.remove('open');
            }
        });

        document.addEventListener('click', (e) => {
            if (navToolsDropdown && !navToolsDropdown.contains(e.target)) {
                navToolsMenu.style.setProperty('display', 'none', 'important');
                navToolsDropdown.classList.remove('open');
            }
        });
    }

    function closeDropdownMenu() {
        if (navToolsMenu) navToolsMenu.style.setProperty('display', 'none', 'important');
        if (navToolsDropdown) navToolsDropdown.classList.remove('open');
    }

    // Page View Switcher
    function switchPageView(pageId) {
        closeDropdownMenu();

        document.querySelectorAll('.page-view').forEach(view => view.classList.remove('active'));
        document.querySelectorAll('.nav-btn').forEach(btn => btn.classList.remove('active'));
        document.querySelectorAll('.dropdown-item').forEach(item => item.classList.remove('active'));

        const targetPage = document.getElementById(`page-${pageId}`);
        if (targetPage) targetPage.classList.add('active');

        const targetNav = document.getElementById(`nav-btn-${pageId}`);
        if (targetNav) {
            targetNav.classList.add('active');
        } else {
            if (navBtnTools && ['converter', 'merger', 'cleaner', 'compressor', 'splitter'].includes(pageId)) {
                navBtnTools.classList.add('active');
            }
        }

        const targetDropdownItem = document.getElementById(`nav-item-${pageId}`);
        if (targetDropdownItem) targetDropdownItem.classList.add('active');
        
        // Refresh portal stats
        if (pageId === 'home') {
            const tplCount = Object.keys(state.templates).length;
            const portalTplEl = document.getElementById('portal-templates-count');
            if (portalTplEl) portalTplEl.textContent = tplCount;
        }
    }

    // Platform Navigation Listeners
    document.getElementById('nav-brand-logo').addEventListener('click', () => switchPageView('home'));
    document.getElementById('nav-btn-home').addEventListener('click', () => switchPageView('home'));
    
    // Tools Dropdown & Direct Items Listeners
    ['nav-item-converter', 'nav-btn-converter'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.addEventListener('click', () => switchPageView('converter'));
    });

    ['nav-item-merger', 'nav-btn-merger'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.addEventListener('click', () => switchPageView('merger'));
    });

    ['nav-item-cleaner', 'nav-btn-cleaner'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.addEventListener('click', () => switchPageView('cleaner'));
    });

    ['nav-item-compressor', 'nav-btn-compressor'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.addEventListener('click', () => switchPageView('compressor'));
    });

    ['nav-item-splitter', 'nav-btn-splitter'].forEach(id => {
        const el = document.getElementById(id);
        if (el) {
            el.addEventListener('click', () => {
                alert('【报表按列拆分】模块为扩展预留工具，已上线工具【PDF/Word双向互转】、【50+多表智能合并】与【数据格式化清洗】均可直接使用！');
                switchPageView('merger');
            });
        }
    });

    const navBtnTemplates = document.getElementById('nav-btn-templates');
    if (navBtnTemplates) {
        navBtnTemplates.addEventListener('click', () => {
            alert('模板中心：管理并套用已保存的映射列及数据清洗字段规范。');
            switchPageView('merger');
            switchStep(2);
        });
    }

    // Home Hero & Tool Card Navigation Buttons
    const heroStartMergerBtn = document.getElementById('hero-start-merger-btn');
    if (heroStartMergerBtn) {
        heroStartMergerBtn.addEventListener('click', () => switchPageView('merger'));
    }

    const heroViewTemplatesBtn = document.getElementById('hero-view-templates-btn');
    if (heroViewTemplatesBtn) {
        heroViewTemplatesBtn.addEventListener('click', () => {
            switchPageView('merger');
            switchStep(2);
        });
    }

    const cardToolMerger = document.getElementById('card-tool-merger');
    if (cardToolMerger) {
        cardToolMerger.addEventListener('click', () => switchPageView('merger'));
    }

    const cardToolConverter = document.getElementById('card-tool-converter');
    if (cardToolConverter) {
        cardToolConverter.addEventListener('click', () => switchPageView('converter'));
    }

    const cardToolCompressor = document.getElementById('card-tool-compressor');
    if (cardToolCompressor) cardToolCompressor.addEventListener('click', () => switchPageView('compressor'));

    const compressorBackHomeBtn = document.getElementById('compressor-back-home-btn');
    if (compressorBackHomeBtn) {
        compressorBackHomeBtn.addEventListener('click', () => switchPageView('home'));
    }

    const cardToolCleaner = document.getElementById('card-tool-cleaner');
    if (cardToolCleaner) {
        cardToolCleaner.addEventListener('click', () => switchPageView('cleaner'));
    }

    const mergerBackHomeBtn = document.getElementById('merger-back-home-btn');
    if (mergerBackHomeBtn) {
        mergerBackHomeBtn.addEventListener('click', () => switchPageView('home'));
    }

    const cleanerBackHomeBtn = document.getElementById('cleaner-back-home-btn');
    if (cleanerBackHomeBtn) {
        cleanerBackHomeBtn.addEventListener('click', () => switchPageView('home'));
    }

    const converterBackHomeBtn = document.getElementById('converter-back-home-btn');
    if (converterBackHomeBtn) {
        converterBackHomeBtn.addEventListener('click', () => switchPageView('home'));
    }

    ['card-tool-splitter', 'card-tool-vlookup'].forEach(id => {
        const card = document.getElementById(id);
        if (card) {
            card.addEventListener('click', () => {
                alert('该模块为扩展预留工具，已上线工具【PDF/Word双向互转】、【50+多表智能合并】与【数据格式化清洗】均可直接使用！');
                switchPageView('converter');
            });
        }
    });

    // Navigation Step Switcher
    function switchStep(stepNumber) {
        state.activeStep = stepNumber;
        for (let i = 1; i <= 4; i++) {
            const nav = document.getElementById(`step-nav-${i}`);
            const panel = document.getElementById(`panel-step-${i}`);
            if (i === stepNumber) {
                nav.classList.add('active');
                panel.classList.add('active');
            } else {
                nav.classList.remove('active');
                panel.classList.remove('active');
            }
        }
    }

    // Template Options
    function renderTemplateOptions() {
        templateSelect.innerHTML = '<option value="">默认规则（算法对齐）</option>';
        Object.keys(state.templates).forEach(tplName => {
            const opt = document.createElement('option');
            opt.value = tplName;
            opt.textContent = `📑 ${tplName}`;
            templateSelect.appendChild(opt);
        });
    }

    saveTemplateBtn.addEventListener('click', () => {
        if (state.customRules.length === 0 && Object.keys(state.cleanRules).length === 0) {
            alert('请先配置至少一条自定义映射规则或数据清洗规则，再保存为模板！');
            return;
        }
        const tplName = prompt('请输入新规则方案模板名称（如：每日销售与脱敏报表方案）:');
        if (tplName && tplName.trim()) {
            const cleanName = tplName.trim();
            state.templates[cleanName] = {
                customRules: JSON.parse(JSON.stringify(state.customRules)),
                cleanRules: JSON.parse(JSON.stringify(state.cleanRules))
            };
            saveTemplatesToStorage();
            renderTemplateOptions();
            templateSelect.value = cleanName;
            alert(`✅ 方案模板 "${cleanName}" (包含映射对齐与清洗脱敏规则) 保存成功！`);
        }
    });

    applyTemplateBtn.addEventListener('click', () => {
        const tplName = templateSelect.value;
        if (!tplName) {
            alert('请选择要套用的预设方案模板');
            return;
        }
        const tpl = state.templates[tplName];
        if (tpl) {
            if (Array.isArray(tpl)) {
                state.customRules = JSON.parse(JSON.stringify(tpl));
                state.cleanRules = {};
            } else {
                state.customRules = JSON.parse(JSON.stringify(tpl.customRules || []));
                state.cleanRules = JSON.parse(JSON.stringify(tpl.cleanRules || {}));
            }
            saveRulesToStorage();
            renderRulesList();
            if (state.selectedFiles.length > 0) analyzeFiles();
            alert(`已成功套用方案 "${tplName}" (包含对齐规则与全套清洗脱敏配置)！`);
        }
    });

    // Export Blank Template Event Handlers
    async function triggerExportTemplate() {
        let targetHeaders = [];
        if (state.analysisResult && state.analysisResult.suggested_clusters) {
            targetHeaders = state.analysisResult.suggested_clusters.map(c => c.target);
        }
        if (targetHeaders.length === 0) {
            alert('当前没有可导出的表头结构！');
            return;
        }

        const formData = new FormData();
        formData.append('headers_json', JSON.stringify(targetHeaders));

        try {
            const resp = await fetch('/api/export-template', {
                method: 'POST',
                body: formData
            });
            const data = await resp.json();
            if (data.template_download_id) {
                window.location.href = `/api/download/${data.template_download_id}`;
            }
        } catch (err) {
            alert('导出空白模板失败: ' + err.message);
        }
    }

    exportTemplateStep2Btn.addEventListener('click', triggerExportTemplate);
    exportTemplateStep4Btn.addEventListener('click', triggerExportTemplate);

    // Step 1: File Upload Handlers
    browseBtn.addEventListener('click', () => fileInput.click());
    dropZone.addEventListener('click', (e) => {
        if (e.target !== browseBtn) fileInput.click();
    });

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
        const files = Array.from(e.dataTransfer.files).filter(f => f.name.endsWith('.xlsx') || f.name.endsWith('.xls'));
        addFiles(files);
    });

    fileInput.addEventListener('change', () => {
        const files = Array.from(fileInput.files);
        addFiles(files);
        fileInput.value = '';
    });

    function addFiles(files) {
        files.forEach(f => {
            if (!state.selectedFiles.some(existing => existing.name === f.name && existing.size === f.size)) {
                state.selectedFiles.push(f);
            }
        });
        renderFilesList();
    }

    function renderFilesList() {
        filesGrid.innerHTML = '';
        if (state.selectedFiles.length > 0) {
            filesSummaryBar.style.display = 'flex';
            fileCount.textContent = state.selectedFiles.length;
            gotoStep2Btn.disabled = false;

            state.selectedFiles.forEach((file, index) => {
                const chip = document.createElement('div');
                chip.className = 'file-item-chip';
                chip.innerHTML = `
                    <span class="file-item-name" title="${file.name}">📄 ${file.name}</span>
                    <span class="file-item-remove" data-index="${index}">&times;</span>
                `;
                filesGrid.appendChild(chip);
            });

            document.querySelectorAll('.file-item-remove').forEach(btn => {
                btn.addEventListener('click', (e) => {
                    e.stopPropagation();
                    const idx = parseInt(btn.getAttribute('data-index'));
                    state.selectedFiles.splice(idx, 1);
                    renderFilesList();
                });
            });
        } else {
            filesSummaryBar.style.display = 'none';
            gotoStep2Btn.disabled = true;
        }
    }

    clearFilesBtn.addEventListener('click', () => {
        state.selectedFiles = [];
        renderFilesList();
    });

    // Progress Bar Updater
    function updateProgressUI(percent, text) {
        parseProgressWrapper.style.display = 'block';
        parseProgressPercent.textContent = `${percent}%`;
        parseProgressFill.style.width = `${percent}%`;
        parseProgressText.textContent = text;
    }

    function hideProgressUI() {
        parseProgressWrapper.style.display = 'none';
    }

    // Step 1 -> Step 2
    gotoStep2Btn.addEventListener('click', async () => {
        if (state.selectedFiles.length === 0) return;
        state.ignoredClusters = [];
        state.historyStack = [];
        updateUndoBtnState();
        switchStep(2);
        await analyzeFiles();
    });

    // Step 2: Render Left Sidebar Files
    function renderSidebarFiles(summaries) {
        sidebarFilesList.innerHTML = '';
        sidebarFileCount.textContent = summaries.length;

        summaries.forEach(s => {
            const item = document.createElement('div');
            item.className = 'sidebar-file-item';
            item.innerHTML = `
                <div class="sidebar-file-name" title="${s.filename}">📄 ${s.filename}</div>
                <div class="sidebar-file-meta">${s.rows} 行数据 | ${s.columns_count} 个字段</div>
            `;
            sidebarFilesList.appendChild(item);
        });
    }

    // Step 2: Custom Rules List
    function renderRulesList() {
        rulesListGrid.innerHTML = '';
        rulesCount.textContent = state.customRules.length;

        if (state.customRules.length > 0) {
            clearAllRulesBtn.style.display = 'inline-block';

            state.customRules.forEach((rule, index) => {
                const card = document.createElement('div');
                card.className = 'rule-item-card';
                const aliasesStr = rule.aliases.join('、');
                card.innerHTML = `
                    <span class="rule-target-title">${rule.target}</span>
                    <span class="rule-aliases-text">包含: ${aliasesStr}</span>
                    <span class="rule-delete-btn" data-rule-idx="${index}" title="删除此规则">&times;</span>
                `;
                rulesListGrid.appendChild(card);
            });

            document.querySelectorAll('.rule-delete-btn').forEach(btn => {
                btn.addEventListener('click', (e) => {
                    e.stopPropagation();
                    const idx = parseInt(btn.getAttribute('data-rule-idx'));
                    state.customRules.splice(idx, 1);
                    saveRulesToStorage();
                    renderRulesList();
                    if (state.selectedFiles.length > 0) analyzeFiles();
                });
            });
        } else {
            clearAllRulesBtn.style.display = 'none';
            rulesListGrid.innerHTML = '<div class="empty-rules-tip">暂无自定义规则。</div>';
        }
    }

    clearAllRulesBtn.addEventListener('click', () => {
        if (confirm('确定要清空所有自定义映射规则吗？')) {
            state.customRules = [];
            saveRulesToStorage();
            renderRulesList();
            if (state.selectedFiles.length > 0) analyzeFiles();
        }
    });

    // Step 2: Similarity Slider & Analyze
    similaritySlider.addEventListener('input', () => {
        sliderVal.textContent = `${similaritySlider.value}%`;
        state.similarityThreshold = parseFloat(similaritySlider.value) / 100;
    });

    reanalyzeBtn.addEventListener('click', () => analyzeFiles());

    async function analyzeFiles() {
        clustersContainer.innerHTML = '<div class="loading-spinner">正在智能解析与读取表格真实样本数据...</div>';
        const totalCount = state.selectedFiles.length;
        updateProgressUI(20, `⏳ 正在上传并读取 ${totalCount} 个 Excel 表格...`);

        const formData = new FormData();
        state.selectedFiles.forEach(file => {
            formData.append('files', file);
        });
        formData.append('custom_rules_json', JSON.stringify(state.customRules));
        formData.append('similarity_threshold', state.similarityThreshold);

        try {
            updateProgressUI(60, `📊 正在计算 ${totalCount} 个表格的标题相似度与填充率...`);
            const resp = await fetch('/api/analyze', {
                method: 'POST',
                body: formData
            });
            updateProgressUI(90, `✨ 正在构建聚类规则与生成全空列预警...`);
            const data = await resp.json();

            if (data.error) {
                hideProgressUI();
                alert('解析失败: ' + data.error);
                return;
            }

            renderSidebarFiles(data.file_summaries);

            if (state.ignoredClusters.length > 0) {
                data.suggested_clusters = data.suggested_clusters.filter(c => !state.ignoredClusters.includes(c.target));
            }

            state.analysisResult = data;
            renderClusters(data.suggested_clusters);
            updateCoverageBadge(data);
            updateIgnoredBadge();
            updateProgressUI(100, `✅ 解析完成！包含 ${data.total_files} 个文件。`);
            setTimeout(hideProgressUI, 1200);
        } catch (err) {
            hideProgressUI();
            clustersContainer.innerHTML = `<div class="loading-spinner" style="color:#EF4444">连接服务器失败: ${err.message}</div>`;
        }
    }

    function updateCoverageBadge(data) {
        const totalRawHeaders = data.all_raw_headers ? data.all_raw_headers.length : 0;
        let mappedCount = 0;
        if (data.suggested_clusters) {
            data.suggested_clusters.forEach(c => {
                mappedCount += c.raw_headers ? c.raw_headers.length : 0;
            });
        }
        const rate = totalRawHeaders > 0 ? Math.round((mappedCount / totalRawHeaders) * 100) : 0;
        coverageBadge.textContent = `📊 列对齐覆盖率: ${rate}% (${mappedCount}/${totalRawHeaders})`;
    }

    function updateIgnoredBadge() {
        if (state.ignoredClusters.length > 0) {
            ignoredColsBadge.style.display = 'inline-block';
            ignoredCount.textContent = state.ignoredClusters.length;
        } else {
            ignoredColsBadge.style.display = 'none';
        }
    }

    function renderClusters(clusters) {
        clustersContainer.innerHTML = '';
        if (clusters.length === 0) {
            clustersContainer.innerHTML = '<div class="empty-rules-tip" style="text-align:center;padding:2rem;">所有列均已被删除或忽略</div>';
            return;
        }

        clusters.forEach((cluster, idx) => {
            const card = document.createElement('div');
            card.className = 'cluster-card';

            const tagsHtml = cluster.raw_headers.map(h => `
                <span class="alias-tag">
                    ${h}
                    <span class="tag-close-btn" data-cluster-idx="${idx}" data-tag="${h}" title="剥离此别名">&times;</span>
                </span>
            `).join('');
            
            let qualityHtml = '';
            const q = cluster.quality_info;

            // AI Recommendations badges
            let recsHtml = '';
            if (cluster.recommendations && cluster.recommendations.length > 0) {
                recsHtml = `
                    <div class="recommend-badge-group">
                        ${cluster.recommendations.map(r => `<span class="recommend-pill" title="${r.desc}">${r.label}</span>`).join('')}
                    </div>
                `;
            }

            // Active clean rule indicator
            let activeCleanBadge = '';
            const activeRule = state.cleanRules[cluster.target];
            if (activeRule) {
                let ruleDescParts = [];
                if (activeRule.mask_type && activeRule.mask_type !== 'none') ruleDescParts.push('脱敏');
                if (activeRule.date_format && activeRule.date_format !== 'none') ruleDescParts.push(`日期: ${activeRule.date_format}`);
                if (activeRule.trim_space) ruleDescParts.push('Trim');
                if (activeRule.fill_empty) ruleDescParts.push(`空值:${activeRule.fill_empty}`);
                
                const descStr = ruleDescParts.length > 0 ? ruleDescParts.join(' | ') : '有规则';
                activeCleanBadge = `<span class="clean-active-badge">✨ 已配置清洗: ${descStr}</span>`;
            }

            if (q) {
                if (q.is_all_empty) {
                    qualityHtml = `
                        <div class="empty-column-warning">
                            <span>⚠️ <strong>全空列警示:</strong> 该列在所有源表格中填充率为 <strong>0%</strong> (全部为无数据空值)，建议点击右侧按钮直接删除。</span>
                        </div>
                    `;
                } else {
                    const rateClass = q.non_null_rate >= 80 ? 'rate-pill-high' : 'rate-pill-low';
                    const sampleRowsHtml = q.samples_by_file.map(sf => {
                        const pills = sf.samples.map(v => `<span class="sample-value-pill" title="${v}">${v}</span>`).join('');
                        return `
                            <div class="source-sample-row">
                                <span class="source-file-name">📄 ${sf.filename}:</span>
                                ${pills}
                            </div>
                        `;
                    }).join('');

                    qualityHtml = `
                        <div class="quality-metrics-bar">
                            <span class="quality-pill">类型: ${q.data_type}</span>
                            <span class="${rateClass}">有效数据填充率: ${q.non_null_rate}% (${q.non_null_count}/${q.total_rows} 行)</span>
                        </div>
                        <div class="source-samples-container">
                            ${sampleRowsHtml}
                        </div>
                    `;
                }
            }

            card.innerHTML = `
                <div class="cluster-target" style="flex-direction:column;align-items:flex-start;flex:1;">
                    <input type="text" class="target-input" value="${cluster.target}" data-idx="${idx}">
                    ${qualityHtml}
                </div>
                <div class="cluster-aliases" style="align-self:flex-start;">
                    ${tagsHtml}
                </div>
                <button class="cluster-delete-btn" data-delete-idx="${idx}" style="align-self:flex-start;">🗑️ 删除此整列</button>
            `;
            clustersContainer.appendChild(card);
        });

        document.querySelectorAll('.target-input').forEach(input => {
            input.addEventListener('change', (e) => {
                const idx = parseInt(input.getAttribute('data-idx'));
                const newTarget = input.value.trim();
                if (newTarget && state.analysisResult.suggested_clusters[idx]) {
                    state.analysisResult.suggested_clusters[idx].target = newTarget;
                }
            });
        });

        // Individual Tag close listener (with Undo History Snapshot)
        document.querySelectorAll('.tag-close-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.stopPropagation();
                pushHistorySnapshot();

                const cIdx = parseInt(btn.getAttribute('data-cluster-idx'));
                const tagToRemove = btn.getAttribute('data-tag');

                if (state.analysisResult && state.analysisResult.suggested_clusters[cIdx]) {
                    const cluster = state.analysisResult.suggested_clusters[cIdx];
                    cluster.raw_headers = cluster.raw_headers.filter(h => h !== tagToRemove);

                    if (cluster.raw_headers.length === 0) {
                        state.analysisResult.suggested_clusters.splice(cIdx, 1);
                    }
                    renderClusters(state.analysisResult.suggested_clusters);
                    updateCoverageBadge(state.analysisResult);
                }
            });
        });

        // Entire cluster deletion listener (with Undo History Snapshot)
        document.querySelectorAll('.cluster-delete-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                pushHistorySnapshot();

                const idx = parseInt(btn.getAttribute('data-delete-idx'));
                const deletedTarget = state.analysisResult.suggested_clusters[idx].target;
                
                state.ignoredClusters.push(deletedTarget);
                state.analysisResult.suggested_clusters.splice(idx, 1);
                
                renderClusters(state.analysisResult.suggested_clusters);
                updateCoverageBadge(state.analysisResult);
                updateIgnoredBadge();
            });
        });
    }

    // Modal Handlers
    addRuleBtn.addEventListener('click', () => {
        modalTargetInput.value = '';
        modalAliasesInput.value = '';
        ruleModal.style.display = 'flex';
        modalTargetInput.focus();
    });

    modalCancelBtn.addEventListener('click', () => {
        ruleModal.style.display = 'none';
    });

    modalSaveBtn.addEventListener('click', () => {
        const target = modalTargetInput.value.trim();
        const aliasesStr = modalAliasesInput.value.trim();
        if (!target) {
            alert('请输入目标统一表头名称（如：客户名称）');
            return;
        }
        const aliases = aliasesStr.split(/[,，]/).map(s => s.trim()).filter(Boolean);
        
        const existingRule = state.customRules.find(r => r.target === target);
        if (existingRule) {
            aliases.forEach(a => {
                if (!existingRule.aliases.includes(a)) existingRule.aliases.push(a);
            });
        } else {
            state.customRules.push({ target, aliases });
        }

        saveRulesToStorage();
        renderRulesList();

        modalTargetInput.value = '';
        modalAliasesInput.value = '';
        ruleModal.style.display = 'none';

        if (state.selectedFiles.length > 0) {
            analyzeFiles();
        }
    });

    // Visual Cleaning Canvas Modal Handlers
    function openCleanCanvasModal(clusterTarget, qualityInfo) {
        state.activeCleanTarget = clusterTarget;
        state.activeCleanQualityInfo = qualityInfo;
        if (cleanTargetName) cleanTargetName.textContent = clusterTarget;

        const existingRule = state.cleanRules[clusterTarget] || {};
        if (cleanMaskSelect) cleanMaskSelect.value = existingRule.mask_type || 'none';
        if (cleanDateSelect) cleanDateSelect.value = existingRule.date_format || 'none';
        if (cleanTrimCb) cleanTrimCb.checked = existingRule.trim_space !== undefined ? existingRule.trim_space : true;
        if (cleanLinebreaksCb) cleanLinebreaksCb.checked = existingRule.remove_linebreaks || false;
        if (cleanCaseSelect) cleanCaseSelect.value = existingRule.case_transform || 'none';
        if (cleanFillEmptyInput) cleanFillEmptyInput.value = existingRule.fill_empty || '';

        if (cleanCanvasModal) cleanCanvasModal.style.display = 'flex';
        updateLiveContrastPreview();
    }

    function updateLiveContrastPreview() {
        if (!sampleContrastContainer) return;
        sampleContrastContainer.innerHTML = '';
        const qualityInfo = state.activeCleanQualityInfo;
        if (!qualityInfo || !qualityInfo.samples_by_file) {
            sampleContrastContainer.innerHTML = '<div class="empty-rules-tip" style="text-align:center;padding:1rem;">暂无真实数据样本</div>';
            return;
        }

        let samples = [];
        qualityInfo.samples_by_file.forEach(sf => {
            if (sf.samples) {
                sf.samples.forEach(s => {
                    if (samples.length < 6 && s !== null && s !== undefined && String(s).trim() !== "") {
                        samples.push(String(s));
                    }
                });
            }
        });

        if (samples.length === 0) {
            sampleContrastContainer.innerHTML = '<div class="empty-rules-tip" style="text-align:center;padding:1rem;">暂无有效数据样本</div>';
            return;
        }

        const maskType = cleanMaskSelect ? cleanMaskSelect.value : 'none';
        const dateFormat = cleanDateSelect ? cleanDateSelect.value : 'none';
        const trimSpace = cleanTrimCb ? cleanTrimCb.checked : true;
        const removeLinebreaks = cleanLinebreaksCb ? cleanLinebreaksCb.checked : false;
        const caseTransform = cleanCaseSelect ? cleanCaseSelect.value : 'none';
        const fillEmpty = cleanFillEmptyInput ? cleanFillEmptyInput.value.trim() : '';

        samples.forEach(rawVal => {
            let cleaned = rawVal;

            if (dateFormat && dateFormat !== 'none') {
                cleaned = mockDateTransform(cleaned, dateFormat);
            }
            if (maskType && maskType !== 'none') {
                cleaned = mockMaskTransform(cleaned, maskType);
            }
            if (trimSpace && typeof cleaned === 'string') cleaned = cleaned.trim();
            if (removeLinebreaks && typeof cleaned === 'string') cleaned = cleaned.replace(/[\r\n]+/g, ' ');
            if (caseTransform === 'upper' && typeof cleaned === 'string') cleaned = cleaned.toUpperCase();
            if (caseTransform === 'lower' && typeof cleaned === 'string') cleaned = cleaned.toLowerCase();
            if ((!cleaned || cleaned.trim() === '') && fillEmpty) cleaned = fillEmpty;

            const row = document.createElement('div');
            row.className = 'contrast-row';
            row.innerHTML = `
                <span class="contrast-before" title="${rawVal}">${rawVal}</span>
                <span class="contrast-arrow">➔</span>
                <span class="contrast-after" title="${cleaned}">${cleaned}</span>
            `;
            sampleContrastContainer.appendChild(row);
        });
    }

    function mockMaskTransform(val, maskType) {
        if (!val) return val;
        const s = String(val).trim();
        if (maskType === 'phone_3_4') {
            const digits = s.replace(/\D/g, '');
            if (digits.length === 11) return digits.slice(0, 3) + '****' + digits.slice(7);
            if (s.length >= 7) return s.slice(0, 3) + '****' + s.slice(-4);
            return s;
        } else if (maskType === 'email') {
            if (s.includes('@')) {
                const parts = s.split('@');
                const uname = parts[0];
                const domain = parts[1];
                const maskedU = uname.length <= 2 ? uname[0] + '*' : uname[0] + '***' + uname[uname.length - 1];
                return maskedU + '@' + domain;
            }
            return s;
        } else if (maskType === 'idcard') {
            if (s.length === 18) return s.slice(0, 6) + '******' + s.slice(14);
            if (s.length === 15) return s.slice(0, 6) + '******' + s.slice(12);
            return s;
        } else if (maskType === 'name') {
            if (s.length === 2) return s[0] + '*';
            if (s.length > 2) return s[0] + '*'.repeat(s.length - 2) + s[s.length - 1];
            return s;
        } else if (maskType === 'full_star') {
            return '****';
        }
        return s;
    }

    function mockDateTransform(val, dateFormat) {
        if (!val) return val;
        const s = String(val).trim();
        if (!s) return val;
        const d = new Date(s);
        if (isNaN(d.getTime())) {
            if (/^\d{8}$/.test(s)) {
                const y = s.slice(0, 4), m = s.slice(4, 6), day = s.slice(6, 8);
                if (dateFormat === 'YYYY-MM-DD') return `${y}-${m}-${day}`;
                if (dateFormat === 'YYYY/MM/DD') return `${y}/${m}/${day}`;
                if (dateFormat === 'YYYY年MM月DD日') return `${y}年${m}月${day}日`;
                if (dateFormat === 'YYYYMMDD') return `${y}${m}${day}`;
            }
            return s;
        }
        const y = d.getFullYear();
        const m = String(d.getMonth() + 1).padStart(2, '0');
        const day = String(d.getDate()).padStart(2, '0');
        if (dateFormat === 'YYYY-MM-DD') return `${y}-${m}-${day}`;
        if (dateFormat === 'YYYY/MM/DD') return `${y}/${m}/${day}`;
        if (dateFormat === 'YYYY年MM月DD日') return `${y}年${m}月${day}日`;
        if (dateFormat === 'YYYY-MM-DD HH:mm:ss') return `${y}-${m}-${day} 00:00:00`;
        if (dateFormat === 'YYYYMMDD') return `${y}${m}${day}`;
        return `${y}-${m}-${day}`;
    }

    [cleanMaskSelect, cleanDateSelect, cleanTrimCb, cleanLinebreaksCb, cleanCaseSelect, cleanFillEmptyInput].forEach(el => {
        if (el) {
            el.addEventListener('change', updateLiveContrastPreview);
            el.addEventListener('input', updateLiveContrastPreview);
        }
    });

    if (cleanModalCloseBtn) cleanModalCloseBtn.addEventListener('click', () => cleanCanvasModal.style.display = 'none');
    if (cleanCancelBtn) cleanCancelBtn.addEventListener('click', () => cleanCanvasModal.style.display = 'none');

    if (cleanSaveRuleBtn) {
        cleanSaveRuleBtn.addEventListener('click', () => {
            const target = state.activeCleanTarget;
            if (!target) return;

            const rule = {
                mask_type: cleanMaskSelect ? cleanMaskSelect.value : 'none',
                date_format: cleanDateSelect ? cleanDateSelect.value : 'none',
                trim_space: cleanTrimCb ? cleanTrimCb.checked : true,
                remove_linebreaks: cleanLinebreaksCb ? cleanLinebreaksCb.checked : false,
                case_transform: cleanCaseSelect ? cleanCaseSelect.value : 'none',
                fill_empty: cleanFillEmptyInput ? cleanFillEmptyInput.value.trim() : ''
            };

            state.cleanRules[target] = rule;
            if (cleanCanvasModal) cleanCanvasModal.style.display = 'none';

            if (state.analysisResult && state.analysisResult.suggested_clusters) {
                renderClusters(state.analysisResult.suggested_clusters);
            }
        });
    }

    if (cleanClearRuleBtn) {
        cleanClearRuleBtn.addEventListener('click', () => {
            const target = state.activeCleanTarget;
            if (target && state.cleanRules[target]) {
                delete state.cleanRules[target];
            }
            if (cleanCanvasModal) cleanCanvasModal.style.display = 'none';
            if (state.analysisResult && state.analysisResult.suggested_clusters) {
                renderClusters(state.analysisResult.suggested_clusters);
            }
        });
    }

    backStep1Btn.addEventListener('click', () => switchStep(1));

    // Step 2 -> Step 3
    gotoStep3Btn.addEventListener('click', () => {
        switchStep(3);
        renderColumnSelector();
        renderDedupOptions();
    });

    function renderDedupOptions() {
        if (!dedupColsGrid) return;
        dedupColsGrid.innerHTML = '';
        if (!state.analysisResult || !state.analysisResult.suggested_clusters) return;

        const targetCols = state.analysisResult.suggested_clusters.map(c => c.target);
        if (targetCols.length === 0) {
            dedupColsGrid.innerHTML = '<div class="empty-rules-tip">暂无可用列</div>';
            updateDedupSummary();
            return;
        }

        targetCols.forEach(col => {
            const isChecked = state.dedupCols.includes(col);
            const chip = document.createElement('label');
            chip.className = `dedup-col-chip ${isChecked ? 'active' : ''}`;
            chip.innerHTML = `
                <input type="checkbox" value="${col}" ${isChecked ? 'checked' : ''} class="dedup-col-cb">
                <span>🔑 ${col}</span>
            `;

            const cb = chip.querySelector('input');
            cb.addEventListener('change', () => {
                if (cb.checked) {
                    if (!state.dedupCols.includes(col)) state.dedupCols.push(col);
                    chip.classList.add('active');
                } else {
                    state.dedupCols = state.dedupCols.filter(c => c !== col);
                    chip.classList.remove('active');
                }
                updateDedupSummary();
            });

            dedupColsGrid.appendChild(chip);
        });

        updateDedupSummary();
    }

    function updateDedupSummary() {
        if (!dedupActiveSummary) return;
        if (!state.dedupCols || state.dedupCols.length === 0) {
            dedupActiveSummary.innerHTML = '当前去重策略: <strong>全量保留 (不去重)</strong>';
        } else if (state.dedupCols.length === 1) {
            dedupActiveSummary.innerHTML = `当前去重策略: 单列唯一主键 <strong>[ ${state.dedupCols[0]} ]</strong> 去重`;
        } else {
            dedupActiveSummary.innerHTML = `当前去重策略: <strong>${state.dedupCols.length} 列复合主键 [ ${state.dedupCols.join(' + ')} ]</strong> 联合唯一去重`;
        }
    }

    if (clearDedupColsBtn) {
        clearDedupColsBtn.addEventListener('click', () => {
            state.dedupCols = [];
            document.querySelectorAll('.dedup-col-cb').forEach(cb => {
                cb.checked = false;
                const chip = cb.closest('.dedup-col-chip');
                if (chip) chip.classList.remove('active');
            });
            updateDedupSummary();
        });
    }

    // Step 3.2: Export Retained Columns Selection (Toggle Switch Cards)
    function renderColumnSelector() {
        columnsSelectorGrid.innerHTML = '';
        if (!state.analysisResult) return;

        const targetCols = state.analysisResult.suggested_clusters.map(c => c.target);

        if (targetCols.length === 0) {
            columnsSelectorGrid.innerHTML = '<div class="empty-rules-tip">没有可合并的列，请返回上一步勾选或恢复列。</div>';
            return;
        }

        targetCols.forEach(col => {
            const card = document.createElement('div');
            card.className = 'col-switch-card active';
            card.setAttribute('data-col-name', col.toLowerCase());

            card.innerHTML = `
                <div class="col-switch-left">
                    <span class="col-switch-name" title="${col}">${col}</span>
                    <span class="col-switch-status-badge">保留</span>
                </div>
                <label class="switch-toggle-wrapper">
                    <input type="checkbox" value="${col}" checked class="col-checkbox">
                    <span class="switch-slider"></span>
                </label>
            `;

            const cb = card.querySelector('.col-checkbox');
            const badge = card.querySelector('.col-switch-status-badge');

            function updateCardState() {
                if (cb.checked) {
                    card.classList.add('active');
                    badge.textContent = '保留';
                } else {
                    card.classList.remove('active');
                    badge.textContent = '剔除';
                }
            }

            cb.addEventListener('change', updateCardState);
            card.addEventListener('click', (e) => {
                if (e.target !== cb && !e.target.classList.contains('switch-slider')) {
                    cb.checked = !cb.checked;
                    updateCardState();
                }
            });

            columnsSelectorGrid.appendChild(card);
        });
    }

    if (colSearchInput) {
        colSearchInput.addEventListener('input', () => {
            const query = colSearchInput.value.trim().toLowerCase();
            document.querySelectorAll('.col-switch-card').forEach(item => {
                const colName = item.getAttribute('data-col-name');
                if (!query || colName.includes(query)) {
                    item.style.display = 'flex';
                } else {
                    item.style.display = 'none';
                }
            });
        });
    }

    function updateAllSwitchesState() {
        document.querySelectorAll('.col-switch-card').forEach(card => {
            const cb = card.querySelector('.col-checkbox');
            const badge = card.querySelector('.col-switch-status-badge');
            if (cb.checked) {
                card.classList.add('active');
                badge.textContent = '保留';
            } else {
                card.classList.remove('active');
                badge.textContent = '剔除';
            }
        });
    }

    selectAllColsBtn.addEventListener('click', () => {
        document.querySelectorAll('.col-checkbox').forEach(cb => cb.checked = true);
        updateAllSwitchesState();
    });

    invertColsBtn.addEventListener('click', () => {
        document.querySelectorAll('.col-checkbox').forEach(cb => cb.checked = !cb.checked);
        updateAllSwitchesState();
    });

    deselectAllColsBtn.addEventListener('click', () => {
        document.querySelectorAll('.col-checkbox').forEach(cb => cb.checked = false);
        updateAllSwitchesState();
    });

    backStep2Btn.addEventListener('click', () => switchStep(2));

    function updateMergeProgressUI(percent, text, detail) {
        if (mergeProgressPanel) mergeProgressPanel.style.display = 'block';
        if (mergeProgressPercent) mergeProgressPercent.textContent = `${percent}%`;
        if (mergeProgressFill) mergeProgressFill.style.width = `${percent}%`;
        if (mergeProgressText && text) mergeProgressText.textContent = text;
        if (mergeProgressDetail && detail) mergeProgressDetail.textContent = detail;
    }

    // Step 3 -> Step 4 (Execute Merge with High-Grade Progress Animation)
    startMergeBtn.addEventListener('click', async () => {
        const selectedCols = Array.from(document.querySelectorAll('.col-checkbox:checked')).map(cb => cb.value);

        if (selectedCols.length === 0) {
            alert('请至少勾选或选择保留一列进行合并导出！');
            return;
        }

        saveLastConfig(selectedCols);

        const dedupKeep = dedupKeepSelect ? dedupKeepSelect.value : 'first';
        const dedupCols = state.dedupCols || [];
        const fileCountNum = state.selectedFiles.length || 0;

        switchStep(4);
        previewThead.innerHTML = '';
        previewTbody.innerHTML = '<tr><td colspan="10" class="loading-spinner">正在高速初始化高吞吐内存 ETL 拼接引擎...</td></tr>';
        
        if (statsCardsRow) statsCardsRow.style.opacity = '0.3';

        // Start Animated Progress Sequence
        updateMergeProgressUI(15, `⚡ [1/4] 正在读取所有 ${fileCountNum} 个源文件数据...`, '💡 包含高并发内存解析、格式统一与基础类型预判...');

        let progressStep = 15;
        const progressTimer = setInterval(() => {
            if (progressStep < 40) {
                progressStep += 5;
                updateMergeProgressUI(progressStep, `⚡ [1/4] 正在读取所有 ${fileCountNum} 个源文件数据...`, '💡 包含高并发内存解析、格式统一与基础类型预判...');
            } else if (progressStep < 70) {
                progressStep += 6;
                updateMergeProgressUI(progressStep, `🔄 [2/4] 正在智能对齐字段表头与转换数据...`, `💡 套用包含 ${state.customRules.length} 条自定义映射规则与相似度聚类...`);
            } else if (progressStep < 90) {
                progressStep += 3;
                const dedupText = dedupCols.length > 0 ? `依据 ${dedupCols.join('+')} 主键` : '全量行保留';
                updateMergeProgressUI(progressStep, `🔑 [3/4] 正在执行多维度行记录去重 (${dedupText})...`, '💡 快速遍历并过滤重复行数据...');
            }
        }, 150);

        const formData = new FormData();
        state.selectedFiles.forEach(file => {
            formData.append('files', file);
        });
        formData.append('custom_rules_json', JSON.stringify(state.customRules));
        formData.append('similarity_threshold', state.similarityThreshold);
        formData.append('selected_columns_json', JSON.stringify(selectedCols));
        formData.append('dedup_columns_json', JSON.stringify(dedupCols));
        formData.append('dedup_keep', dedupKeep);
        formData.append('clean_rules_json', JSON.stringify(state.cleanRules));

        try {
            const resp = await fetch('/api/merge', {
                method: 'POST',
                body: formData
            });
            const data = await resp.json();

            clearInterval(progressTimer);

            if (data.error) {
                updateMergeProgressUI(0, '❌ 合并失败', data.error);
                alert('合并错误: ' + data.error);
                return;
            }

            updateMergeProgressUI(95, '📊 [4/4] 正在构建 Excel 总表与自动强行展开列宽...', '💡 强行自适应扩展 ID/手机/Email/时间 等关键列宽度...');

            setTimeout(() => {
                updateMergeProgressUI(100, `✅ [完成] 已成功合并 ${fileCountNum} 个表格文件！`, `🎉 共处理 ${data.initial_rows || data.total_rows} 行数据，多维去重剔除 ${data.dedup_removed_count || 0} 行！`);
                
                state.downloadId = data.download_id;
                state.templateDownloadId = data.template_download_id;

                // Render stats
                statTotalRows.textContent = data.total_rows;
                statDedupRemoved.textContent = data.dedup_removed_count || 0;
                statTotalCols.textContent = data.total_cols;
                statTotalFiles.textContent = state.selectedFiles.length;

                if (statsCardsRow) statsCardsRow.style.opacity = '1';

                renderPreviewTable(data.preview);
            }, 400);

        } catch (err) {
            clearInterval(progressTimer);
            updateMergeProgressUI(0, '❌ 合并出错', err.message);
            previewTbody.innerHTML = `<tr><td colspan="10" style="color:#EF4444;text-align:center;">合并失败: ${err.message}</td></tr>`;
        }
    });

    function renderPreviewTable(preview) {
        previewThead.innerHTML = '';
        previewTbody.innerHTML = '';

        const trHead = document.createElement('tr');
        preview.headers.forEach(h => {
            const th = document.createElement('th');
            th.textContent = h;
            trHead.appendChild(th);
        });
        previewThead.appendChild(trHead);

        preview.rows.forEach(row => {
            const tr = document.createElement('tr');
            row.forEach(cell => {
                const td = document.createElement('td');
                td.textContent = cell !== null && cell !== undefined ? cell : '';
                tr.appendChild(td);
            });
            previewTbody.appendChild(tr);
        });
    }

    // Download handlers
    function triggerDownload() {
        if (state.downloadId) {
            window.location.href = `/api/download/${state.downloadId}`;
        } else {
            alert('尚无合并文件可供下载');
        }
    }

    downloadBtn.addEventListener('click', triggerDownload);
    downloadBtn2.addEventListener('click', triggerDownload);

    // Restart batch
    restartBtn.addEventListener('click', () => {
        state.selectedFiles = [];
        state.ignoredClusters = [];
        state.historyStack = [];
        state.downloadId = null;
        renderFilesList();
        checkQuickStart();
        switchStep(1);
    });

    // Dedicated Data Cleaner Tool Page (#page-cleaner) - Standard Unified UI
    const tabCleanAuto = document.getElementById('tab-clean-auto');
    const tabCleanCustom = document.getElementById('tab-clean-custom');
    const cleanerDropZone = document.getElementById('cleaner-drop-zone');
    const cleanerFileInput = document.getElementById('cleaner-file-input');
    const cleanerBrowseBtn = document.getElementById('cleaner-browse-btn');
    const cleanerFilesSummary = document.getElementById('cleaner-files-summary');
    const cleanerFileCount = document.getElementById('cleaner-file-count');
    const cleanerFilesGrid = document.getElementById('cleaner-files-grid');
    const cleanerClearFilesBtn = document.getElementById('cleaner-clear-files-btn');
    const cleanerWorkspace = document.getElementById('cleaner-workspace');
    const cleanerCardsWall = document.getElementById('cleaner-cards-wall');
    const cleanerToggleWallBtn = document.getElementById('cleaner-toggle-wall-btn');
    const cleanerColsCountBadge = document.getElementById('cleaner-cols-count-badge');
    const cleanerAiRecommendBtn = document.getElementById('cleaner-ai-recommend-btn');
    const cleanerProgressWrapper = document.getElementById('cleaner-progress-wrapper');
    const cleanerProgressText = document.getElementById('cleaner-progress-text');
    const cleanerProgressPercent = document.getElementById('cleaner-progress-percent');
    const cleanerProgressFill = document.getElementById('cleaner-progress-fill');
    const cleanerResetBtn = document.getElementById('cleaner-reset-btn');
    const cleanerExecuteBtn = document.getElementById('cleaner-execute-btn');
    const cleanerDownloadBtn = document.getElementById('cleaner-download-btn');

    let cleanerMode = 'auto'; // 'auto' | 'custom'
    let cleanerFiles = [];
    let cleanerAnalysis = null;
    let cleanerRules = {};
    let cleanerDownloadId = null;

    if (tabCleanAuto && tabCleanCustom) {
        tabCleanAuto.addEventListener('click', () => setCleanerMode('auto'));
        tabCleanCustom.addEventListener('click', () => setCleanerMode('custom'));
    }

    function setCleanerMode(mode) {
        cleanerMode = mode;
        if (mode === 'auto') {
            tabCleanAuto.classList.add('active');
            tabCleanCustom.classList.remove('active');
            if (cleanerCardsWall) cleanerCardsWall.style.display = 'none';
            if (cleanerToggleWallBtn) cleanerToggleWallBtn.textContent = '📁 展开字段微调墙 ▾';
        } else {
            tabCleanCustom.classList.add('active');
            tabCleanAuto.classList.remove('active');
            if (cleanerCardsWall) cleanerCardsWall.style.display = 'grid';
            if (cleanerToggleWallBtn) cleanerToggleWallBtn.textContent = '📁 收起字段微调墙 ▴';
        }
    }

    if (cleanerToggleWallBtn) {
        cleanerToggleWallBtn.addEventListener('click', () => {
            if (!cleanerCardsWall) return;
            const isHidden = cleanerCardsWall.style.display === 'none' || getComputedStyle(cleanerCardsWall).display === 'none';
            if (isHidden) {
                cleanerCardsWall.style.display = 'grid';
                cleanerToggleWallBtn.textContent = '📁 收起字段微调墙 ▴';
                if (tabCleanCustom) setCleanerMode('custom');
            } else {
                cleanerCardsWall.style.display = 'none';
                cleanerToggleWallBtn.textContent = '📁 展开字段微调墙 ▾';
            }
        });
    }

    if (cleanerBrowseBtn) cleanerBrowseBtn.addEventListener('click', () => cleanerFileInput.click());
    if (cleanerDropZone) {
        cleanerDropZone.addEventListener('click', (e) => {
            if (e.target !== cleanerBrowseBtn) cleanerFileInput.click();
        });
        ['dragenter', 'dragover'].forEach(evt => {
            cleanerDropZone.addEventListener(evt, (e) => {
                e.preventDefault();
                cleanerDropZone.classList.add('dragover');
            });
        });
        ['dragleave', 'drop'].forEach(evt => {
            cleanerDropZone.addEventListener(evt, (e) => {
                e.preventDefault();
                cleanerDropZone.classList.remove('dragover');
            });
        });
        cleanerDropZone.addEventListener('drop', (e) => {
            const files = Array.from(e.dataTransfer.files).filter(f => f.name.endsWith('.xlsx') || f.name.endsWith('.xls'));
            if (files.length > 0) addCleanerFiles(files);
        });
    }

    if (cleanerFileInput) {
        cleanerFileInput.addEventListener('change', () => {
            const files = Array.from(cleanerFileInput.files);
            if (files.length > 0) addCleanerFiles(files);
            cleanerFileInput.value = '';
        });
    }

    function addCleanerFiles(files) {
        files.forEach(f => {
            if (!cleanerFiles.some(existing => existing.name === f.name && existing.size === f.size)) {
                cleanerFiles.push(f);
            }
        });
        renderCleanerFiles();
        if (cleanerFiles.length > 0) {
            analyzeCleanerFiles();
        }
    }

    function renderCleanerFiles() {
        if (!cleanerFilesGrid) return;
        cleanerFilesGrid.innerHTML = '';

        if (cleanerFiles.length > 0) {
            if (cleanerFilesSummary) cleanerFilesSummary.style.display = 'flex';
            if (cleanerFileCount) cleanerFileCount.textContent = cleanerFiles.length;
            if (cleanerExecuteBtn) cleanerExecuteBtn.disabled = false;

            cleanerFiles.forEach((file, index) => {
                const chip = document.createElement('div');
                chip.className = 'file-item-chip';
                chip.innerHTML = `
                    <span class="file-item-name" title="${file.name}">🧹 ${file.name}</span>
                    <span class="file-item-remove" data-cleaner-idx="${index}">&times;</span>
                `;
                cleanerFilesGrid.appendChild(chip);
            });

            cleanerFilesGrid.querySelectorAll('.file-item-remove').forEach(btn => {
                btn.addEventListener('click', (e) => {
                    e.stopPropagation();
                    const idx = parseInt(btn.getAttribute('data-cleaner-idx'));
                    cleanerFiles.splice(idx, 1);
                    renderCleanerFiles();
                    if (cleanerFiles.length > 0) {
                        analyzeCleanerFiles();
                    } else {
                        resetCleanerWorkspace();
                    }
                });
            });
        } else {
            if (cleanerFilesSummary) cleanerFilesSummary.style.display = 'none';
            if (cleanerExecuteBtn) cleanerExecuteBtn.disabled = true;
        }
    }

    async function analyzeCleanerFiles() {
        if (cleanerFiles.length === 0) return;
        const formData = new FormData();
        cleanerFiles.forEach(f => formData.append('files', f));

        try {
            const resp = await fetch('/api/analyze', { method: 'POST', body: formData });
            const data = await resp.json();
            cleanerAnalysis = data;

            applyDefaultAiRules();
            if (cleanerWorkspace) cleanerWorkspace.style.display = 'block';
            renderCleanerCardsWall();
            
            // Respect cleanerMode: collapsed by default if 'auto'
            if (cleanerCardsWall) {
                cleanerCardsWall.style.display = cleanerMode === 'custom' ? 'grid' : 'none';
            }
        } catch (err) {
            alert('读取解析表格失败: ' + err.message);
        }
    }

    function applyDefaultAiRules() {
        if (!cleanerAnalysis || !cleanerAnalysis.suggested_clusters) return;
        cleanerRules = {};

        cleanerAnalysis.suggested_clusters.forEach(c => {
            const colName = c.target;
            const colLower = colName.toLowerCase();
            
            let mask = 'none';
            let dateFormat = 'none';

            if (colLower.includes('手机') || colLower.includes('电话') || colLower.includes('phone') || colLower.includes('mobile')) {
                mask = 'phone_3_4';
            } else if (colLower.includes('邮箱') || colLower.includes('email')) {
                mask = 'email';
            } else if (colLower.includes('身份证') || colLower.includes('idcard')) {
                mask = 'idcard';
            } else if (colLower.includes('姓名') || colLower.includes('名字') || colLower.includes('name')) {
                mask = 'name';
            } else if (colLower.includes('日期') || colLower.includes('时间') || colLower.includes('生日') || colLower.includes('date') || colLower.includes('time')) {
                dateFormat = 'YYYY-MM-DD';
            }

            cleanerRules[colName] = {
                mask_type: mask,
                date_format: dateFormat,
                trim_space: true,
                remove_linebreaks: false,
                case_transform: 'none',
                fill_empty: ''
            };
        });
    }

    if (cleanerAiRecommendBtn) {
        cleanerAiRecommendBtn.addEventListener('click', () => {
            applyDefaultAiRules();
            renderCleanerCardsWall();
        });
    }

    function renderCleanerCardsWall() {
        if (!cleanerCardsWall || !cleanerAnalysis || !cleanerAnalysis.suggested_clusters) return;
        cleanerCardsWall.innerHTML = '';
        const clusters = cleanerAnalysis.suggested_clusters;
        if (cleanerColsCountBadge) cleanerColsCountBadge.textContent = `共 ${clusters.length} 个字段`;

        clusters.forEach((c, idx) => {
            const colName = c.target;
            const rule = cleanerRules[colName] || {};
            const colLower = colName.toLowerCase();

            // Extract up to 2 real sample values
            let samplesArr = [];
            if (c.quality_info && c.quality_info.samples_by_file) {
                c.quality_info.samples_by_file.forEach(sf => {
                    if (sf.samples) {
                        sf.samples.forEach(s => {
                            if (samplesArr.length < 2 && s !== null && s !== undefined && String(s).trim() !== '') {
                                const strVal = String(s).trim();
                                if (!samplesArr.includes(strVal)) {
                                    samplesArr.push(strVal);
                                }
                            }
                        });
                    }
                });
            }

            if (samplesArr.length === 0) {
                if (colLower.includes('手机') || colLower.includes('电话')) samplesArr = ['13812345678', '13988886666'];
                else if (colLower.includes('日期') || colLower.includes('时间') || colLower.includes('生日')) samplesArr = ['2026/09/08 14:30:00', '2026-9-1'];
                else if (colLower.includes('邮箱')) samplesArr = ['user@domain.com', 'service@qq.com'];
                else if (colLower.includes('身份证')) samplesArr = ['110101199003072345', '310104198512128888'];
                else if (colLower.includes('姓名') || colLower.includes('名字')) samplesArr = ['张三', '李四'];
                else samplesArr = [' 示例数据字符串 ', '样本数据2 '];
            }

            const isAiDetected = colLower.includes('手机') || colLower.includes('电话') || colLower.includes('日期') || colLower.includes('时间') || colLower.includes('邮箱') || colLower.includes('身份证');

            const card = document.createElement('div');
            card.className = 'cleaner-field-card';

            card.innerHTML = `
                <div class="cleaner-field-header">
                    <div class="cleaner-field-title-group">
                        <span class="cleaner-field-index">字段 #${idx + 1}</span>
                        <span class="cleaner-field-title">🏷️ ${colName}</span>
                    </div>
                    ${isAiDetected ? '<span class="cleaner-ai-tag">🤖 AI 推荐</span>' : ''}
                </div>

                <div class="cleaner-preview-box">
                    <div class="cleaner-preview-label">
                        <span>⚡ 卡片内清洗对比 (RAW ➔ CLEANED)</span>
                    </div>
                    <div class="cleaner-preview-rows" id="card-preview-rows-${idx}"></div>
                </div>

                <div class="cleaner-config-section">
                    <div class="cleaner-control-item">
                        <label>🔒 脱敏规则 (Data Masking)</label>
                        <select class="input-control card-mask-select" data-col="${colName}">
                            <option value="none" ${rule.mask_type === 'none' ? 'selected' : ''}>不脱敏 (保留明文)</option>
                            <option value="phone_3_4" ${rule.mask_type === 'phone_3_4' ? 'selected' : ''}>📱 手机号脱敏 (前3后4，中间****)</option>
                            <option value="email" ${rule.mask_type === 'email' ? 'selected' : ''}>📧 邮箱掩码 (例: u***r@domain.com)</option>
                            <option value="idcard" ${rule.mask_type === 'idcard' ? 'selected' : ''}>🪪 身份证隐藏生日 (110101******1234)</option>
                            <option value="name" ${rule.mask_type === 'name' ? 'selected' : ''}>👤 姓名脱敏 (例: 张*三)</option>
                            <option value="full_star" ${rule.mask_type === 'full_star' ? 'selected' : ''}>✳️ 全部星号 (****)</option>
                        </select>
                    </div>

                    <div class="cleaner-control-item">
                        <label>📅 日期统一格式 (Date Standardizer)</label>
                        <select class="input-control card-date-select" data-col="${colName}">
                            <option value="none" ${rule.date_format === 'none' ? 'selected' : ''}>保持原始格式</option>
                            <option value="YYYY-MM-DD" ${rule.date_format === 'YYYY-MM-DD' ? 'selected' : ''}>2026-09-08 (标准 YYYY-MM-DD)</option>
                            <option value="YYYY/MM/DD" ${rule.date_format === 'YYYY/MM/DD' ? 'selected' : ''}>2026/09/08 (斜杠 YYYY/MM/DD)</option>
                            <option value="YYYY年MM月DD日" ${rule.date_format === 'YYYY年MM月DD日' ? 'selected' : ''}>2026年09月08日 (中文年月)</option>
                            <option value="YYYYMMDD" ${rule.date_format === 'YYYYMMDD' ? 'selected' : ''}>20260908 (紧凑纯数字)</option>
                        </select>
                    </div>

                    <div class="cleaner-checkboxes-row">
                        <label class="checkbox-label-inline">
                            <input type="checkbox" class="card-trim-cb" data-col="${colName}" ${rule.trim_space ? 'checked' : ''}>
                            <span>✂️ 剔除首尾空格 (Trim)</span>
                        </label>
                        <label class="checkbox-label-inline">
                            <input type="checkbox" class="card-linebreaks-cb" data-col="${colName}" ${rule.remove_linebreaks ? 'checked' : ''}>
                            <span>🚫 清除文本强行换行</span>
                        </label>
                    </div>
                </div>
            `;

            cleanerCardsWall.appendChild(card);

            // Calculate initial card live preview
            updateCardLivePreview(card, colName, samplesArr, idx);

            // Bind change handlers for card controls
            const maskSelect = card.querySelector('.card-mask-select');
            const dateSelect = card.querySelector('.card-date-select');
            const trimCb = card.querySelector('.card-trim-cb');
            const linebreaksCb = card.querySelector('.card-linebreaks-cb');

            const syncCardRule = () => {
                cleanerRules[colName] = {
                    mask_type: maskSelect.value,
                    date_format: dateSelect.value,
                    trim_space: trimCb.checked,
                    remove_linebreaks: linebreaksCb.checked,
                    case_transform: 'none',
                    fill_empty: ''
                };
                updateCardLivePreview(card, colName, samplesArr, idx);
            };

            [maskSelect, dateSelect, trimCb, linebreaksCb].forEach(el => {
                if (el) el.addEventListener('change', syncCardRule);
            });
        });
    }

    function updateCardLivePreview(cardEl, colName, samplesArr, idx) {
        const previewRowsContainer = cardEl.querySelector(`#card-preview-rows-${idx}`);
        if (!previewRowsContainer) return;
        previewRowsContainer.innerHTML = '';

        const rule = cleanerRules[colName] || {};

        samplesArr.forEach(rawSample => {
            let cleaned = rawSample;

            if (rule.date_format && rule.date_format !== 'none') cleaned = mockDateTransform(cleaned, rule.date_format);
            if (rule.mask_type && rule.mask_type !== 'none') cleaned = mockMaskTransform(cleaned, rule.mask_type);
            if (rule.trim_space && typeof cleaned === 'string') cleaned = cleaned.trim();
            if (rule.remove_linebreaks && typeof cleaned === 'string') cleaned = cleaned.replace(/[\r\n]+/g, ' ');

            const row = document.createElement('div');
            row.className = 'cleaner-preview-flow';
            row.style.marginBottom = '0.3rem';
            row.innerHTML = `
                <span class="cleaner-val-before" title="${rawSample}">${rawSample}</span>
                <span class="cleaner-val-arrow">➔</span>
                <span class="cleaner-val-after" title="${cleaned}">${cleaned}</span>
            `;
            previewRowsContainer.appendChild(row);
        });
    }

    function resetCleanerWorkspace() {
        cleanerFiles = [];
        cleanerAnalysis = null;
        cleanerRules = {};
        cleanerDownloadId = null;
        renderCleanerFiles();
        if (cleanerCardsWall) cleanerCardsWall.innerHTML = '';
        if (cleanerWorkspace) cleanerWorkspace.style.display = 'none';
        if (cleanerProgressWrapper) cleanerProgressWrapper.style.display = 'none';
        if (cleanerDownloadBtn) cleanerDownloadBtn.style.display = 'none';
        if (cleanerExecuteBtn) {
            cleanerExecuteBtn.style.display = 'inline-block';
            cleanerExecuteBtn.disabled = true;
        }
    }

    if (cleanerResetBtn) cleanerResetBtn.addEventListener('click', resetCleanerWorkspace);
    if (cleanerClearFilesBtn) cleanerClearFilesBtn.addEventListener('click', resetCleanerWorkspace);

    if (cleanerDownloadBtn) {
        cleanerDownloadBtn.addEventListener('click', () => {
            if (cleanerDownloadId) {
                window.location.href = `/api/download/${cleanerDownloadId}`;
            } else {
                alert('暂无清洗完成的文件可供下载！');
            }
        });
    }

    if (cleanerExecuteBtn) {
        cleanerExecuteBtn.addEventListener('click', async () => {
            if (cleanerFiles.length === 0) {
                alert('请先上传 Excel 表格文件！');
                return;
            }

            if (cleanerDownloadBtn) cleanerDownloadBtn.style.display = 'none';
            if (cleanerProgressWrapper) cleanerProgressWrapper.style.display = 'block';
            if (cleanerProgressFill) cleanerProgressFill.style.width = '15%';
            if (cleanerProgressPercent) cleanerProgressPercent.textContent = '15%';
            if (cleanerProgressText) cleanerProgressText.textContent = `⚡ 正在批量规范清洗与转换 ${cleanerFiles.length} 个表格文件...`;

            let p = 15;
            const timer = setInterval(() => {
                if (p < 85) {
                    p += 10;
                    if (cleanerProgressFill) cleanerProgressFill.style.width = `${p}%`;
                    if (cleanerProgressPercent) cleanerProgressPercent.textContent = `${p}%`;
                }
            }, 250);

            const formData = new FormData();
            cleanerFiles.forEach(f => formData.append('files', f));
            formData.append('clean_rules_json', JSON.stringify(cleanerRules));

            try {
                const resp = await fetch('/api/clean', { method: 'POST', body: formData });
                const data = await resp.json();
                clearInterval(timer);

                if (data.error) {
                    if (cleanerProgressText) cleanerProgressText.textContent = '❌ 清洗失败: ' + data.error;
                    alert('格式化清洗失败: ' + data.error);
                    return;
                }

                if (cleanerProgressFill) cleanerProgressFill.style.width = '100%';
                if (cleanerProgressPercent) cleanerProgressPercent.textContent = '100%';
                if (cleanerProgressText) cleanerProgressText.textContent = `✅ 数据格式化与规范清洗成功！共完成 ${data.count || 1} 个表格清洗，请点击右下角按钮下载。`;

                if (data.download_id) {
                    cleanerDownloadId = data.download_id;
                    if (cleanerDownloadBtn) {
                        cleanerDownloadBtn.style.display = 'inline-block';
                        cleanerDownloadBtn.textContent = `📥 点击下载清洗后的 Excel (${data.filename || '点击下载'}) ➔`;
                    }
                    if (cleanerExecuteBtn) cleanerExecuteBtn.style.display = 'none';
                }
            } catch (err) {
                clearInterval(timer);
                if (cleanerProgressText) cleanerProgressText.textContent = '❌ 清洗出错: ' + err.message;
                alert('格式化清洗失败: ' + err.message);
            }
        });
    }

    // ==========================================
    // Dedicated PDF/Word Bidirectional Converter (#page-converter)
    // ==========================================
    const tabPdfToWord = document.getElementById('tab-mode-pdf-to-word');
    const tabWordToPdf = document.getElementById('tab-mode-word-to-pdf');
    const converterDropZone = document.getElementById('converter-drop-zone');
    const converterFileInput = document.getElementById('converter-file-input');
    const converterBrowseBtn = document.getElementById('converter-browse-btn');
    const converterDropIcon = document.getElementById('converter-drop-icon');
    const converterDropTitle = document.getElementById('converter-drop-title');
    const converterDropDesc = document.getElementById('converter-drop-desc');
    const converterFilesSummary = document.getElementById('converter-files-summary');
    const converterFileCount = document.getElementById('converter-file-count');
    const converterFilesGrid = document.getElementById('converter-files-grid');
    const converterClearFilesBtn = document.getElementById('converter-clear-files-btn');
    const converterProgressWrapper = document.getElementById('converter-progress-wrapper');
    const converterProgressText = document.getElementById('converter-progress-text');
    const converterProgressPercent = document.getElementById('converter-progress-percent');
    const converterProgressFill = document.getElementById('converter-progress-fill');
    const converterResetBtn = document.getElementById('converter-reset-btn');
    const converterExecuteBtn = document.getElementById('converter-execute-btn');

    let converterMode = 'pdf_to_word';
    let converterFiles = [];

    if (tabPdfToWord && tabWordToPdf) {
        tabPdfToWord.addEventListener('click', () => setConverterMode('pdf_to_word'));
        tabWordToPdf.addEventListener('click', () => setConverterMode('word_to_pdf'));
    }

    function setConverterMode(mode) {
        converterMode = mode;
        converterFiles = [];
        renderConverterFiles();

        if (mode === 'pdf_to_word') {
            tabPdfToWord.classList.add('active');
            tabWordToPdf.classList.remove('active');
            converterDropIcon.textContent = '📄';
            converterDropTitle.innerHTML = '拖拽 PDF 文件到此处，或 <span class="browse-btn" id="converter-browse-btn">点击浏览选择文件</span>';
            converterDropDesc.textContent = '支持选择单个或多个 PDF (.pdf) 文件转为可编辑 Word';
            converterFileInput.setAttribute('accept', '.pdf');
        } else {
            tabWordToPdf.classList.add('active');
            tabPdfToWord.classList.remove('active');
            converterDropIcon.textContent = '📝';
            converterDropTitle.innerHTML = '拖拽 Word 文档到此处，或 <span class="browse-btn" id="converter-browse-btn">点击浏览选择文件</span>';
            converterDropDesc.textContent = '支持选择单个或多个 Word (.docx, .doc) 文档打包导出为 PDF';
            converterFileInput.setAttribute('accept', '.docx, .doc');
        }

        const newBrowseBtn = document.getElementById('converter-browse-btn');
        if (newBrowseBtn) newBrowseBtn.addEventListener('click', () => converterFileInput.click());
    }

    if (converterBrowseBtn) converterBrowseBtn.addEventListener('click', () => converterFileInput.click());
    if (converterDropZone) {
        converterDropZone.addEventListener('click', (e) => {
            if (e.target.id !== 'converter-browse-btn') converterFileInput.click();
        });
        ['dragenter', 'dragover'].forEach(evt => {
            converterDropZone.addEventListener(evt, (e) => {
                e.preventDefault();
                converterDropZone.classList.add('dragover');
            });
        });
        ['dragleave', 'drop'].forEach(evt => {
            converterDropZone.addEventListener(evt, (e) => {
                e.preventDefault();
                converterDropZone.classList.remove('dragover');
            });
        });
        converterDropZone.addEventListener('drop', (e) => {
            let files = Array.from(e.dataTransfer.files);
            if (converterMode === 'pdf_to_word') {
                files = files.filter(f => f.name.toLowerCase().endsWith('.pdf'));
            } else {
                files = files.filter(f => f.name.toLowerCase().endsWith('.docx') || f.name.toLowerCase().endsWith('.doc'));
            }
            if (files.length > 0) addConverterFiles(files);
        });
    }

    if (converterFileInput) {
        converterFileInput.addEventListener('change', () => {
            const files = Array.from(converterFileInput.files);
            if (files.length > 0) addConverterFiles(files);
            converterFileInput.value = '';
        });
    }

    function addConverterFiles(files) {
        files.forEach(f => {
            if (!converterFiles.some(existing => existing.name === f.name && existing.size === f.size)) {
                converterFiles.push(f);
            }
        });
        renderConverterFiles();
    }

    function renderConverterFiles() {
        if (!converterFilesGrid) return;
        converterFilesGrid.innerHTML = '';

        if (converterFiles.length > 0) {
            converterFilesSummary.style.display = 'flex';
            converterFileCount.textContent = converterFiles.length;
            converterExecuteBtn.disabled = false;

            converterFiles.forEach((file, index) => {
                const chip = document.createElement('div');
                chip.className = 'file-item-chip';
                const icon = converterMode === 'pdf_to_word' ? '📄' : '📝';
                chip.innerHTML = `
                    <span class="file-item-name" title="${file.name}">${icon} ${file.name}</span>
                    <span class="file-item-remove" data-converter-idx="${index}">&times;</span>
                `;
                converterFilesGrid.appendChild(chip);
            });

            converterFilesGrid.querySelectorAll('.file-item-remove').forEach(btn => {
                btn.addEventListener('click', (e) => {
                    e.stopPropagation();
                    const idx = parseInt(btn.getAttribute('data-converter-idx'));
                    converterFiles.splice(idx, 1);
                    renderConverterFiles();
                });
            });
        } else {
            converterFilesSummary.style.display = 'none';
            converterExecuteBtn.disabled = true;
        }
    }

    if (converterClearFilesBtn) {
        converterClearFilesBtn.addEventListener('click', () => {
            converterFiles = [];
            renderConverterFiles();
        });
    }

    const converterDownloadBtn = document.getElementById('converter-download-btn');
    let converterDownloadId = null;

    if (converterDownloadBtn) {
        converterDownloadBtn.addEventListener('click', () => {
            if (converterDownloadId) {
                window.location.href = `/api/download/${converterDownloadId}`;
            } else {
                alert('暂无转换完成的文件可供下载！');
            }
        });
    }

    if (converterResetBtn) {
        converterResetBtn.addEventListener('click', () => {
            converterFiles = [];
            converterDownloadId = null;
            renderConverterFiles();
            if (converterProgressWrapper) converterProgressWrapper.style.display = 'none';
            if (converterDownloadBtn) converterDownloadBtn.style.display = 'none';
            if (converterExecuteBtn) converterExecuteBtn.style.display = 'inline-block';
        });
    }

    if (converterExecuteBtn) {
        converterExecuteBtn.addEventListener('click', async () => {
            if (converterFiles.length === 0) {
                alert('请先选择需要转换的文件！');
                return;
            }

            if (converterDownloadBtn) converterDownloadBtn.style.display = 'none';
            converterProgressWrapper.style.display = 'block';
            converterProgressFill.style.width = '15%';
            converterProgressPercent.textContent = '15%';
            const modeText = converterMode === 'pdf_to_word' ? 'PDF ➔ Word' : 'Word ➔ PDF';
            converterProgressText.textContent = `⏳ 正在读取 ${converterFiles.length} 个文件，启动 ${modeText} 转换引擎...`;

            let p = 15;
            const timer = setInterval(() => {
                if (p < 85) {
                    p += 8;
                    converterProgressFill.style.width = `${p}%`;
                    converterProgressPercent.textContent = `${p}%`;
                }
            }, 300);

            const formData = new FormData();
            converterFiles.forEach(f => formData.append('files', f));
            formData.append('mode', converterMode);

            try {
                const resp = await fetch('/api/convert-pdf-word', {
                    method: 'POST',
                    body: formData
                });
                const data = await resp.json();
                clearInterval(timer);

                if (data.error) {
                    converterProgressText.textContent = '❌ 转换失败: ' + data.error;
                    alert('转换失败: ' + data.error);
                    return;
                }

                converterProgressFill.style.width = '100%';
                converterProgressPercent.textContent = '100%';
                converterProgressText.textContent = `✅ 转换成功！共完成 ${data.count || 1} 个文件转换，请点击右下角按钮下载。`;

                if (data.download_id) {
                    converterDownloadId = data.download_id;
                    if (converterDownloadBtn) {
                        converterDownloadBtn.style.display = 'inline-block';
                        converterDownloadBtn.textContent = `📥 点击下载转换后的文件 (${data.filename || '点击下载'}) ➔`;
                    }
                    if (converterExecuteBtn) converterExecuteBtn.style.display = 'none';
                }
            } catch (err) {
                clearInterval(timer);
                converterProgressText.textContent = '❌ 转换出错: ' + err.message;
                alert('转换出错: ' + err.message);
            }
        });
    }

    // ==========================================
    // Dedicated Image Compressor Tool Page (#page-compressor)
    // ==========================================
    const tabHighCompress = document.getElementById('tab-mode-high-compress');
    const tabHighQuality = document.getElementById('tab-mode-high-quality');
    const compressorDropZone = document.getElementById('compressor-drop-zone');
    const compressorFileInput = document.getElementById('compressor-file-input');
    const compressorBrowseBtn = document.getElementById('compressor-browse-btn');
    const compressorFilesSummary = document.getElementById('compressor-files-summary');
    const compressorFileCount = document.getElementById('compressor-file-count');
    const compressorFilesGrid = document.getElementById('compressor-files-grid');
    const compressorClearFilesBtn = document.getElementById('compressor-clear-files-btn');
    const compressorProgressWrapper = document.getElementById('compressor-progress-wrapper');
    const compressorProgressText = document.getElementById('compressor-progress-text');
    const compressorProgressPercent = document.getElementById('compressor-progress-percent');
    const compressorProgressFill = document.getElementById('compressor-progress-fill');
    const compressorStatsRow = document.getElementById('compressor-stats-row');
    const statOrigSize = document.getElementById('stat-orig-size');
    const statCompressedSize = document.getElementById('stat-compressed-size');
    const statSavedPercent = document.getElementById('stat-saved-percent');
    const compressorResetBtn = document.getElementById('compressor-reset-btn');
    const compressorExecuteBtn = document.getElementById('compressor-execute-btn');
    const compressorDownloadBtn = document.getElementById('compressor-download-btn');

    let compressorMode = 'high_compress';
    let compressorFiles = [];
    let compressorDownloadId = null;

    if (tabHighCompress && tabHighQuality) {
        tabHighCompress.addEventListener('click', () => setCompressorMode('high_compress'));
        tabHighQuality.addEventListener('click', () => setCompressorMode('high_quality'));
    }

    function setCompressorMode(mode) {
        compressorMode = mode;
        if (mode === 'high_compress') {
            tabHighCompress.classList.add('active');
            tabHighQuality.classList.remove('active');
        } else {
            tabHighQuality.classList.add('active');
            tabHighCompress.classList.remove('active');
        }
    }

    if (compressorBrowseBtn) compressorBrowseBtn.addEventListener('click', () => compressorFileInput.click());
    if (compressorDropZone) {
        compressorDropZone.addEventListener('click', (e) => {
            if (e.target.id !== 'compressor-browse-btn') compressorFileInput.click();
        });
        ['dragenter', 'dragover'].forEach(evt => {
            compressorDropZone.addEventListener(evt, (e) => {
                e.preventDefault();
                compressorDropZone.classList.add('dragover');
            });
        });
        ['dragleave', 'drop'].forEach(evt => {
            compressorDropZone.addEventListener(evt, (e) => {
                e.preventDefault();
                compressorDropZone.classList.remove('dragover');
            });
        });
        compressorDropZone.addEventListener('drop', (e) => {
            const validExts = ['.jpg', '.jpeg', '.png', '.webp', '.bmp'];
            const files = Array.from(e.dataTransfer.files).filter(f => validExts.some(ext => f.name.toLowerCase().endsWith(ext)));
            if (files.length > 0) addCompressorFiles(files);
        });
    }

    if (compressorFileInput) {
        compressorFileInput.addEventListener('change', () => {
            const files = Array.from(compressorFileInput.files);
            if (files.length > 0) addCompressorFiles(files);
            compressorFileInput.value = '';
        });
    }

    function addCompressorFiles(files) {
        files.forEach(f => {
            if (!compressorFiles.some(existing => existing.name === f.name && existing.size === f.size)) {
                compressorFiles.push(f);
            }
        });
        renderCompressorFiles();
    }

    function renderCompressorFiles() {
        if (!compressorFilesGrid) return;
        compressorFilesGrid.innerHTML = '';

        if (compressorFiles.length > 0) {
            compressorFilesSummary.style.display = 'flex';
            compressorFileCount.textContent = compressorFiles.length;
            compressorExecuteBtn.disabled = false;

            compressorFiles.forEach((file, index) => {
                const chip = document.createElement('div');
                chip.className = 'file-item-chip';
                chip.innerHTML = `
                    <span class="file-item-name" title="${file.name}">🖼️ ${file.name}</span>
                    <span class="file-item-remove" data-compressor-idx="${index}">&times;</span>
                `;
                compressorFilesGrid.appendChild(chip);
            });

            compressorFilesGrid.querySelectorAll('.file-item-remove').forEach(btn => {
                btn.addEventListener('click', (e) => {
                    e.stopPropagation();
                    const idx = parseInt(btn.getAttribute('data-compressor-idx'));
                    compressorFiles.splice(idx, 1);
                    renderCompressorFiles();
                });
            });
        } else {
            compressorFilesSummary.style.display = 'none';
            compressorExecuteBtn.disabled = true;
        }
    }

    if (compressorClearFilesBtn) {
        compressorClearFilesBtn.addEventListener('click', () => {
            compressorFiles = [];
            renderCompressorFiles();
        });
    }

    if (compressorResetBtn) {
        compressorResetBtn.addEventListener('click', () => {
            compressorFiles = [];
            compressorDownloadId = null;
            renderCompressorFiles();
            if (compressorProgressWrapper) compressorProgressWrapper.style.display = 'none';
            if (compressorStatsRow) compressorStatsRow.style.display = 'none';
            if (compressorDownloadBtn) compressorDownloadBtn.style.display = 'none';
            if (compressorExecuteBtn) compressorExecuteBtn.style.display = 'inline-block';
        });
    }

    if (compressorDownloadBtn) {
        compressorDownloadBtn.addEventListener('click', () => {
            if (compressorDownloadId) {
                window.location.href = `/api/download/${compressorDownloadId}`;
            } else {
                alert('暂无压缩完成的图片可供下载！');
            }
        });
    }

    if (compressorExecuteBtn) {
        compressorExecuteBtn.addEventListener('click', async () => {
            if (compressorFiles.length === 0) {
                alert('请先选择需要压缩的图片文件！');
                return;
            }

            if (compressorDownloadBtn) compressorDownloadBtn.style.display = 'none';
            if (compressorStatsRow) compressorStatsRow.style.display = 'none';
            compressorProgressWrapper.style.display = 'block';
            compressorProgressFill.style.width = '15%';
            compressorProgressPercent.textContent = '15%';
            const modeText = compressorMode === 'high_compress' ? '高压缩模式 (几十~几百KB)' : '高质量模式 (~1MB优化)';
            compressorProgressText.textContent = `⚡ 正在并行压缩 ${compressorFiles.length} 张图片 (${modeText})...`;

            let p = 15;
            const timer = setInterval(() => {
                if (p < 85) {
                    p += 10;
                    compressorProgressFill.style.width = `${p}%`;
                    compressorProgressPercent.textContent = `${p}%`;
                }
            }, 250);

            const formData = new FormData();
            compressorFiles.forEach(f => formData.append('files', f));
            formData.append('mode', compressorMode);

            try {
                const resp = await fetch('/api/compress-images', {
                    method: 'POST',
                    body: formData
                });
                const data = await resp.json();
                clearInterval(timer);

                if (data.error) {
                    compressorProgressText.textContent = '❌ 压缩失败: ' + data.error;
                    alert('压缩失败: ' + data.error);
                    return;
                }

                compressorProgressFill.style.width = '100%';
                compressorProgressPercent.textContent = '100%';
                compressorProgressText.textContent = `✅ 图片压缩成功！共完成 ${data.count || 1} 张图片压缩。`;

                // Update Stats
                if (statOrigSize) statOrigSize.textContent = data.orig_size_str || '0 MB';
                if (statCompressedSize) statCompressedSize.textContent = data.compressed_size_str || '0 KB';
                if (statSavedPercent) statSavedPercent.textContent = data.saved_percent_str || '0%';
                if (compressorStatsRow) compressorStatsRow.style.display = 'flex';

                if (data.download_id) {
                    compressorDownloadId = data.download_id;
                    if (compressorDownloadBtn) {
                        compressorDownloadBtn.style.display = 'inline-block';
                        compressorDownloadBtn.textContent = `📥 点击下载压缩图片 (${data.filename || '点击下载'}) ➔`;
                    }
                    if (compressorExecuteBtn) compressorExecuteBtn.style.display = 'none';
                }
            } catch (err) {
                clearInterval(timer);
                compressorProgressText.textContent = '❌ 压缩出错: ' + err.message;
                alert('压缩出错: ' + err.message);
            }
        });
    }
});
