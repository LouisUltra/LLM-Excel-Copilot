/**
 * Excel 智能助手 - 前端逻辑 (Dark Tech UI Refactor)
 */

// ============ 全局错误捕获 ============
window.addEventListener('error', (event) => {
    console.error('💥 全局错误:', event.error);
});

// ============ 状态管理 ============
const state = {
    fileId: null,
    sessionId: null,
    metadata: null,
    downloadUrl: null,
    isProcessing: false,
    currentAnswers: {},
    currentRequestId: 0,
    // 多文件支持
    files: [],  // [{fileId, metadata, originalName}...]
    isMultiFileMode: false,
    currentFileIndex: 0,  // 当前显示的文件索引
    // 🔒 防止死循环的状态变量
    lastSubmitTime: 0,  // 上次提交时间戳
    sameInputCount: 0,  // 相同输入计数
    // 重试支持
    lastUserInput: null,  // 保存最后一次用户输入用于重试
    lastUserMessageEl: null,  // 最后一条用户消息DOM元素，用于重试时替换
    lastAssistantMessageEl: null,  // 最后一条助手消息DOM元素，用于重试时移除
    // 操作上下文
    lastOperationPlan: null  // 上一次成功执行的操作计划，用于继续编辑时的上下文
};

// ============ DOM 元素 ============
const elements = {
    uploadSection: document.getElementById('upload-section'),
    uploadArea: document.getElementById('upload-area'),
    fileInput: document.getElementById('file-input'),
    uploadProgress: document.getElementById('upload-progress'),
    progressFill: document.getElementById('progress-fill'),
    progressText: document.getElementById('progress-text'),
    // 多文件 UI
    uploadedFilesList: document.getElementById('uploaded-files-list'),
    multiFileActions: document.getElementById('multi-file-actions'),
    btnAddMoreFiles: document.getElementById('btn-add-more-files'),
    btnStartProcessing: document.getElementById('btn-start-processing'),

    fileInfoSection: document.getElementById('file-info-section'),
    fileName: document.getElementById('file-name'),
    fileSummary: document.getElementById('file-summary'),
    fileHeaders: document.getElementById('file-headers'),
    fileCardsContainer: document.getElementById('file-cards-container'),
    fileDotsNav: document.getElementById('file-dots-nav'),
    btnRemoveFile: document.getElementById('btn-remove-file'),

    workspaceArea: document.getElementById('workspace-area'), // New wrapper
    chatSection: document.getElementById('chat-section'),
    chatContainer: document.getElementById('chat-container'),
    userInput: document.getElementById('user-input'),
    btnSend: document.getElementById('btn-send'),

    resultSection: document.getElementById('result-section'),
    statusCard: document.getElementById('status-card'),
    statusTitle: document.getElementById('status-title'),
    statusDesc: document.getElementById('status-desc'),
    statusIcon: document.getElementById('status-icon'),
    statusIconBg: document.getElementById('status-icon-bg'),

    actionCard: document.getElementById('action-card'),
    btnDownload: document.getElementById('btn-download'),
    btnContinue: document.getElementById('btn-continue'),
    btnNewTask: document.getElementById('btn-new-task'),

    loadingOverlay: document.getElementById('loading-overlay'),
    loadingText: document.getElementById('loading-text')
};

// ============ API 调用 (保持不变) ============
const api = {
    async upload(file) {
        const formData = new FormData();
        formData.append('file', file);
        const response = await fetch('/api/upload', { method: 'POST', body: formData });
        if (!response.ok) throw new Error((await response.json()).detail || '上传失败');
        return response.json();
    },

    async refine(fileId, userInput, sessionId = null, answers = null, fileIds = [], previousOperations = null) {
        const body = { file_id: fileId, user_input: userInput };
        if (sessionId) body.session_id = sessionId;
        if (answers) body.answers = answers;
        if (fileIds.length > 0) body.file_ids = fileIds;
        if (previousOperations) body.previous_operations = previousOperations;
        const response = await fetch('/api/refine', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(body)
        });
        if (!response.ok) throw new Error((await response.json()).detail || '请求失败');
        return response.json();
    },

    async process(fileId, sessionId) {
        const response = await fetch('/api/process', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ file_id: fileId, session_id: sessionId, confirmed: true })
        });
        if (!response.ok) throw new Error((await response.json()).detail || '处理失败');
        return response.json();
    },

    async continueProcessing(fileId) {
        const response = await fetch(`/api/continue/${fileId}`, { method: 'POST' });
        if (!response.ok) throw new Error((await response.json()).detail || '继续处理失败');
        return response.json();
    }
};

// ============ 工具函数 ============

// 简单 Markdown 解析器
function parseSimpleMarkdown(text) {
    if (!text) return '';

    let result = text
        // 转义 HTML
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        // 粗体 **text** 或 __text__
        .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
        .replace(/__(.+?)__/g, '<strong>$1</strong>')
        // 斜体 *text* 或 _text_
        .replace(/\*(.+?)\*/g, '<em>$1</em>')
        .replace(/_(.+?)_/g, '<em>$1</em>')
        // 行内代码 `code`
        .replace(/`(.+?)`/g, '<code class="px-1.5 py-0.5 bg-slate-700/50 rounded text-xs font-mono text-blue-300">$1</code>');

    // 处理列表项：先将连续的列表项包装在 ul 中
    const lines = result.split('\n');
    let inList = false;
    let processedLines = [];

    for (let line of lines) {
        if (line.startsWith('- ')) {
            if (!inList) {
                processedLines.push('<ul class="my-1 space-y-0.5">');
                inList = true;
            }
            processedLines.push(`<li class="ml-4 list-disc">${line.substring(2)}</li>`);
        } else {
            if (inList) {
                processedLines.push('</ul>');
                inList = false;
            }
            processedLines.push(line);
        }
    }
    if (inList) {
        processedLines.push('</ul>');
    }

    // 非列表项之间用 <br> 分隔
    return processedLines.join('').replace(/<\/ul>/g, '</ul><br>').replace(/<br><ul/g, '<ul');
}

// ============ UI 更新函数 ============

function showSection(sectionName) {
    if (sectionName === 'upload') {
        elements.uploadSection.classList.remove('hidden');
        elements.fileInfoSection.classList.add('hidden');
        elements.workspaceArea.classList.add('hidden');
    } else if (sectionName === 'chat') {
        elements.uploadSection.classList.add('hidden');
        elements.fileInfoSection.classList.remove('hidden');
        elements.workspaceArea.classList.remove('hidden');
        // Reset Status Card
        elements.statusCard.classList.add('hidden');
        elements.actionCard.classList.add('hidden');
    } else if (sectionName === 'result') {
        elements.statusCard.classList.remove('hidden');
        elements.actionCard.classList.remove('hidden');
    }
}

function showLoading(text = '处理中...') {
    elements.loadingText.textContent = text;
    elements.loadingOverlay.classList.remove('hidden');
}

function hideLoading() {
    elements.loadingOverlay.classList.add('hidden');
}

function updateFileInfo(metadata) {
    // 多文件模式：使用轮播显示
    if (state.files.length > 1) {
        renderMultiFileCarousel();
        return;
    }

    // 单文件模式：直接显示
    renderSingleFileCard(metadata);
}

// 渲染单个文件卡片内容
function renderSingleFileCard(metadata) {
    elements.fileName.textContent = metadata.file_name;
    const sheets = metadata.sheets;
    const totalRows = sheets.reduce((sum, s) => sum + s.total_rows, 0);
    const totalCols = sheets[0]?.total_cols || 0;
    elements.fileSummary.textContent = `${sheets.length} 个工作表 | ${totalRows} 行 | ${totalCols} 列`;

    // 显示表头信息
    if (elements.fileHeaders && sheets[0]?.columns) {
        elements.fileHeaders.innerHTML = '';
        const columns = sheets[0].columns;

        // 限制显示的列数，避免太多
        const displayCols = columns.slice(0, 15);
        displayCols.forEach(col => {
            const tag = document.createElement('span');
            tag.className = 'inline-flex items-center gap-1.5 px-2.5 py-1 rounded-md text-xs bg-slate-800/50 border border-slate-700/50 text-slate-300';

            // 根据数据类型添加图标
            let icon = '📝';
            if (col.data_type === '数字') icon = '🔢';
            else if (col.data_type === '日期') icon = '📅';
            else if (col.data_type === '布尔') icon = '✓';

            tag.innerHTML = `<span class="opacity-70">${icon}</span>${col.name}`;
            elements.fileHeaders.appendChild(tag);
        });

        // 如果还有更多列
        if (columns.length > 15) {
            const moreTag = document.createElement('span');
            moreTag.className = 'inline-flex items-center px-2.5 py-1 rounded-md text-xs bg-blue-500/10 border border-blue-500/20 text-blue-400';
            moreTag.textContent = `+${columns.length - 15} 列`;
            elements.fileHeaders.appendChild(moreTag);
        }
    }

    // 隐藏点导航
    if (elements.fileDotsNav) {
        elements.fileDotsNav.classList.add('hidden');
    }
}

// 渲染多文件轮播
function renderMultiFileCarousel() {
    const currentFile = state.files[state.currentFileIndex];
    if (!currentFile) return;

    // 更新当前显示的文件内容
    renderSingleFileCard(currentFile.metadata);

    // 显示点导航
    if (elements.fileDotsNav && state.files.length > 1) {
        elements.fileDotsNav.classList.remove('hidden');
        elements.fileDotsNav.innerHTML = '';

        state.files.forEach((file, index) => {
            const dot = document.createElement('button');
            dot.className = `w-2.5 h-2.5 rounded-full transition-all ${index === state.currentFileIndex
                ? 'bg-blue-500 scale-110'
                : 'bg-slate-600 hover:bg-slate-500'
                }`;
            dot.title = file.metadata.file_name;
            dot.onclick = () => switchToFile(index);
            elements.fileDotsNav.appendChild(dot);
        });

        // 添加文件名提示
        const hint = document.createElement('span');
        hint.className = 'ml-3 text-xs text-slate-500';
        hint.textContent = `${state.currentFileIndex + 1} / ${state.files.length}: ${currentFile.metadata.file_name}`;
        elements.fileDotsNav.appendChild(hint);
    }
}

// 切换显示的文件
function switchToFile(index) {
    if (index >= 0 && index < state.files.length) {
        state.currentFileIndex = index;
        renderMultiFileCarousel();
    }
}

function scrollToBottom() {
    elements.chatContainer.scrollTo({
        top: elements.chatContainer.scrollHeight,
        behavior: 'smooth'
    });
}

function addMessage(role, content) {
    const isUser = role === 'user';

    // Wrapper
    const wrapper = document.createElement('div');
    wrapper.className = `flex gap-3 ${isUser ? 'flex-row-reverse' : ''} animate-fade-in`;

    // Icon
    const iconDiv = document.createElement('div');
    iconDiv.className = `w-8 h-8 rounded-full flex-shrink-0 flex items-center justify-center border ${isUser ? 'bg-slate-700 border-slate-600' : 'bg-blue-600/20 border-blue-500/30'
        }`;
    iconDiv.innerHTML = `<i data-lucide="${isUser ? 'user' : 'bot'}" class="w-4 h-4 ${isUser ? 'text-slate-300' : 'text-blue-400'}"></i>`;

    // Bubble
    const bubble = document.createElement('div');
    bubble.className = `p-4 max-w-[85%] rounded-2xl text-sm leading-relaxed shadow-sm ${isUser
        ? 'bg-blue-600 text-white rounded-tr-none'
        : 'glass-panel rounded-tl-none text-slate-200 border-slate-700/50'
        }`;

    // Content - 支持简单 Markdown 渲染
    if (typeof content === 'string') {
        bubble.innerHTML = parseSimpleMarkdown(content);
    } else {
        bubble.appendChild(content);
    }

    wrapper.appendChild(iconDiv);
    wrapper.appendChild(bubble);
    elements.chatContainer.appendChild(wrapper);

    lucide.createIcons({ root: wrapper });
    scrollToBottom();

    // 追踪最后的消息元素（用于重试时移除）
    if (isUser) {
        state.lastUserMessageEl = wrapper;
    } else {
        state.lastAssistantMessageEl = wrapper;
    }

    return wrapper;
}

function addTypingIndicator() {
    const wrapper = document.createElement('div');
    wrapper.className = 'flex gap-3 animate-fade-in typing-message';

    const iconDiv = document.createElement('div');
    iconDiv.className = 'w-8 h-8 rounded-full bg-blue-600/20 flex-shrink-0 flex items-center justify-center border border-blue-500/30';
    iconDiv.innerHTML = '<i data-lucide="bot" class="w-4 h-4 text-blue-400"></i>';

    const bubble = document.createElement('div');
    bubble.className = 'glass-panel p-4 rounded-2xl rounded-tl-none border-slate-700/50 flex items-center gap-1.5';

    // Dots
    [1, 2, 3].forEach(i => {
        const dot = document.createElement('div');
        dot.className = 'w-1.5 h-1.5 bg-blue-400/50 rounded-full animate-pulse';
        dot.style.animationDelay = `${i * 0.15}s`;
        bubble.appendChild(dot);
    });

    wrapper.appendChild(iconDiv);
    wrapper.appendChild(bubble);
    elements.chatContainer.appendChild(wrapper);

    lucide.createIcons({ root: wrapper });
    scrollToBottom();
    return wrapper;
}

function removeTypingIndicator() {
    const typing = elements.chatContainer.querySelector('.typing-message');
    if (typing) typing.remove();
}

// 继续编辑分隔线
function addContinueSessionDivider() {
    const divider = document.createElement('div');
    divider.className = 'continue-session-divider flex items-center gap-4 my-6 animate-fade-in';
    divider.innerHTML = `
        <div class="flex-1 h-px bg-gradient-to-r from-transparent via-slate-600 to-transparent"></div>
        <div class="flex items-center gap-2 px-4 py-1.5 rounded-full bg-slate-800/50 border border-slate-700/50">
            <i data-lucide="refresh-cw" class="w-3.5 h-3.5 text-blue-400"></i>
            <span class="text-xs text-slate-400">继续编辑</span>
        </div>
        <div class="flex-1 h-px bg-gradient-to-r from-transparent via-slate-600 to-transparent"></div>
    `;
    elements.chatContainer.appendChild(divider);
    lucide.createIcons({ root: divider });
    scrollToBottom();
}

// 创建带重试按钮的错误消息
function createErrorWithRetry(errorMessage, onRetry) {
    const container = document.createElement('div');
    container.className = 'bg-red-500/10 border border-red-500/30 rounded-xl p-4 space-y-3';

    // 错误信息
    const errorText = document.createElement('div');
    errorText.className = 'flex items-start gap-2 text-sm text-red-400';
    errorText.innerHTML = `
        <i data-lucide="alert-circle" class="w-4 h-4 mt-0.5 flex-shrink-0"></i>
        <span>${errorMessage}</span>
    `;
    container.appendChild(errorText);

    // 操作按钮
    const actions = document.createElement('div');
    actions.className = 'flex gap-2 mt-2';

    // 重试按钮
    const retryBtn = document.createElement('button');
    retryBtn.className = 'flex items-center gap-1.5 px-3 py-1.5 text-xs bg-blue-600 hover:bg-blue-500 text-white rounded-lg transition-colors';
    retryBtn.innerHTML = '<i data-lucide="refresh-cw" class="w-3 h-3"></i> 重试';
    retryBtn.onclick = () => {
        container.remove();
        if (onRetry) onRetry();
    };
    actions.appendChild(retryBtn);

    // 设置按钮（切换API配置）
    const settingsBtn = document.createElement('button');
    settingsBtn.className = 'flex items-center gap-1.5 px-3 py-1.5 text-xs bg-slate-700 hover:bg-slate-600 text-slate-200 rounded-lg transition-colors';
    settingsBtn.innerHTML = '<i data-lucide="settings" class="w-3 h-3"></i> 切换 API';
    settingsBtn.onclick = () => {
        document.getElementById('settings-modal')?.classList.remove('hidden');
    };
    actions.appendChild(settingsBtn);

    container.appendChild(actions);

    setTimeout(() => lucide.createIcons({ root: container }), 0);
    return container;
}

// 交互式组件：问题块
function createQuestionBlock(questions, onConfirm, onModify) {
    const container = document.createElement('div');
    container.className = 'flex flex-col gap-3 mt-2';

    questions.forEach(q => {
        const block = document.createElement('div');
        block.className = 'bg-slate-800/50 rounded-lg p-3 border border-slate-700';

        const qText = document.createElement('p');
        qText.className = 'font-medium text-white mb-2';
        qText.textContent = q.question;
        block.appendChild(qText);

        const optionsDiv = document.createElement('div');
        optionsDiv.className = 'space-y-2';

        q.options.forEach(opt => {
            const label = document.createElement('label');
            label.className = 'flex items-center gap-3 p-2 rounded hover:bg-slate-700/50 cursor-pointer transition-colors';

            const input = document.createElement('input');
            input.type = q.question_type === 'multiple' ? 'checkbox' : 'radio';
            input.name = `q-${q.question_id}`;
            input.className = 'accent-blue-500 w-4 h-4';
            input.addEventListener('change', () => {
                state.currentAnswers[q.question_id] = opt.key;
            });

            const text = document.createElement('span');
            text.textContent = opt.label;
            text.className = 'text-slate-300 text-sm';

            label.appendChild(input);
            label.appendChild(text);
            optionsDiv.appendChild(label);
        });
        block.appendChild(optionsDiv);
        container.appendChild(block);
    });

    const btnRow = document.createElement('div');
    btnRow.className = 'flex gap-2 mt-2';

    const confirmBtn = document.createElement('button');
    confirmBtn.className = 'px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white rounded-lg text-xs font-medium transition-colors';
    confirmBtn.textContent = '确认选择';
    confirmBtn.onclick = async () => {
        confirmBtn.textContent = '提交中...';
        confirmBtn.disabled = true;
        await onConfirm();
    };

    const modifyBtn = document.createElement('button');
    modifyBtn.className = 'px-4 py-2 bg-slate-700 hover:bg-slate-600 text-slate-200 rounded-lg text-xs font-medium transition-colors';
    modifyBtn.textContent = '修改需求';
    modifyBtn.onclick = onModify;

    btnRow.appendChild(confirmBtn);
    btnRow.appendChild(modifyBtn);
    container.appendChild(btnRow);

    return container;
}

// 交互式组件：执行计划
function createPlanConfirmation(plan, onConfirm) {
    const container = document.createElement('div');
    container.className = 'bg-slate-800/30 rounded-lg border border-slate-700/50 p-4 mt-2 space-y-3';

    // Summary
    const summary = document.createElement('div');
    summary.innerHTML = `<h4 class="text-white font-medium mb-1">计划摘要</h4><p class="text-sm text-slate-400">${plan.summary}</p>`;
    container.appendChild(summary);

    // List
    if (plan.operations?.length > 0) {
        const ul = document.createElement('ul');
        ul.className = 'space-y-2 mt-2';
        plan.operations.forEach(op => {
            const li = document.createElement('li');
            li.className = 'flex items-start gap-2 text-sm text-slate-300 bg-slate-900/40 p-2 rounded';
            li.innerHTML = `<i data-lucide="chevron-right" class="w-4 h-4 text-blue-500 mt-0.5"></i> <span>${op.description || op.type}</span>`;
            ul.appendChild(li);
        });
        container.appendChild(ul);
    }

    // Button
    const btn = document.createElement('button');
    btn.className = 'w-full py-2.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg text-sm font-medium mt-2 transition-all shadow-lg shadow-emerald-900/20 flex items-center justify-center gap-2';
    btn.innerHTML = '<i data-lucide="play" class="w-4 h-4"></i> 立即执行';
    btn.onclick = onConfirm;

    container.appendChild(btn);
    // Be sure to init icons for the new content
    setTimeout(() => lucide.createIcons({ root: container }), 0);

    return container;
}


// ============ 核心逻辑 ============

async function handleFileUpload(file) {
    const ext = file.name.split('.').pop().toLowerCase();
    if (!['xlsx', 'xls'].includes(ext)) {
        alert('只支持 .xlsx 和 .xls 格式的文件');
        return null;
    }

    elements.uploadProgress.classList.remove('hidden');
    elements.progressFill.style.width = '30%';
    elements.progressText.textContent = `正在上传 ${file.name}...`;

    try {
        elements.progressFill.style.width = '60%';
        const result = await api.upload(file);
        elements.progressFill.style.width = '100%';

        return {
            fileId: result.file_id,
            metadata: result.metadata,
            originalName: result.metadata.file_name
        };
    } catch (error) {
        elements.progressText.textContent = `失败: ${error.message}`;
        elements.progressFill.classList.add('bg-red-500');
        return null;
    }
}

// 处理多文件上传
async function handleMultiFileUpload(files) {
    const fileArray = Array.from(files);

    for (let i = 0; i < fileArray.length; i++) {
        const file = fileArray[i];
        elements.progressText.textContent = `上传中 (${i + 1}/${fileArray.length}): ${file.name}`;

        const result = await handleFileUpload(file);
        if (result) {
            state.files.push(result);
        }
    }

    elements.uploadProgress.classList.add('hidden');
    elements.progressFill.style.width = '0%';

    if (state.files.length > 0) {
        state.isMultiFileMode = state.files.length > 1;

        if (state.isMultiFileMode) {
            // 多文件模式：显示文件列表
            renderUploadedFilesList();
            elements.uploadedFilesList.classList.remove('hidden');
            elements.multiFileActions.classList.remove('hidden');
            elements.uploadArea.classList.add('hidden');
            lucide.createIcons();
        } else {
            // 单文件模式：直接进入聊天
            const fileInfo = state.files[0];
            state.fileId = fileInfo.fileId;
            state.metadata = fileInfo.metadata;
            updateFileInfo(fileInfo.metadata);
            showSection('chat');
            addMessage('assistant', `👋 文件 **${fileInfo.metadata.file_name}** 已就绪。\n\n请告诉我您想如何处理这个表格？`);
        }
    }
}

// 渲染已上传文件列表
function renderUploadedFilesList() {
    elements.uploadedFilesList.innerHTML = '';

    state.files.forEach((fileInfo, index) => {
        const card = document.createElement('div');
        card.className = 'glass-panel rounded-xl p-4 animate-fade-in';

        const metadata = fileInfo.metadata;
        const sheet = metadata.sheets[0];
        const headers = sheet?.columns?.slice(0, 10) || [];

        card.innerHTML = `
            <div class="flex items-start justify-between mb-3">
                <div class="flex items-center gap-3">
                    <div class="w-10 h-10 rounded-lg bg-blue-500/10 border border-blue-500/20 flex items-center justify-center text-blue-400">
                        <i data-lucide="file-spreadsheet" class="w-5 h-5"></i>
                    </div>
                    <div>
                        <h4 class="font-medium text-white text-sm">${metadata.file_name}</h4>
                        <p class="text-xs text-slate-500">${sheet?.total_rows || 0} 行 | ${sheet?.total_cols || 0} 列</p>
                    </div>
                </div>
                <button class="btn-remove-uploaded-file p-1.5 text-slate-500 hover:text-red-400 transition-colors" data-index="${index}">
                    <i data-lucide="x" class="w-4 h-4"></i>
                </button>
            </div>
            <div class="flex flex-wrap gap-1.5">
                ${headers.map(col => `
                    <span class="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs bg-slate-800/50 border border-slate-700/50 text-slate-400">
                        ${col.data_type === '数字' ? '🔢' : col.data_type === '日期' ? '📅' : '📝'}${col.name}
                    </span>
                `).join('')}
                ${sheet?.columns?.length > 10 ? `<span class="text-xs text-blue-400">+${sheet.columns.length - 10} 列</span>` : ''}
            </div>
        `;

        elements.uploadedFilesList.appendChild(card);
    });

    // 绑定删除按钮事件
    document.querySelectorAll('.btn-remove-uploaded-file').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const index = parseInt(e.currentTarget.dataset.index);
            state.files.splice(index, 1);
            if (state.files.length === 0) {
                elements.uploadedFilesList.classList.add('hidden');
                elements.multiFileActions.classList.add('hidden');
                elements.uploadArea.classList.remove('hidden');
            } else {
                renderUploadedFilesList();
            }
            lucide.createIcons();
        });
    });

    lucide.createIcons();
}

// 开始处理多文件（进入聊天）
function startMultiFileProcessing() {
    // 使用第一个文件作为主文件
    const primaryFile = state.files[0];
    state.fileId = primaryFile.fileId;
    state.metadata = primaryFile.metadata;

    updateFileInfo(primaryFile.metadata);
    showSection('chat');

    // 生成多文件说明消息
    const fileNames = state.files.map(f => `**${f.metadata.file_name}**`).join('、');
    const msg = state.files.length > 1
        ? `📁 已加载 ${state.files.length} 个文件：${fileNames}\n\n请告诉我您想如何处理这些文件？例如：\n- "合并这些表格"\n- "按订单号匹配合并"\n- "纵向追加所有数据"`
        : `👋 文件 **${primaryFile.metadata.file_name}** 已就绪。\n\n请告诉我您想如何处理这个表格？`;

    addMessage('assistant', msg);
}

async function handleSendMessage() {
    const input = elements.userInput.value.trim();
    if (!input || state.isProcessing) return;
    
    // 🔒 防止短时间内重复提交（限流保护）
    const now = Date.now();
    if (state.lastSubmitTime && (now - state.lastSubmitTime) < 1000) {
        console.warn('⚠️ 请求过于频繁，已忽略');
        return;
    }
    state.lastSubmitTime = now;
    
    // 🔍 循环检测：如果连续 5 次相同输入，警告用户
    if (state.lastUserInput === input) {
        state.sameInputCount = (state.sameInputCount || 0) + 1;
        if (state.sameInputCount >= 3) {
            const continueAnyway = confirm('⚠️ 检测到您连续提交了相同的请求。\n\n可能原因：\n1. 智能助手理解有误\n2. API 配置问题\n3. 网络延迟\n\n建议您：\n- 尝试换个方式描述需求\n- 检查 API 配置\n- 查看浏览器控制台日志\n\n是否继续提交？');
            if (!continueAnyway) {
                state.sameInputCount = 0;
                return;
            }
        }
    } else {
        state.sameInputCount = 0;
    }
    
    if (!input || state.isProcessing) return;

    state.isProcessing = true;
    state.currentRequestId++;
    const thisRequestId = state.currentRequestId;

    // 保存用户输入用于重试
    state.lastUserInput = input;

    elements.btnSend.disabled = true;
    elements.userInput.disabled = true;

    addMessage('user', input);
    elements.userInput.value = '';

    addTypingIndicator();

    try {
        // 收集所有文件ID用于多文件场景
        const fileIds = state.files.map(f => f.fileId);
        // 传递上一次操作计划用于上下文
        const response = await api.refine(state.fileId, input, state.sessionId, state.currentAnswers, fileIds, state.lastOperationPlan);

        if (state.currentRequestId !== thisRequestId) return;
        removeTypingIndicator();
        state.sessionId = response.session_id;

        if (response.status === 'need_clarification') {
            addMessage('assistant', response.message);
            if (response.questions?.length) {
                const qBlock = createQuestionBlock(response.questions, handleAnswerConfirm, () => {
                    elements.userInput.focus();
                });
                addMessage('assistant', qBlock);
            }
        } else if (response.status === 'ready') {
            // 保存操作计划用于继续编辑时的上下文
            state.lastOperationPlan = response.operation_plan;
            const planBlock = createPlanConfirmation(response.operation_plan, executeProcessing);
            addMessage('assistant', planBlock);
        } else if (response.status === 'error') {
            // 使用带重试按钮的错误消息
            const errorBlock = createErrorWithRetry(response.message || '处理请求时出错', retryLastMessage);
            addMessage('assistant', errorBlock);
        } else {
            addMessage('assistant', response.message || '我不确定如何处理，请重试。');
        }

    } catch (error) {
        removeTypingIndicator();
        console.error('❌ [API 请求错误]', error);
        
        // 🎯 更友好的错误提示，区分不同错误类型
        let errorMessage = '请求失败';
        let errorDetails = error.message || '未知错误';
        
        if (error.message.includes('Failed to fetch') || error.message.includes('NetworkError')) {
            errorMessage = '网络连接失败';
            errorDetails = '可能原因：\n• 网络不稳定\n• 服务器未响应\n• 跨域问题\n\n建议：检查网络连接，或稍后重试。';
        } else if (error.message.includes('timeout')) {
            errorMessage = '请求超时';
            errorDetails = 'LLM API 响应时间过长。\n建议：切换到响应更快的 API 配置。';
        } else if (error.message.includes('API') || error.message.includes('401') || error.message.includes('403')) {
            errorMessage = 'API 配置错误';
            errorDetails = 'API Key 或配置可能有问题。\n建议：点击"切换 API"重新配置。';
        } else if (error.message.includes('JSON') || error.message.includes('parse')) {
            errorMessage = 'LLM 响应格式异常';
            errorDetails = '智能助手返回了无效的数据格式。\n建议：重试或切换不同的 LLM 模型。';
        }
        
        const errorBlock = createErrorWithRetry(
            `<strong>${errorMessage}</strong><br><span class="text-xs">${errorDetails}</span>`, 
            retryLastMessage
        );
        addMessage('assistant', errorBlock);
    } finally {
        state.isProcessing = false;
        elements.btnSend.disabled = false;
        elements.userInput.disabled = false;
        elements.userInput.focus();
    }
}

// 重试最后一次消息（由 createErrorWithRetry 的重试按钮调用）
function retryLastMessage() {
    if (state.lastUserInput) {
        // 移除错误消息和之前的用户消息（避免重复）
        if (state.lastAssistantMessageEl) {
            state.lastAssistantMessageEl.remove();
            state.lastAssistantMessageEl = null;
        }
        if (state.lastUserMessageEl) {
            state.lastUserMessageEl.remove();
            state.lastUserMessageEl = null;
        }
        elements.userInput.value = state.lastUserInput;
        handleSendMessage();
    }
}

async function handleAnswerConfirm() {
    if (Object.keys(state.currentAnswers).length === 0) {
        alert('请先选择选项');
        return;
    }

    state.isProcessing = true;
    state.currentRequestId++;
    const thisRequestId = state.currentRequestId;

    // 保存当前回答用于重试
    const savedAnswers = { ...state.currentAnswers };

    addTypingIndicator();

    try {
        const fileIds = state.files.map(f => f.fileId);
        const response = await api.refine(state.fileId, '用户已确认', state.sessionId, state.currentAnswers, fileIds);
        if (state.currentRequestId !== thisRequestId) return;

        removeTypingIndicator();
        state.sessionId = response.session_id;

        if (response.status === 'ready') {
            const planBlock = createPlanConfirmation(response.operation_plan, executeProcessing);
            addMessage('assistant', planBlock);
        } else if (response.status === 'error') {
            // 使用带重试按钮的错误消息
            const errorBlock = createErrorWithRetry(response.message || '处理请求时出错', () => {
                state.currentAnswers = savedAnswers;
                handleAnswerConfirm();
            });
            addMessage('assistant', errorBlock);
        } else {
            addMessage('assistant', response.message);
            if (response.questions) {
                const qBlock = createQuestionBlock(response.questions, handleAnswerConfirm, () => elements.userInput.focus());
                addMessage('assistant', qBlock);
            }
        }
    } catch (e) {
        removeTypingIndicator();
        // 使用带重试按钮的错误消息
        const errorBlock = createErrorWithRetry(`请求失败: ${e.message}`, () => {
            state.currentAnswers = savedAnswers;
            handleAnswerConfirm();
        });
        addMessage('assistant', errorBlock);
    } finally {
        state.isProcessing = false;
        state.currentAnswers = {};
    }
}

async function executeProcessing() {
    elements.statusCard.classList.remove('hidden');
    elements.statusTitle.textContent = 'AI 正在处理...';
    elements.statusDesc.textContent = '正在执行您的Excel操作计划';
    elements.statusIcon.classList.add('animate-spin');
    elements.actionCard.classList.add('hidden');

    try {
        const result = await api.process(state.fileId, state.sessionId);

        if (result.success) {
            elements.statusTitle.textContent = '处理完成!';
            elements.statusDesc.textContent = result.summary || '操作已成功执行';
            elements.statusIcon.classList.remove('animate-spin');
            // 使用 statusIconBg 而不是 parentElement（避免 null 引用）
            elements.statusIconBg.classList.remove('bg-blue-500/20');
            elements.statusIconBg.classList.add('bg-emerald-500/20');
            elements.statusIcon.classList.remove('text-blue-400');
            elements.statusIcon.classList.add('text-emerald-400');
            elements.statusIcon.setAttribute('data-lucide', 'check-circle');
            // 添加成功脉冲动画
            elements.statusIconBg.classList.add('animate-success-pulse');
            lucide.createIcons();

            elements.actionCard.classList.remove('hidden');
            elements.btnDownload.classList.remove('hidden');
            elements.btnContinue.classList.remove('hidden');

            state.downloadUrl = result.download_url;
            addMessage('assistant', '✅ 处理完成！您可以下载文件或继续操作。');
        } else {
            throw new Error(result.message);
        }
    } catch (error) {
        elements.statusTitle.textContent = '处理失败';
        elements.statusDesc.textContent = error.message;
        elements.statusIcon.classList.remove('animate-spin');
        elements.statusIcon.setAttribute('data-lucide', 'alert-triangle');
        // 使用 remove + add 而不是 replace（避免兼容性问题）
        elements.statusIconBg.classList.remove('bg-blue-500/20');
        elements.statusIconBg.classList.add('bg-red-500/20');
        elements.statusIcon.classList.remove('text-blue-400');
        elements.statusIcon.classList.add('text-red-400');
        lucide.createIcons();
        addMessage('assistant', `❌ 处理失败: ${error.message}`);
    }
}


// ============ 事件监听 ============

// Drag & Drop
elements.uploadArea.addEventListener('dragover', (e) => { e.preventDefault(); elements.uploadArea.classList.add('border-blue-500'); });
elements.uploadArea.addEventListener('dragleave', () => { elements.uploadArea.classList.remove('border-blue-500'); });
elements.uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    elements.uploadArea.classList.remove('border-blue-500');
    if (e.dataTransfer.files.length) handleMultiFileUpload(e.dataTransfer.files);
});
elements.uploadArea.addEventListener('click', () => elements.fileInput.click());
elements.fileInput.addEventListener('change', (e) => { if (e.target.files.length) handleMultiFileUpload(e.target.files); });

// Multi-file actions
elements.btnAddMoreFiles?.addEventListener('click', () => elements.fileInput.click());
elements.btnStartProcessing?.addEventListener('click', startMultiFileProcessing);

// Chat
elements.btnSend.addEventListener('click', handleSendMessage);
elements.userInput.addEventListener('keydown', (e) => { if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); handleSendMessage(); } });

// Actions
elements.btnRemoveFile.addEventListener('click', () => {
    if (confirm('确定要移除当前文件吗?')) showSection('upload');
});
elements.btnDownload.addEventListener('click', () => { if (state.downloadUrl) window.location.href = state.downloadUrl; });
elements.btnNewTask.addEventListener('click', () => showSection('upload'));
elements.btnContinue.addEventListener('click', async () => {
    const outputFileId = state.downloadUrl.split('/').pop();
    showLoading('正在加载新文件...');
    try {
        const result = await api.continueProcessing(outputFileId);
        state.fileId = result.file_id;
        state.metadata = result.metadata;
        // 必须清空 sessionId，因为后端的 session 在处理完成后已被清理
        // 下次发送消息时会为新文件创建新的 session
        state.sessionId = null;
        state.downloadUrl = null;
        // 重置多文件状态 - 继续编辑时只有一个文件
        state.files = [{
            fileId: result.file_id,
            metadata: result.metadata,
            originalName: result.metadata.file_name
        }];
        state.isMultiFileMode = false;
        state.currentFileIndex = 0;

        updateFileInfo(result.metadata);
        hideLoading()
        showSection('chat');
        // 保留聊天记录作为 UI 上下文展示，添加分隔线
        addContinueSessionDivider();
        addMessage('assistant', `📁 已加载处理后的文件 **${result.metadata.file_name}**\n\n您可以继续对这个文件进行操作，请告诉我接下来需要做什么？`);
    } catch (e) {
        hideLoading();
        alert(e.message);
    }
});

// Settings Modal
const settingsModal = document.getElementById('settings-modal');
document.getElementById('btn-settings').addEventListener('click', () => {
    settingsModal.classList.remove('hidden');
    if (typeof loadAllConfigs === 'function') loadAllConfigs();
});
document.getElementById('btn-close-settings').addEventListener('click', () => settingsModal.classList.add('hidden'));

// Settings Elements
const settingsElements = {
    modal: settingsModal,
    configsList: document.getElementById('configs-list'),
    configForm: document.getElementById('config-form'),
    btnAdd: document.getElementById('btn-add-config'),
    btnSave: document.getElementById('btn-save-config'),
    btnCancel: document.getElementById('btn-cancel-edit'),
    // Inputs
    name: document.getElementById('config-name'),
    base: document.getElementById('config-api-base'),
    key: document.getElementById('config-api-key'),
    model: document.getElementById('config-model'),
    isDefault: document.getElementById('config-is-default'),
    editId: document.getElementById('edit-config-id'),
    status: document.getElementById('connection-status'),
    btnFetch: document.getElementById('btn-fetch-models'),
    btnTest: document.getElementById('btn-test-connection')
};

// Toggle logic
settingsElements.btnAdd.addEventListener('click', () => {
    settingsElements.configsList.parentElement.classList.add('hidden');
    settingsElements.configForm.classList.remove('hidden');
    // Clear form
    settingsElements.editId.value = '';
    settingsElements.name.value = '';
    settingsElements.base.value = 'https://api.openai.com/v1';
    settingsElements.key.value = '';
    settingsElements.model.innerHTML = '<option>请先获取模型</option>';
});

settingsElements.btnCancel.addEventListener('click', () => {
    settingsElements.configForm.classList.add('hidden');
    settingsElements.configsList.parentElement.classList.remove('hidden');
});

// Load Configs (Updated with state logic)
const settingsState = {
    configs: [],
    editingId: null,
    isEditing: false
};

async function loadAllConfigs() {
    try {
        const res = await fetch('/api/configs');
        const data = await res.json();
        settingsState.configs = data.configs || [];
        renderConfigsList();
    } catch (e) { console.error(e); }
}

// 渲染配置列表 (Event Delegation version)
function renderConfigsList() {
    if (settingsState.configs.length === 0) {
        settingsElements.configsList.innerHTML = '<p class="empty-message text-slate-400 text-center py-4">还没有保存任何配置，点击上方按钮添加</p>';
        return;
    }

    settingsElements.configsList.innerHTML = '';
    settingsState.configs.forEach(config => {
        const card = document.createElement('div');
        card.className = 'config-card bg-slate-800/50 p-3 rounded-lg border border-slate-700 mb-2';
        if (config.is_default) {
            card.classList.add('ring-1', 'ring-blue-500/50');
        }

        card.innerHTML = `
            <div class="flex justify-between items-start">
                <div class="config-info">
                    <h4 class="font-medium text-slate-200 text-sm flex items-center gap-2">
                        ${config.name}
                        ${config.is_default ? '<span class="text-[10px] bg-blue-900/50 text-blue-300 px-1.5 py-0.5 rounded border border-blue-500/20">默认</span>' : ''}
                    </h4>
                    <p class="text-xs text-slate-500 mt-1">${config.model}</p>
                </div>
                <div class="flex gap-2">
                    <button type="button" class="btn-icon btn-edit text-slate-400 hover:text-white transition-colors" data-id="${config.id}" title="编辑">✏️</button>
                    ${!config.is_default ? `<button type="button" class="btn-icon btn-default text-slate-400 hover:text-yellow-400 transition-colors" data-id="${config.id}" title="设为默认">⭐</button>` : ''}
                    <button type="button" class="btn-icon btn-delete text-slate-400 hover:text-red-400 transition-colors" data-id="${config.id}" title="删除">🗑️</button>
                </div>
            </div>
            <div class="text-[10px] text-slate-600 mt-2 font-mono truncate">
                ${config.api_base}
            </div>
        `;

        settingsElements.configsList.appendChild(card);
    });
}

// 事件委托处理配置列表点击
settingsElements.configsList.addEventListener('click', async (e) => {
    // 向上寻找 button
    const btn = e.target.closest('button');
    if (!btn) return;

    // Prevent any default form submission or bubbling
    e.preventDefault();
    e.stopPropagation();

    const id = btn.dataset.id;
    if (!id) return;

    if (btn.classList.contains('btn-edit')) {
        await editConfig(id);
    } else if (btn.classList.contains('btn-delete')) {
        await deleteConfig(id);
    } else if (btn.classList.contains('btn-default')) {
        await setDefaultConfig(id);
    }
});

function showConfigForm(configId) {
    settingsElements.configsList.parentElement.classList.add('hidden');
    settingsElements.configForm.classList.remove('hidden');

    // Find config
    const config = settingsState.configs.find(c => c.id === configId);
    if (config) {
        settingsElements.editId.value = config.id;
        settingsElements.name.value = config.name;
        settingsElements.base.value = config.api_base;
        settingsElements.key.value = ''; // Don't show
        settingsElements.key.placeholder = '保留原密钥';
        settingsElements.model.innerHTML = `<option value="${config.model}">${config.model}</option>`;
        settingsElements.isDefault.checked = config.is_default;
    }
}

// Local helper functions for actions
async function editConfig(configId) {
    showConfigForm(configId);
}

async function deleteConfig(configId) {
    // No timeout needed with proper event handling
    if (!confirm('确定要删除这个配置吗？')) return;

    try {
        const response = await fetch(`/api/configs/${configId}`, { method: 'DELETE' });
        const result = await response.json();
        if (result.success) await loadAllConfigs();
        else alert('删除失败: ' + result.message);
    } catch (error) { alert('删除失败: ' + error.message); }
}

async function setDefaultConfig(configId) {
    try {
        const response = await fetch(`/api/configs/${configId}/set-default`, { method: 'POST' });
        const result = await response.json();
        if (result.success) await loadAllConfigs();
        else alert('设置失败: ' + result.message);
    } catch (error) { alert('设置失败: ' + error.message); }
}

// Save
settingsElements.btnSave.addEventListener('click', async () => {
    const id = settingsElements.editId.value;
    const body = {
        name: settingsElements.name.value,
        api_base: settingsElements.base.value,
        model: settingsElements.model.value,
        set_as_default: settingsElements.isDefault.checked,
        is_default: settingsElements.isDefault.checked
    };
    if (settingsElements.key.value) body.api_key = settingsElements.key.value;

    const method = id ? 'PUT' : 'POST';
    const url = id ? `/api/configs/${id}` : '/api/configs';

    settingsElements.status.classList.remove('hidden', 'success', 'error');
    settingsElements.status.textContent = '保存中...';

    try {
        const res = await fetch(url, {
            method,
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(body)
        });
        const data = await res.json();
        if (data.success) {
            settingsElements.status.textContent = '保存成功';
            settingsElements.status.classList.add('success');
            setTimeout(() => {
                settingsElements.status.classList.add('hidden');
                settingsElements.configForm.classList.add('hidden');
                settingsElements.configsList.parentElement.classList.remove('hidden');
                loadAllConfigs();
            }, 500);
        } else {
            settingsElements.status.textContent = '失败: ' + data.message;
            settingsElements.status.classList.add('error');
        }
    } catch (e) {
        settingsElements.status.textContent = '错误: ' + e.message;
        settingsElements.status.classList.add('error');
    }
});

// Fetch Models
settingsElements.btnFetch.addEventListener('click', async () => {
    try {
        const res = await fetch('/api/models', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                api_base: settingsElements.base.value,
                api_key: settingsElements.key.value
            })
        });
        const data = await res.json();
        if (data.success) {
            settingsElements.model.innerHTML = '';
            data.models.forEach(m => {
                const opt = document.createElement('option');
                opt.value = m.id;
                opt.textContent = m.name;
                settingsElements.model.appendChild(opt);
            });
        } else {
            alert(data.message);
        }
    } catch (e) { alert(e.message); }
});

// Test Connection
settingsElements.btnTest.addEventListener('click', async () => {
    if (!settingsElements.key.value || !settingsElements.base.value || !settingsElements.model.value) {
        alert('请先填写 API 地址、API Key 和选择模型');
        return;
    }

    settingsElements.status.classList.remove('hidden', 'success', 'error');
    settingsElements.status.textContent = '测试中...';

    try {
        const res = await fetch('/api/test-connection', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                api_base: settingsElements.base.value,
                api_key: settingsElements.key.value,
                model: settingsElements.model.value
            })
        });
        const data = await res.json();
        if (data.success) {
            settingsElements.status.textContent = '✓ ' + data.message;
            settingsElements.status.classList.add('success');
        } else {
            settingsElements.status.textContent = '✗ ' + data.message;
            settingsElements.status.classList.add('error');
        }
    } catch (e) {
        settingsElements.status.textContent = '✗ 连接失败: ' + e.message;
        settingsElements.status.classList.add('error');
    }
});


// Init
document.addEventListener('DOMContentLoaded', () => {
    showSection('upload');
    lucide.createIcons();
    loadAllConfigs();
});
