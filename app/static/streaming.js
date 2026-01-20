/**
 * LLM 流式输出管理器
 * 实现类似 ChatGPT 的实时显示效果
 */

class StreamingManager {
    constructor() {
        this.activeStreams = new Map();
        this.streamingContainer = null;
        this.initStreamingUI();
    }

    /**
     * 初始化流式显示UI
     */
    initStreamingUI() {
        // 创建流式显示容器（浮动窗口）
        const container = document.createElement('div');
        container.className = 'streaming-container hidden';
        container.innerHTML = `
            <div class="streaming-header">
                <span class="streaming-title">🤖 AI 思考中...</span>
                <button class="btn-close-stream" title="最小化">−</button>
            </div>
            <div class="streaming-content" id="streaming-content"></div>
            <div class="streaming-status">
                <span class="status-text">连接中...</span>
                <div class="pulse-indicator"></div>
            </div>
        `;
        
        document.body.appendChild(container);
        this.streamingContainer = container;

        // 绑定最小化/还原按钮
        container.querySelector('.btn-close-stream').addEventListener('click', () => {
            this.toggleMinimize();
        });
    }

    /**
     * 显示流式窗口
     */
    show() {
        this.streamingContainer.classList.remove('hidden');
        this.streamingContainer.classList.add('streaming-active');
        console.log('📺 流式窗口已显示');
    }

    /**
     * 隐藏流式窗口
     */
    hide() {
        this.streamingContainer.classList.add('hidden');
        this.streamingContainer.classList.remove('streaming-active');
        console.log('📺 流式窗口已隐藏');
    }

    /**
     * 最小化窗口
     */
    minimize() {
        this.streamingContainer.classList.add('minimized');
        // 更新按钮为还原图标
        const btn = this.streamingContainer.querySelector('.btn-close-stream');
        btn.textContent = '□';
        btn.title = '还原';
    }

    /**
     * 还原窗口
     */
    restore() {
        this.streamingContainer.classList.remove('minimized');
        // 更新按钮为最小化图标
        const btn = this.streamingContainer.querySelector('.btn-close-stream');
        btn.textContent = '−';
        btn.title = '最小化';
    }

    /**
     * 切换最小化/还原
     */
    toggleMinimize() {
        if (this.streamingContainer.classList.contains('minimized')) {
            this.restore();
        } else {
            this.minimize();
        }
    }

    /**
     * 更新状态
     */
    updateStatus(text, type = 'info') {
        const statusText = this.streamingContainer.querySelector('.status-text');
        statusText.textContent = text;
        statusText.className = `status-text status-${type}`;
    }

    /**
     * 添加消息（支持打字机效果）
     */
    addMessage(text, options = {}) {
        const { 
            type = 'assistant', 
            streaming = false,
            id = Date.now().toString()
        } = options;

        const contentDiv = this.streamingContainer.querySelector('.streaming-content');
        
        // 检查是否已存在该消息
        let messageDiv = contentDiv.querySelector(`[data-message-id="${id}"]`);
        
        if (!messageDiv) {
            messageDiv = document.createElement('div');
            messageDiv.className = `stream-message stream-${type}`;
            messageDiv.setAttribute('data-message-id', id);
            contentDiv.appendChild(messageDiv);
        }

        if (streaming) {
            // 打字机效果
            messageDiv.textContent = text;
        } else {
            messageDiv.innerHTML = text;
        }

        // 自动滚动到底部
        contentDiv.scrollTop = contentDiv.scrollHeight;

        return id;
    }

    /**
     * 更新消息（用于流式追加）
     */
    updateMessage(id, text) {
        const messageDiv = this.streamingContainer.querySelector(`[data-message-id="${id}"]`);
        if (messageDiv) {
            messageDiv.textContent = text;
            
            // 自动滚动
            const contentDiv = this.streamingContainer.querySelector('.streaming-content');
            contentDiv.scrollTop = contentDiv.scrollHeight;
        }
    }

    /**
     * 清空内容
     */
    clear() {
        const contentDiv = this.streamingContainer.querySelector('.streaming-content');
        contentDiv.innerHTML = '';
    }

    /**
     * 添加进度信息
     */
    addProgress(current, total, description = '') {
        const progressHtml = `
            <div class="stream-progress">
                <div class="progress-bar-wrapper">
                    <div class="progress-bar-fill" style="width: ${(current/total*100)}%"></div>
                </div>
                <div class="progress-text">
                    ${description} (${current}/${total})
                </div>
            </div>
        `;
        
        this.addMessage(progressHtml, { type: 'system', id: 'progress' });
    }

    /**
     * 模拟 LLM 思考过程
     */
    simulateThinking(stage = 1) {
        const stages = [
            '🔍 正在分析 Excel 文件结构...',
            '🤔 正在理解您的需求...',
            '📋 正在生成操作计划...',
            '✨ 即将完成...'
        ];

        const text = stages[Math.min(stage - 1, stages.length - 1)];
        this.updateStatus(text, 'thinking');
    }

    /**
     * 显示错误
     */
    showError(error) {
        this.addMessage(`❌ 错误: ${error}`, { type: 'error' });
        this.updateStatus('发生错误', 'error');
    }

    /**
     * 显示完成
     */
    showComplete() {
        this.updateStatus('✅ 完成', 'success');
        
        // 3秒后自动隐藏
        setTimeout(() => {
            this.hide();
        }, 3000);
    }
}

// 创建全局实例
const streamingManager = new StreamingManager();

// 导出到全局
window.streamingManager = streamingManager;
