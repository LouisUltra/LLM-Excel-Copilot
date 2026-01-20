# 贡献指南 | Contributing Guide

感谢你对 LLM-Excel-Copilot 的关注！

## 🤝 如何贡献

### 报告 Bug
- 使用 GitHub Issues 提交 Bug
- 描述清楚复现步骤
- 提供错误日志和截图

### 提交功能建议
- 先检查 Issues 中是否已有类似建议
- 描述功能的使用场景和价值
- 欢迎提供设计方案

### 提交代码
1. Fork 本仓库
2. 创建你的特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交你的改动 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 创建 Pull Request

## 🧪 开发环境

```bash
# 1. 克隆项目
git clone https://github.com/LouisUltra/LLM-Excel-Copilot.git
cd LLM-Excel-Copilot

# 2. 创建虚拟环境
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate

# 3. 安装依赖
pip install -r requirements.txt

# 4. 配置环境变量
cp .env.example .env
# 编辑 .env 文件，填入你的 API Key

# 5. 运行测试
pytest tests/

# 6. 启动开发服务器
python -m app.main
```

## 📝 代码规范

- 遵循 PEP 8 Python 代码风格
- 使用有意义的变量名和函数名
- 添加必要的注释和文档字符串
- 新功能需要添加对应的单元测试

## 🔍 提交前检查

- [ ] 代码通过 `pytest` 测试
- [ ] 代码符合 PEP 8 规范
- [ ] 添加了必要的注释
- [ ] 更新了相关文档

## 💡 开发建议

### 项目架构
- `app/main.py` - FastAPI 应用入口
- `app/core/` - 核心业务逻辑
  - `excel_parser.py` - Excel 解析
  - `excel_executor.py` - Excel 操作执行
  - `llm_client.py` - LLM API 交互
  - `requirement_refiner.py` - 需求精化
- `app/static/` - 前端资源
- `tests/` - 单元测试

### 常见开发场景

**添加新的 Excel 操作**：
1. 在 `models.py` 中的 `OperationType` 添加新类型
2. 在 `excel_executor.py` 中实现对应方法
3. 在 `llm_client.py` 的 `SYSTEM_PROMPT` 中添加操作说明
4. 添加单元测试

**优化 LLM 提示词**：
- 修改 `llm_client.py` 中的 `SYSTEM_PROMPT` 和 `REFINE_SYSTEM_PROMPT`
- 测试不同场景下的效果
- 记录优化前后的对比

## 📬 联系方式

如有问题，欢迎：
- 提交 Issue
- 发起 Discussion
- 联系维护者

再次感谢你的贡献！🎉
