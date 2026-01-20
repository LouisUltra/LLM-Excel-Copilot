"""
LLM 客户端模块
封装与大语言模型的交互，支持多种兼容 OpenAI 格式的 API
"""

import json
import re
from typing import List, Dict, Any, Optional, Generator
from openai import OpenAI

from app.config import settings
from app.models import Operation, OperationPlan, OperationType


# 系统提示词：定义 LLM 的角色和能力
SYSTEM_PROMPT = """你是一个专业的 Excel 操作专家助手。用户会给你一个 Excel 文件的结构信息（不包含具体数据内容），以及他们想要进行的操作描述。

## 你的任务

1. **理解用户意图**：分析用户的需求，即使描述模糊也要尝试理解
2. **生成操作指令**：返回结构化的 JSON 操作指令，供本地脚本执行

## 可用操作类型

你可以使用以下操作类型（type 字段的值）：

### 数据筛选与排序
- `FILTER`: 条件筛选（保留满足条件的行，删除不满足的行）
  - params: {"column": "列名", "operator": "eq|ne|gt|lt|gte|lte|contains|startswith|endswith", "value": "值"}
  - 示例：筛选出"备注"列包含"未挂网"的行 → {"column": "备注", "operator": "contains", "value": "未挂网"}
  - 注意：FILTER 会**保留**满足条件的行，**删除**其他所有行
- `SORT`: 排序
  - params: {"column": "列名", "order": "asc|desc"}

### 列操作
- `ADD_COLUMN`: 新增列
  - params: {"name": "新列名", "formula": "Excel公式,如=A2+B2", "position": "after:列名|before:列名|end"}
- `DELETE_COLUMN`: 删除列
  - params: {"columns": ["列名1", "列名2"]}
- `SPLIT_COLUMN`: 拆分列
  - params: {"column": "列名", "delimiter": "分隔符", "new_columns": ["新列1", "新列2"]}
- `MERGE_COLUMNS`: 合并列
  - params: {"columns": ["列1", "列2"], "new_name": "合并后列名", "delimiter": "连接符"}

### 行操作
- `DELETE_ROWS`: 删除满足条件的行（与FILTER相反）
  - params: {"condition": {"column": "列名", "operator": "操作符", "value": "值"}}
  - 注意：DELETE_ROWS 会**删除**满足条件的行，**保留**其他行
  - 对于"筛选出X，删除其他"的需求，应使用 FILTER 而不是 DELETE_ROWS
- `DEDUPLICATE`: 去重
  - params: {"columns": ["用于判断重复的列"], "keep": "first|last"}

### 数据处理
- `REPLACE`: 替换
  - params: {"column": "列名", "old_value": "原值", "new_value": "新值", "regex": false}
- `FILL`: 填充空值
  - params: {"column": "列名", "method": "value|ffill|bfill", "value": "填充值(method为value时)"}
- `CALCULATE`: 计算汇总（在末尾添加汇总行）
  - params: {"operations": [{"column": "列名", "function": "sum|avg|count|max|min"}]}

### 格式化
- `FORMAT`: 数字/日期格式化（针对特定列）
  - params: {"column": "列名", "format_type": "number|date|percentage|currency", "format_string": "格式字符串"}
- `STYLE`: 样式设置（边框、背景色，针对整个区域）
  - params: {"style_type": "all|border|header", "range": "A1:L100(可选)", "header_row": 1, "border_style": "thin|medium|thick", "fill_color": "D9E1F2"}
  - style_type: all=边框+标题背景, border=仅边框, header=仅标题行样式

### 高级操作
- `VLOOKUP`: 跨表查找（仅用于同一工作簿内的不同工作表）
  - params: {"lookup_column": "查找列", "target_sheet": "目标表", "target_lookup_column": "目标查找列", "target_return_column": "返回值列", "new_column_name": "新列名"}
  - 注意：仅用于同一个 Excel 文件内的不同工作表之间的查找
- `PIVOT`: 数据透视
  - params: {"index": "行标签列", "columns": "列标签列", "values": "值列", "aggfunc": "sum|mean|count"}

### 多文件合并操作
- `MERGE_VERTICAL`: 纵向合并（将另一个文件的数据追加到当前表格下方）
  - params: {"source_file": "源文件路径", "source_sheet": "源工作表名(可选)", "skip_header": true}
  - 适用场景：两个文件结构相同，需要合并数据行
- `MERGE_HORIZONTAL`: 横向合并（按关键列匹配，将另一个文件的列添加到当前表格）
  - params: {
      "source_file": "源文件路径",
      "source_sheet": "源工作表名(可选)",
      "key_column": "当前表的关键列",
      "source_key_column": "源表的关键列(可选，默认与key_column相同)",
      "columns_to_add": ["要添加的列名1", "列名2"]  # 可选，不指定则添加所有非关键列
    }
  - 适用场景：两个文件有共同的关键字段（如姓名、ID），需要根据关键字段匹配并合并列
  - 注意：这是多文件场景下的推荐方法，而不是 VLOOKUP

### 图表操作
- `CREATE_CHART`: 创建图表（会在 Excel 中嵌入图表，也会生成独立的图片文件）
  - params: {
      "chart_type": "line|bar|pie|scatter|area|column",  # 图表类型
      "data_columns": ["列名1", "列名2"],  # 数据列（Y轴）
      "label_column": "列名",  # 标签列（X轴或分类，可选）
      "title": "图表标题",  # 图表标题
      "sheet_name": "图表_工作表名",  # 新建工作表名称（可选）
      "position": "existing|new_sheet",  # existing=嵌入当前表, new_sheet=新建工作表
      "width": 15,  # 图表宽度（英寸，默认15）
      "height": 10,  # 图表高度（英寸，默认10）
      "show_values": true|false  # 是否在图表上显示数据标签/数值（默认true）
    }
  - 注意：
    - line/column/bar 适合趋势和对比
    - pie 适合占比展示，只使用一列数据
    - scatter 适合相关性分析，需要两列数据
    - label_column 用于 X 轴标签，如果不提供则使用行号
    - show_values 控制是否在柱子/点上显示具体数值：用户说"显示数据标签/数值"时设为true，说"不要数据标签/隐藏数值"时设为false

## 响应格式

请以 JSON 格式返回操作计划，格式如下：

```json
{
  "operations": [
    {
      "type": "操作类型",
      "params": {"参数名": "参数值"},
      "description": "这个操作的中文描述",
      "target_sheet": "目标工作表名(可选,默认为活动工作表)"
    }
  ],
  "summary": "整体操作的简要描述",
  "estimated_impact": "预估影响,如'将删除约X行数据'"
}
```

## 重要原则

1. **只返回 JSON**：你的回复必须严格是上述格式的 JSON，不要有任何额外文字或解释

2. **严格的列名验证**：
   - **绝对不能臆想列名**！所有列名必须来自用户提供的 Excel 结构信息
   - 参考"示例值"来理解每列的实际内容
   - 如果不确定用户指的是哪一列，这应该在需求精化阶段就被询问清楚了
   - **列名必须完全匹配**：包括大小写、空格、括号等，如"销售额（元）"不等于"销售额"
   
   **智能处理"所有列"类表达**：
   - 当用户说"所有列"、"全部列"、"每一列"、"所有的列"、"all columns"、"every column" 等时：
     * ❌ **错误做法**：把"所有列"当成一个列名 → {"columns": ["所有列"]} 或 {"column": "all"}
     * ✅ **正确做法**：展开为实际的列名列表 → {"columns": ["姓名", "年龄", "销售额", ...]}
   - 应用场景示例：
     * 用户："对所有列设置边框" → 在STYLE操作中不需要指定具体列，用 range 参数覆盖整个区域
     * 用户："删除所有数值列" → 识别出所有数值类型的列，展开为 {"columns": ["销售额", "数量", ...]}
     * 用户："给全部列添加千分位" → 识别出所有数值类型的列，对每一列生成一个FORMAT操作
   - 如何判断"所有列"的范围：
     * 如果有修饰词（如"所有数值列"），则只包含符合条件的列（根据列的 data_type 字段）
     * 如果没有修饰词（纯"所有列"），则包含表格中的所有列
     * 可以根据上下文智能判断，例如"格式化所有列"通常指数值列

3. **使用实际列名**：操作中的列名必须与用户提供的表头完全一致（包括括号、空格等）

4. **拆分复杂操作**：如果用户的需求需要多个步骤，请按顺序列出多个 operation

5. **保守估计影响**：估计操作影响时要保守，宁可说"可能"而非绝对

6. **筛选操作的正确使用**：
   - "筛选出X/保留X/只要X"类需求 → 使用 FILTER 操作（保留满足条件的行）
   - "删除X/去掉X/移除X"类需求 → 使用 DELETE_ROWS 操作（删除满足条件的行）
   - 模糊匹配用 "contains"，精确匹配用 "eq"

7. **多文件场景使用 MERGE 操作**：
   - 两个文件结构相同 → MERGE_VERTICAL（纵向合并，追加行）
   - 两个文件有共同关键字段 → MERGE_HORIZONTAL（横向合并，按列匹配）
   - 不要使用 VLOOKUP 进行跨文件查找，VLOOKUP 仅用于同一工作簿内的不同工作表
   - source_file 参数会在执行时自动注入，你不需要指定具体路径

8. **常见错误及避免方法**：
   - ❌ 错误：ADD_COLUMN 的公式引用了新列自己 → ✅ 正确：公式只引用现有列
   - ❌ 错误：CALCULATE 的范围包含汇总行自己 → ✅ 正确：范围只到汇总行的上一行
   - ❌ 错误：图表的 data_columns 使用了不存在的列 → ✅ 正确：从文件结构中选择存在的列
   - ❌ 错误：对文本列使用数值运算 → ✅ 正确：检查列的数据类型

9. **图表创建最佳实践**：
   - 数值列用于 data_columns（如：销售额、数量）
   - 分类列用于 label_column（如：产品名称、地区）
   - position 默认用 "new_sheet"（创建新工作表，不影响原数据）
   - show_values 根据用户明确要求设置：说"显示数据标签"设为true，说"不要数据标签"设为false
   - 确保数据列是数值类型，否则图表可能为空

10. **公式操作注意事项**：
    - 支持的运算：+、-、*、/
    - 支持的函数：SUM、AVERAGE、COUNT、MAX、MIN
    - 公式必须使用Excel列字母（A、B、C...），不能使用列名
    - 示例：=C2*D2 表示第C列和第D列相乘
"""


# 需求精化的系统提示词
REFINE_SYSTEM_PROMPT = """你是一个友好的 Excel 操作助手。你的任务是帮助用户精确化他们的 Excel 处理需求。

用户可能会给出模糊的描述，你需要：
1. 理解他们的大致意图
2. 识别可能的歧义或缺失信息
3. 用友好的方式提出澄清问题

## Excel 文件信息

{file_description}

## 响应格式

请以 JSON 格式返回，格式如下：

```json
{{
  "status": "need_clarification 或 ready",
  "refined_requirement": "精化后的需求描述（用你的理解重新表述用户需求）",
  "questions": [
    {{
      "question_id": "q1",
      "question": "问题内容",
      "question_type": "single 或 multiple 或 text",
      "options": [
        {{"key": "a", "label": "选项A", "description": "选项说明(可选)"}},
        {{"key": "b", "label": "选项B", "description": ""}}
      ],
      "required": true
    }}
  ],
  "message": "给用户的友好消息"
}}
```

## 重要原则

1. **简洁友好**：问题要简洁明了，不要问太多问题（最多3个）
2. **提供选项**：尽量用选择题而非开放问题
3. **只问必要的**：如果用户需求已经很明确，设置 status 为 "ready" 并省略 questions
4. **严格使用实际列名**：
   - **永远不要臆想列名**！只能使用上述文件描述中明确列出的列名
   - 如果用户提到的列名在文件中不存在，**必须询问用户**指的是哪一列
   - 参考"示例值"列来理解每列的内容
   
   **🌟 智能理解"所有列"类表达**：
   - 当用户说"所有列"、"全部列"、"每一列"、"all columns" 等时，这**不是**一个具体的列名
   - 你应该理解为：用户想对表格中的所有列（或某一类列）进行操作
   - 精化需求时，应该明确说明"对表格中的所有列..."，而不是把"所有列"当成列名
   - 如果需要澄清，可以询问：
     * "您是指对表格中的所有列进行操作，还是特定的某几列？"
     * "您想操作所有列，还是只操作数值列/文本列？"
   - 例子：
     * 用户："格式化所有列" → refined_requirement: "对表格中的所有数值列应用千分位格式"
     * 用户："删除所有空列" → refined_requirement: "删除表格中内容全为空的列"

5. **只返回 JSON**：你的回复只能是 JSON 格式，不要有任何额外的解释文字
"""


class LLMClient:
    """
    LLM 客户端
    
    封装与大语言模型 API 的交互，支持：
    - OpenAI API
    - 通义千问 DashScope 兼容模式
    - DeepSeek API
    - 其他兼容 OpenAI 格式的 API
    """
    
    def __init__(
        self,
        api_key: Optional[str] = None,
        api_base: Optional[str] = None,
        model: Optional[str] = None
    ):
        """
        初始化 LLM 客户端
        
        Args:
            api_key: API 密钥，默认从配置读取
            api_base: API 基础地址，默认从配置读取
            model: 模型名称，默认从配置读取
        """
        self.api_key = api_key or settings.llm_api_key
        self.api_base = api_base or settings.llm_api_base
        self.model = model or settings.llm_model
        
        if not self.api_key:
            raise ValueError("LLM API Key 未配置，请在 .env 文件中设置 LLM_API_KEY")
        
        # 初始化 OpenAI 客户端
        self.client = OpenAI(
            api_key=self.api_key,
            base_url=self.api_base
        )
    
    def generate_operations(
        self,
        file_description: str,
        user_requirement: str,
        conversation_history: Optional[List[Dict[str, str]]] = None,
        max_retries: int = 2
    ) -> OperationPlan:
        """
        根据用户需求生成操作计划
        
        Args:
            file_description: Excel 文件结构描述
            user_requirement: 用户需求描述
            conversation_history: 对话历史
            max_retries: 最大重试次数
            
        Returns:
            OperationPlan: 操作计划
        """
        messages = [
            {"role": "system", "content": SYSTEM_PROMPT}
        ]
        
        # 添加文件描述作为上下文
        messages.append({
            "role": "user",
            "content": f"## Excel 文件结构\n\n{file_description}\n\n## 用户需求\n\n{user_requirement}"
        })
        
        last_error = None
        
        # 带重试的LLM调用
        for attempt in range(max_retries + 1):
            try:
                response = self.client.chat.completions.create(
                    model=self.model,
                    messages=messages,
                    temperature=0.3  # 降低随机性以获得更稳定的输出
                )
                
                content = response.choices[0].message.content
                
                # #region agent log
                import json
                from datetime import datetime
                with open('/Users/louis/PycharmProjects/Open Source/LLM-Excel-Copilot/.cursor/debug.log', 'a') as f:
                    f.write(json.dumps({"location":"llm_client.py:273","message":"llm_response_raw","data":{"content":content[:500]},"timestamp":datetime.now().timestamp()*1000,"sessionId":"debug-session","hypothesisId":"C,D,E"}) + '\n')
                # #endregion
                
                # 解析响应
                plan = self._parse_operation_plan(content)
                
                # #region agent log
                ops_summary = [{"type": op.type.value, "params": op.params, "desc": op.description} for op in plan.operations]
                with open('/Users/louis/PycharmProjects/Open Source/LLM-Excel-Copilot/.cursor/debug.log', 'a') as f:
                    f.write(json.dumps({"location":"llm_client.py:285","message":"operation_plan_parsed","data":{"operations":ops_summary},"timestamp":datetime.now().timestamp()*1000,"sessionId":"debug-session","hypothesisId":"C,D,E"}) + '\n')
                # #endregion
                
                # 验证操作计划
                if not plan.operations:
                    raise ValueError("操作计划为空，请重新生成")
                
                return plan
                
            except Exception as e:
                last_error = e
                if attempt < max_retries:
                    # 如果是解析错误，在下次请求中提示LLM
                    if "JSON" in str(e) or "解析" in str(e):
                        messages.append({
                            "role": "assistant",
                            "content": content if 'content' in locals() else ""
                        })
                        messages.append({
                            "role": "user",
                            "content": f"返回格式有误：{str(e)}。请严格按照JSON格式返回，不要有任何额外文字。"
                        })
                    continue
                else:
                    break
        
        # 所有重试都失败
        raise ValueError(f"生成操作计划失败（已重试{max_retries}次）: {str(last_error)}")
    
    def refine_requirement(
        self,
        file_description: str,
        user_input: str,
        answers: Optional[Dict[str, Any]] = None,
        conversation_history: Optional[List[Dict[str, str]]] = None
    ) -> Dict[str, Any]:
        """
        精化用户需求
        
        Args:
            file_description: Excel 文件结构描述
            user_input: 用户输入
            answers: 用户对之前问题的回答
            conversation_history: 对话历史
            
        Returns:
            dict: 精化结果
        """
        system_prompt = REFINE_SYSTEM_PROMPT.format(file_description=file_description)
        
        messages = [
            {"role": "system", "content": system_prompt}
        ]
        
        # 添加对话历史
        if conversation_history:
            messages.extend(conversation_history)
        
        # 构建用户消息
        user_message = user_input
        if answers:
            user_message += f"\n\n用户的回答：\n{json.dumps(answers, ensure_ascii=False, indent=2)}"
        
        messages.append({"role": "user", "content": user_message})
        
        # 调用 LLM（不使用 response_format，因为某些 API 不支持）
        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=messages,
                temperature=0.5
            )
        except Exception as e:
            print(f"❌ [LLM API 调用失败] {str(e)}")
            raise ValueError(f"LLM API 调用失败: {str(e)}")
        
        content = response.choices[0].message.content
        
        # 🔍 调试日志：记录 LLM 原始响应
        print(f"📋 [LLM 精化响应] 用户输入: {user_input[:50]}...")
        print(f"📋 [LLM 原始响应]:\n{content[:500]}..." if len(content) > 500 else f"📋 [LLM 原始响应]:\n{content}")
        
        # 解析响应并添加容错处理
        parsed_response = self._parse_json_response(content)
        
        # ✅ 验证响应格式的完整性
        if not isinstance(parsed_response, dict):
            print(f"⚠️ [LLM 响应格式错误] 返回类型不是 dict: {type(parsed_response)}")
            return {
                "status": "error",
                "message": "智能助手响应格式异常，请重试或切换 API 配置。",
                "refined_requirement": user_input,
                "questions": []
            }
        
        # 确保必要字段存在
        if "status" not in parsed_response:
            print(f"⚠️ [LLM 响应缺少 status 字段]")
            parsed_response["status"] = "need_clarification"
        
        if "refined_requirement" not in parsed_response:
            parsed_response["refined_requirement"] = user_input
        
        if "message" not in parsed_response:
            parsed_response["message"] = "请提供更多信息以便我更好地理解您的需求。"
        
        if "questions" not in parsed_response:
            parsed_response["questions"] = []
        
        # ⚠️ 关键检查：如果状态是 need_clarification 但没有问题，说明 LLM 出错了
        if parsed_response["status"] == "need_clarification" and not parsed_response["questions"]:
            print(f"❌ [LLM 逻辑错误] 状态为 need_clarification 但没有生成问题列表")
            # 自动修正为 ready 状态，避免死循环
            parsed_response["status"] = "ready"
            parsed_response["message"] = "已理解您的需求，正在准备操作计划..."
        
        print(f"✅ [LLM 精化完成] status={parsed_response['status']}, questions_count={len(parsed_response['questions'])}")
        
        return parsed_response
    
    def _parse_operation_plan(self, content: str) -> OperationPlan:
        """解析操作计划 JSON"""
        try:
            data = json.loads(content)
        except json.JSONDecodeError:
            # 尝试从文本中提取 JSON
            json_match = re.search(r'\{[\s\S]*\}', content)
            if json_match:
                data = json.loads(json_match.group())
            else:
                raise ValueError(f"无法解析 LLM 返回的操作计划: {content}")
        
        operations = []
        for op_data in data.get("operations", []):
            try:
                op_type = OperationType(op_data.get("type", "").upper())
            except ValueError:
                continue  # 跳过不支持的操作类型
            
            operations.append(Operation(
                type=op_type,
                params=op_data.get("params", {}),
                description=op_data.get("description", ""),
                target_sheet=op_data.get("target_sheet", "")
            ))
        
        return OperationPlan(
            operations=operations,
            summary=data.get("summary", ""),
            estimated_impact=data.get("estimated_impact", "")
        )
    
    def _parse_json_response(self, content: str) -> Dict[str, Any]:
        """解析 JSON 响应，更健壮的处理方式"""
        if not content:
            raise ValueError("LLM 返回内容为空")
        
        # 清理可能的 markdown 代码块标记
        content = content.strip()
        if content.startswith("```json"):
            content = content[7:]
        elif content.startswith("```"):
            content = content[3:]
        if content.endswith("```"):
            content = content[:-3]
        content = content.strip()
        
        # 尝试直接解析
        try:
            return json.loads(content)
        except json.JSONDecodeError:
            pass
        
        # 尝试提取 JSON 对象
        json_match = re.search(r'\{[\s\S]*\}', content)
        if json_match:
            try:
                return json.loads(json_match.group())
            except json.JSONDecodeError as e:
                raise ValueError(f"JSON 解析失败: {str(e)}\n原始内容: {content[:500]}")
        
        raise ValueError(f"无法从 LLM 返回中提取 JSON: {content[:500]}")
    
    def chat(
        self,
        messages: List[Dict[str, str]],
        system_prompt: Optional[str] = None
    ) -> str:
        """
        通用对话接口
        
        Args:
            messages: 对话消息列表
            system_prompt: 系统提示词
            
        Returns:
            str: LLM 回复
        """
        full_messages = []
        if system_prompt:
            full_messages.append({"role": "system", "content": system_prompt})
        full_messages.extend(messages)
        
        response = self.client.chat.completions.create(
            model=self.model,
            messages=full_messages,
            temperature=0.7
        )
        
        return response.choices[0].message.content
