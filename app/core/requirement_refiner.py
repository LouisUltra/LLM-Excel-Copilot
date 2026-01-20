"""
需求精化模块
负责多轮对话精化用户的模糊需求
"""

import uuid
from typing import Dict, Any, Optional, List
from dataclasses import dataclass, field

from app.models import (
    ExcelMetadata, 
    RefineResponse, 
    ClarificationQuestion, 
    ClarificationOption,
    OperationPlan
)
from app.core.llm_client import LLMClient
from app.core.excel_parser import ExcelParser


@dataclass
class RefineSession:
    """需求精化会话"""
    session_id: str
    file_id: str
    metadata: ExcelMetadata
    file_description: str
    conversation_history: List[Dict[str, str]] = field(default_factory=list)
    refined_requirement: str = ""
    is_ready: bool = False
    operation_plan: Optional[OperationPlan] = None
    file_ids: List[str] = field(default_factory=list)  # 多文件ID列表


class RequirementRefiner:
    """
    需求精化器
    
    核心功能：
    1. 管理多轮对话会话
    2. 调用 LLM 分析用户模糊需求
    3. 生成澄清问题供用户确认
    4. 在用户确认后生成最终操作计划
    """
    
    def __init__(self, llm_client: Optional[LLMClient] = None):
        """
        初始化需求精化器
        
        Args:
            llm_client: LLM 客户端实例，不提供则创建新实例
        """
        self.llm_client = llm_client or LLMClient()
        # 会话存储（实际生产环境应使用持久化存储）
        self._sessions: Dict[str, RefineSession] = {}
    
    def create_session(
        self,
        file_id: str,
        metadata: ExcelMetadata,
        file_description: str,
        file_ids: List[str] = None
    ) -> str:
        """
        创建新的精化会话
        
        Args:
            file_id: 主文件 ID
            metadata: Excel 文件元数据
            file_description: 文件结构描述
            file_ids: 所有文件ID列表（多文件场景）
            
        Returns:
            str: 会话 ID
        """
        session_id = str(uuid.uuid4())
        session = RefineSession(
            session_id=session_id,
            file_id=file_id,
            metadata=metadata,
            file_description=file_description,
            file_ids=file_ids or [file_id]
        )
        self._sessions[session_id] = session
        return session_id
    
    def get_session(self, session_id: str) -> Optional[RefineSession]:
        """获取会话"""
        return self._sessions.get(session_id)
    
    def refine(
        self,
        session_id: str,
        user_input: str,
        answers: Optional[Dict[str, Any]] = None,
        previous_operations: Optional[Dict[str, Any]] = None
    ) -> RefineResponse:
        """
        精化用户需求
        
        Args:
            session_id: 会话 ID
            user_input: 用户输入
            answers: 用户对之前问题的回答
            previous_operations: 上一次执行的操作计划（继续编辑时的上下文）
            
        Returns:
            RefineResponse: 精化响应
        """
        session = self._sessions.get(session_id)
        if not session:
            return RefineResponse(
                session_id=session_id,
                status="error",
                message="会话不存在或已过期"
            )
        
        try:
            # 构建上下文信息（如果有上一次操作）
            context_info = ""
            if previous_operations:
                ops_desc = previous_operations.get("summary", "")
                ops_list = previous_operations.get("operations", [])
                if ops_list:
                    ops_details = "\n".join([f"  - {op.get('description', op.get('type', ''))}" for op in ops_list])
                    context_info = f"\n\n【上一次操作记录】\n{ops_desc}\n操作详情:\n{ops_details}\n\n用户现在可能是想基于上一次的操作结果继续修改。"
            
            # 调用 LLM 进行需求精化
            result = self.llm_client.refine_requirement(
                file_description=session.file_description + context_info,
                user_input=user_input,
                answers=answers,
                conversation_history=session.conversation_history
            )
            
            # 更新对话历史
            session.conversation_history.append({
                "role": "user",
                "content": user_input + (f"\n回答: {answers}" if answers else "")
            })
            session.conversation_history.append({
                "role": "assistant",
                "content": str(result)
            })
            
            # 解析 LLM 响应
            status = result.get("status", "need_clarification")
            refined_requirement = result.get("refined_requirement", "")
            session.refined_requirement = refined_requirement
            
            # 构建澄清问题
            questions = []
            for q_data in result.get("questions", []):
                options = [
                    ClarificationOption(
                        key=opt.get("key", ""),
                        label=opt.get("label", ""),
                        description=opt.get("description", "")
                    )
                    for opt in q_data.get("options", [])
                ]
                questions.append(ClarificationQuestion(
                    question_id=q_data.get("question_id", ""),
                    question=q_data.get("question", ""),
                    question_type=q_data.get("question_type", "single"),
                    options=options,
                    required=q_data.get("required", True)
                ))
            
            # 如果需求已经清晰，生成操作计划
            operation_plan = None
            if status == "ready":
                session.is_ready = True
                operation_plan = self.llm_client.generate_operations(
                    file_description=session.file_description,
                    user_requirement=refined_requirement
                )
                
                # 验证操作计划的合理性
                validation_result = self._validate_operation_plan(operation_plan, session.metadata)
                if validation_result["has_warnings"]:
                    # 如果有警告，生成二次确认问题
                    status = "need_clarification"
                    questions = [
                        ClarificationQuestion(
                            question_id="validation_warning",
                            question=f"⚠️ 检测到以下潜在问题：\n\n{validation_result['warning_message']}\n\n是否继续执行？",
                            question_type="single",
                            options=[
                                ClarificationOption(key="yes", label="是，继续执行", description=""),
                                ClarificationOption(key="no", label="否，重新调整", description="")
                            ],
                            required=True
                        )
                    ]
                else:
                    session.operation_plan = operation_plan
            
            return RefineResponse(
                session_id=session_id,
                status=status,
                refined_requirement=refined_requirement,
                questions=questions,
                operation_plan=operation_plan,
                message=result.get("message", "")
            )
            
        except Exception as e:
            return RefineResponse(
                session_id=session_id,
                status="error",
                message=f"处理请求时出错: {str(e)}"
            )
    
    def confirm_and_get_plan(self, session_id: str) -> Optional[OperationPlan]:
        """
        确认需求并获取操作计划
        
        Args:
            session_id: 会话 ID
            
        Returns:
            OperationPlan: 操作计划，如果会话不存在或未准备好则返回 None
        """
        session = self._sessions.get(session_id)
        if not session or not session.is_ready:
            return None
        
        # 如果还没有操作计划，现在生成
        if not session.operation_plan:
            session.operation_plan = self.llm_client.generate_operations(
                file_description=session.file_description,
                user_requirement=session.refined_requirement
            )
        
        return session.operation_plan
    
    def _validate_operation_plan(self, plan: OperationPlan, metadata: ExcelMetadata) -> Dict[str, Any]:
        """
        验证操作计划的合理性
        
        Returns:
            dict: {"has_warnings": bool, "warning_message": str, "warnings": list}
        """
        warnings = []
        
        # 获取所有可用的列名
        all_columns = set()
        column_types = {}  # 列名 -> 数据类型的映射
        for sheet in metadata.sheets:
            all_columns.update(sheet.headers)
            # 收集列的数据类型信息
            for col in sheet.columns:
                column_types[col.name] = col.data_type
        
        # 定义"所有列"的通配符表达（不区分大小写）
        WILDCARD_PATTERNS = [
            "所有列", "全部列", "每一列", "所有的列", "全部的列",
            "all", "all columns", "every column", "每列"
        ]
        
        def is_wildcard_column(col_name: str) -> bool:
            """检查是否是通配符表达"""
            if not col_name:
                return False
            col_lower = col_name.lower().strip()
            return any(pattern.lower() in col_lower for pattern in WILDCARD_PATTERNS)
        
        def suggest_expansion(col_name: str, context: str = "") -> str:
            """为通配符表达提供建议"""
            col_lower = col_name.lower().strip()
            
            # 检查是否包含类型限定词
            if "数值" in col_lower or "数字" in col_lower or "numeric" in col_lower:
                numeric_cols = [c for c, t in column_types.items() if t == "数字"]
                if numeric_cols:
                    return f"检测到 '{col_name}' 可能指所有数值列。建议使用具体列名：{numeric_cols[:5]}"
            elif "文本" in col_lower or "text" in col_lower:
                text_cols = [c for c, t in column_types.items() if t == "文本"]
                if text_cols:
                    return f"检测到 '{col_name}' 可能指所有文本列。建议使用具体列名：{text_cols[:5]}"
            else:
                # 纯"所有列"表达
                col_list = list(all_columns)[:5]
                more = f"等共{len(all_columns)}列" if len(all_columns) > 5 else ""
                return f"检测到 '{col_name}' 可能指表格中的所有列。建议在操作计划中明确列出具体列名：{col_list}{more}"
            
            return f"'{col_name}' 不是有效的列名"
        
        for op in plan.operations:
            # 验证1: 检查列名是否存在
            columns_to_check = []
            op_type = op.type.value
            
            # 收集需要验证的列名
            if op_type in ["FILTER", "SORT", "DELETE_COLUMN", "FORMAT", "REPLACE", "FILL"]:
                if "column" in op.params:
                    columns_to_check.append(op.params["column"])
                if "columns" in op.params:
                    columns_to_check.extend(op.params["columns"])
            elif op_type == "ADD_COLUMN":
                # 检查公式中引用的列是否存在（简单检查）
                formula = op.params.get("formula", "")
                if formula:
                    import re
                    # 提取列字母（如A、B、C）
                    col_refs = re.findall(r'([A-Z]+)\d+', formula)
                    if len(col_refs) > 26:  # 如果引用的列超过Z列，可能有问题
                        warnings.append(f"添加列操作：公式'{formula}'可能引用了过多列")
            elif op_type == "CREATE_CHART":
                data_cols = op.params.get("data_columns", [])
                label_col = op.params.get("label_column", "")
                columns_to_check.extend(data_cols)
                if label_col:
                    columns_to_check.append(label_col)
            elif op_type == "CALCULATE":
                ops = op.params.get("operations", [])
                for calc_op in ops:
                    if "column" in calc_op:
                        columns_to_check.append(calc_op["column"])
            elif op_type == "MERGE_COLUMNS":
                merge_cols = op.params.get("columns", [])
                columns_to_check.extend(merge_cols)
            elif op_type == "SPLIT_COLUMN":
                split_col = op.params.get("column", "")
                if split_col:
                    columns_to_check.append(split_col)
            
            # 🌟 智能检查列名
            for col in columns_to_check:
                if not col:
                    continue
                
                # 检查是否是通配符表达
                if is_wildcard_column(col):
                    suggestion = suggest_expansion(col, op_type)
                    warnings.append(f"⚠️ {suggestion}")
                # 检查列名是否存在
                elif col not in all_columns:
                    # 尝试模糊匹配
                    similar = [c for c in all_columns if col.lower() in c.lower() or c.lower() in col.lower()]
                    if similar:
                        warnings.append(f"列名 '{col}' 不存在，您可能是指：{similar[:3]}")
                    else:
                        warnings.append(f"列名 '{col}' 不存在于表格中")
            
            # 验证2: 检查危险操作
            if op_type == "DELETE_ROWS":
                warnings.append(f"将删除满足条件的行，此操作不可撤销")
            elif op_type == "DELETE_COLUMN":
                cols = op.params.get("columns", [])
                if len(cols) > 3:
                    warnings.append(f"将删除 {len(cols)} 列，请确认")
        
        # 构建警告消息
        warning_message = "\n".join([f"• {w}" for w in warnings])
        
        return {
            "has_warnings": len(warnings) > 0,
            "warning_message": warning_message,
            "warnings": warnings
        }
    
    def clear_session(self, session_id: str) -> bool:
        """
        清除会话
        
        Args:
            session_id: 会话 ID
            
        Returns:
            bool: 是否成功清除
        """
        if session_id in self._sessions:
            del self._sessions[session_id]
            return True
        return False
