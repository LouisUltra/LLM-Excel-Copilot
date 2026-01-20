"""
数据模型定义
使用 Pydantic 定义 API 请求和响应模型
"""

from typing import Optional, List, Dict, Any, Literal
from pydantic import BaseModel, Field
from enum import Enum


# ============ Excel 结构相关模型 ============

class ColumnInfo(BaseModel):
    """列信息"""
    name: str = Field(description="列名/表头")
    index: int = Field(description="列索引(从0开始)")
    data_type: str = Field(description="推断的数据类型")
    sample_values: List[str] = Field(default=[], description="示例值(脱敏后)")
    has_empty: bool = Field(default=False, description="是否存在空值")
    unique_count: Optional[int] = Field(default=None, description="唯一值数量")


class SheetInfo(BaseModel):
    """工作表信息"""
    name: str = Field(description="工作表名称")
    index: int = Field(description="工作表索引")
    total_rows: int = Field(description="总行数(不含表头)")
    total_cols: int = Field(description="总列数")
    headers: List[str] = Field(description="表头列表")
    columns: List[ColumnInfo] = Field(default=[], description="列详细信息")
    has_merged_cells: bool = Field(default=False, description="是否有合并单元格")
    has_formulas: bool = Field(default=False, description="是否包含公式")


class ExcelMetadata(BaseModel):
    """Excel 文件元数据"""
    file_id: str = Field(description="文件唯一标识")
    file_name: str = Field(description="原始文件名")
    file_size: int = Field(description="文件大小(字节)")
    sheets: List[SheetInfo] = Field(description="工作表信息列表")
    active_sheet: str = Field(description="默认活动工作表名称")


# ============ 操作指令相关模型 ============

class OperationType(str, Enum):
    """支持的操作类型"""
    FILTER = "FILTER"           # 筛选
    SORT = "SORT"               # 排序
    ADD_COLUMN = "ADD_COLUMN"   # 新增列
    DELETE_COLUMN = "DELETE_COLUMN"  # 删除列
    DELETE_ROWS = "DELETE_ROWS"      # 删除行
    DEDUPLICATE = "DEDUPLICATE"      # 去重
    CALCULATE = "CALCULATE"          # 计算汇总
    FORMAT = "FORMAT"                # 数字/日期格式化
    STYLE = "STYLE"                  # 样式设置（边框、背景色等）
    VLOOKUP = "VLOOKUP"              # 跨表查找
    PIVOT = "PIVOT"                  # 数据透视
    FILL = "FILL"                    # 填充数据
    REPLACE = "REPLACE"              # 替换
    SPLIT_COLUMN = "SPLIT_COLUMN"    # 拆分列
    MERGE_COLUMNS = "MERGE_COLUMNS"  # 合并列
    # 图表操作
    CREATE_CHART = "CREATE_CHART"    # 创建图表
    # 多文件合并操作
    MERGE_VERTICAL = "MERGE_VERTICAL"      # 纵向合并（追加行）
    MERGE_HORIZONTAL = "MERGE_HORIZONTAL"  # 横向合并（按关键列匹配）


class Operation(BaseModel):
    """单个操作指令"""
    type: OperationType = Field(description="操作类型")
    params: Dict[str, Any] = Field(default={}, description="操作参数")
    description: str = Field(default="", description="操作描述(用于展示)")
    target_sheet: str = Field(default="", description="目标工作表")


class OperationPlan(BaseModel):
    """操作计划(多个操作的组合)"""
    operations: List[Operation] = Field(description="操作列表")
    summary: str = Field(description="操作摘要")
    estimated_impact: str = Field(default="", description="预估影响")


# ============ 需求精化相关模型 ============

class ClarificationOption(BaseModel):
    """澄清选项"""
    key: str = Field(description="选项标识")
    label: str = Field(description="选项显示文本")
    description: str = Field(default="", description="选项说明")


class ClarificationQuestion(BaseModel):
    """澄清问题"""
    question_id: str = Field(description="问题ID")
    question: str = Field(description="问题内容")
    question_type: Literal["single", "multiple", "text"] = Field(
        default="single", description="问题类型: 单选/多选/文本"
    )
    options: List[ClarificationOption] = Field(default=[], description="选项列表")
    required: bool = Field(default=True, description="是否必答")


class RefineResponse(BaseModel):
    """需求精化响应"""
    session_id: str = Field(description="会话ID")
    status: Literal["need_clarification", "ready", "error"] = Field(
        description="状态: 需要澄清/准备就绪/错误"
    )
    refined_requirement: str = Field(default="", description="精化后的需求描述")
    questions: List[ClarificationQuestion] = Field(default=[], description="需要澄清的问题")
    operation_plan: Optional[OperationPlan] = Field(default=None, description="操作计划")
    message: str = Field(default="", description="状态消息")


# ============ API 请求/响应模型 ============

class UploadResponse(BaseModel):
    """文件上传响应"""
    success: bool
    file_id: str = Field(default="")
    metadata: Optional[ExcelMetadata] = None
    message: str = Field(default="")


class RefineRequest(BaseModel):
    """需求精化请求"""
    file_id: str = Field(description="主文件ID")
    file_ids: List[str] = Field(default=[], description="所有文件ID列表(多文件场景)")
    session_id: str = Field(default="", description="会话ID(续接对话时提供)")
    user_input: str = Field(description="用户输入")
    answers: Dict[str, Any] = Field(default={}, description="问题回答")
    previous_operations: Optional[Dict[str, Any]] = Field(default=None, description="上一次执行的操作计划(继续编辑时的上下文)")


class ProcessRequest(BaseModel):
    """处理执行请求"""
    file_id: str = Field(description="文件ID")
    session_id: str = Field(description="会话ID")
    confirmed: bool = Field(default=False, description="用户是否确认执行")


class ProcessResponse(BaseModel):
    """处理执行响应"""
    success: bool
    file_id: str = Field(default="", description="输出文件ID")
    download_url: str = Field(default="", description="下载链接")
    summary: str = Field(default="", description="处理摘要")
    message: str = Field(default="")
