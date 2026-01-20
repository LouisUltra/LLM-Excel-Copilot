"""
FastAPI 应用入口
"""

import os
import uuid
import shutil
from pathlib import Path
from typing import Dict, Optional

from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware

from app.config import settings
from app.models import (
    UploadResponse,
    RefineRequest,
    RefineResponse,
    ProcessRequest,
    ProcessResponse,
    ExcelMetadata
)
from app.core.excel_parser import ExcelParser
from app.core.llm_client import LLMClient
from app.core.requirement_refiner import RequirementRefiner
from app.core.excel_executor import ExcelExecutor
from app.core.api_manager import api_manager
from pydantic import BaseModel


# API 配置请求模型
class TestConnectionRequest(BaseModel):
    api_key: str
    api_base: str
    model: str


class GetModelsRequest(BaseModel):
    api_key: str
    api_base: str


# 创建 FastAPI 应用
app = FastAPI(
    title="Excel 智能处理助手",
    description="隐私安全的 Excel 自动化处理工具",
    version="1.0.0"
)

# CORS 配置
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 全局存储（实际生产环境应使用数据库/Redis）
file_storage: Dict[str, Dict] = {}  # file_id -> {path, metadata, description}
refiner: Optional[RequirementRefiner] = None


def get_refiner() -> RequirementRefiner:
    """获取需求精化器实例"""
    global refiner
    if refiner is None:
        try:
            # 从 api_manager 获取当前配置
            config = api_manager.get_config()
            if config and config.api_key:
                from app.core.llm_client import LLMClient
                llm_client = LLMClient(
                    api_key=config.api_key,
                    api_base=config.api_base,
                    model=config.model
                )
                refiner = RequirementRefiner(llm_client=llm_client)
            else:
                # 回退到环境变量配置
                refiner = RequirementRefiner()
        except ValueError as e:
            # API Key 未配置
            raise HTTPException(status_code=500, detail=str(e))
    return refiner


@app.on_event("startup")
async def startup_event():
    """应用启动时初始化"""
    # 确保目录存在
    settings.upload_dir.mkdir(exist_ok=True)
    settings.output_dir.mkdir(exist_ok=True)
    print(f"📁 上传目录: {settings.upload_dir}")
    print(f"📁 输出目录: {settings.output_dir}")
    print(f"🚀 Excel 智能助手已启动")


# ============ API 路由 ============

@app.post("/api/upload", response_model=UploadResponse)
async def upload_file(file: UploadFile = File(...)):
    """
    上传 Excel 文件
    
    - 支持 .xlsx 和 .xls 格式
    - 返回文件 ID 和解析的元数据
    """
    # 验证文件类型
    if not file.filename:
        raise HTTPException(status_code=400, detail="文件名不能为空")
    
    ext = Path(file.filename).suffix.lower()
    if ext not in [".xlsx", ".xls"]:
        raise HTTPException(status_code=400, detail="只支持 .xlsx 和 .xls 格式")
    
    # 生成文件 ID 并保存
    file_id = str(uuid.uuid4())
    save_path = settings.upload_dir / f"{file_id}{ext}"
    
    try:
        with open(save_path, "wb") as f:
            content = await file.read()
            f.write(content)
        
        # 解析文件
        parser = ExcelParser(save_path)
        metadata = parser.parse(file_id)
        # 覆盖文件名为原始上传文件名（而不是 UUID）
        metadata.file_name = file.filename
        description = parser.generate_description(metadata)
        
        # 存储文件信息
        file_storage[file_id] = {
            "path": str(save_path),
            "original_name": file.filename,
            "metadata": metadata,
            "description": description
        }
        
        return UploadResponse(
            success=True,
            file_id=file_id,
            metadata=metadata,
            message="文件上传成功"
        )
        
    except Exception as e:
        # 清理失败的上传
        if save_path.exists():
            save_path.unlink()
        raise HTTPException(status_code=500, detail=f"文件处理失败: {str(e)}")


@app.post("/api/refine", response_model=RefineResponse)
async def refine_requirement(request: RefineRequest):
    """
    精化用户需求
    
    - 首次调用传入 file_id 和 user_input
    - 后续调用传入 session_id、user_input 和 answers
    - 多文件场景传入 file_ids 列表
    """
    # 验证主文件存在
    if request.file_id not in file_storage:
        raise HTTPException(status_code=404, detail="文件不存在或已过期")
    
    file_info = file_storage[request.file_id]
    refiner_instance = get_refiner()
    
    # 收集多文件信息（如果有）
    all_file_ids = request.file_ids if request.file_ids else [request.file_id]
    all_files_info = []
    combined_description = ""
    
    for fid in all_file_ids:
        if fid in file_storage:
            info = file_storage[fid]
            all_files_info.append({
                "file_id": fid,
                "metadata": info["metadata"],
                "description": info["description"]
            })
    
    # 生成多文件描述
    if len(all_files_info) > 1:
        combined_description = f"## 多文件场景（共 {len(all_files_info)} 个文件）\n\n"
        for i, finfo in enumerate(all_files_info, 1):
            combined_description += f"### 文件 {i}: {finfo['metadata'].file_name}\n"
            combined_description += finfo["description"] + "\n\n"
    else:
        combined_description = file_info["description"]
    
    # 创建或获取会话
    if not request.session_id:
        session_id = refiner_instance.create_session(
            file_id=request.file_id,
            metadata=file_info["metadata"],
            file_description=combined_description,
            file_ids=all_file_ids  # 传递所有文件ID
        )
    else:
        session_id = request.session_id
        if not refiner_instance.get_session(session_id):
            raise HTTPException(status_code=404, detail="会话不存在或已过期")
    
    # 精化需求 - 传递上一次操作上下文
    response = refiner_instance.refine(
        session_id=session_id,
        user_input=request.user_input,
        answers=request.answers if request.answers else None,
        previous_operations=request.previous_operations
    )
    
    return response


@app.post("/api/process", response_model=ProcessResponse)
async def process_file(request: ProcessRequest, background_tasks: BackgroundTasks):
    """
    执行 Excel 处理
    
    - 需要 session_id 和确认标志
    - 返回处理后的文件下载链接
    """
    print(f"📝 开始处理请求: file_id={request.file_id}, session_id={request.session_id}")
    
    refiner_instance = get_refiner()
    session = refiner_instance.get_session(request.session_id)
    
    if not session:
        print(f"❌ 会话不存在: {request.session_id}")
        raise HTTPException(status_code=404, detail="会话不存在或已过期")
    
    if not request.confirmed:
        print(f"❌ 未确认执行")
        raise HTTPException(status_code=400, detail="请先确认执行操作")
    
    print(f"✓ 获取操作计划...")
    # 获取操作计划
    plan = refiner_instance.confirm_and_get_plan(request.session_id)
    if not plan:
        print(f"❌ 操作计划为空")
        raise HTTPException(status_code=400, detail="没有可执行的操作计划")
    
    print(f"✓ 操作计划包含 {len(plan.operations)} 个操作")
    
    file_info = file_storage.get(request.file_id)
    if not file_info:
        print(f"❌ 源文件不存在: {request.file_id}")
        raise HTTPException(status_code=404, detail="源文件不存在")
    
    try:
        print(f"🔧 开始执行操作...")
        
        # 为合并操作和跨文件查找操作注入实际文件路径
        # LLM 可能生成 file_index 引用或文件名，需要转换为实际文件路径
        if hasattr(session, 'file_ids') and len(session.file_ids) > 1:
            # 构建文件名到路径的映射
            filename_to_path = {}
            for fid in session.file_ids:
                if fid in file_storage:
                    info = file_storage[fid]
                    original_name = info.get('original_name', '')
                    filename_to_path[original_name] = info['path']
                    # 也尝试不带扩展名的匹配
                    name_without_ext = Path(original_name).stem
                    filename_to_path[name_without_ext] = info['path']
            
            for op in plan.operations:
                if op.type.value in ['MERGE_VERTICAL', 'MERGE_HORIZONTAL', 'VLOOKUP']:
                    # 情况1: 有 source_file_index，转换为实际路径
                    file_index = op.params.get('source_file_index')
                    if file_index is not None and isinstance(file_index, int):
                        if 0 <= file_index < len(session.file_ids):
                            source_fid = session.file_ids[file_index]
                            if source_fid in file_storage:
                                op.params['source_file'] = file_storage[source_fid]['path']
                                print(f"  注入源文件路径(via index): {file_storage[source_fid]['path']}")
                    
                    # 情况2: 有 source_file 但是是文件名而不是路径，尝试解析
                    elif 'source_file' in op.params:
                        source_file = op.params['source_file']
                        # 如果不是绝对路径且不是现有文件，尝试通过文件名查找
                        if not Path(source_file).is_absolute() and not Path(source_file).exists():
                            # 尝试直接匹配文件名
                            if source_file in filename_to_path:
                                op.params['source_file'] = filename_to_path[source_file]
                                print(f"  解析文件名 '{source_file}' -> {op.params['source_file']}")
                            else:
                                # 尝试模糊匹配（包含关系）
                                for fname, fpath in filename_to_path.items():
                                    if source_file in fname or fname in source_file:
                                        op.params['source_file'] = fpath
                                        print(f"  模糊匹配文件名 '{source_file}' -> {fpath}")
                                        break
                                else:
                                    # 都没匹配到，默认使用第二个文件
                                    if len(session.file_ids) > 1:
                                        second_fid = session.file_ids[1]
                                        if second_fid in file_storage:
                                            op.params['source_file'] = file_storage[second_fid]['path']
                                            print(f"  无法匹配 '{source_file}'，使用第二个文件: {file_storage[second_fid]['path']}")
                    
                    # 情况3: 完全没有 source_file，默认使用第二个文件
                    elif 'source_file' not in op.params:
                        second_fid = session.file_ids[1]
                        if second_fid in file_storage:
                            op.params['source_file'] = file_storage[second_fid]['path']
                            print(f"  默认使用第二个文件: {file_storage[second_fid]['path']}")
                            
                    # 特殊处理：如果 target_sheet 包含文件名前缀（如 "测试 2.xlsx!Sheet1"），去掉文件名部分
                    if op.type.value == 'VLOOKUP' and 'target_sheet' in op.params:
                        target_sheet = op.params['target_sheet']
                        if '!' in target_sheet:
                            # 提取工作表名（去掉文件名前缀）
                            op.params['target_sheet'] = target_sheet.split('!')[-1]
                            print(f"  修正 target_sheet: {target_sheet} -> {op.params['target_sheet']}")
        
        # 执行操作
        executor = ExcelExecutor(file_info["path"])
        output_path = executor.execute_plan(plan)
        
        print(f"✓ 操作执行完成，输出路径: {output_path}")
        
        # 打印操作日志
        for log in executor.get_log():
            print(f"  {log}")
        
        executor.close()
        
        # 生成输出文件 ID
        output_file_id = str(uuid.uuid4())
        
        # 确保下载文件名使用 .xlsx 扩展名（因为输出总是 .xlsx 格式）
        original_name = file_info['original_name']
        name_without_ext = Path(original_name).stem
        download_name = f"processed_{name_without_ext}.xlsx"
        
        output_info = {
            "path": output_path,
            "original_name": download_name
        }
        file_storage[output_file_id] = output_info
        
        print(f"✅ 处理完成！输出文件ID: {output_file_id}")
        
        # 清理会话
        background_tasks.add_task(refiner_instance.clear_session, request.session_id)
        
        return ProcessResponse(
            success=True,
            file_id=output_file_id,
            download_url=f"/api/download/{output_file_id}",
            summary=plan.summary,
            message="处理完成"
        )
        
    except Exception as e:
        print(f"❌ 处理失败: {str(e)}")
        import traceback
        print(traceback.format_exc())
        raise HTTPException(status_code=500, detail=f"处理失败: {str(e)}")


@app.get("/api/download/{file_id}")
async def download_file(file_id: str):
    """下载处理后的文件"""
    if file_id not in file_storage:
        raise HTTPException(status_code=404, detail="文件不存在")
    
    file_info = file_storage[file_id]
    file_path = Path(file_info["path"])
    
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="文件不存在")
    
    return FileResponse(
        path=file_path,
        filename=file_info.get("original_name", file_path.name),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.post("/api/continue/{file_id}", response_model=UploadResponse)
async def continue_processing(file_id: str):
    """
    继续处理已处理的文件
    
    - 将输出文件作为新的输入文件
    - 重新解析文件结构
    - 返回新的文件 ID 和元数据
    """
    # 获取输出文件信息
    if file_id not in file_storage:
        raise HTTPException(status_code=404, detail="文件不存在")
    
    file_info = file_storage[file_id]
    output_path = Path(file_info["path"])
    
    if not output_path.exists():
        raise HTTPException(status_code=404, detail="文件不存在")
    
    try:
        # 生成新的文件 ID
        new_file_id = str(uuid.uuid4())
        
        # 将输出文件复制到上传目录（作为新的输入文件）
        new_file_path = settings.upload_dir / f"{new_file_id}.xlsx"
        shutil.copy2(output_path, new_file_path)
        
        # 解析文件结构
        parser = ExcelParser(new_file_path)
        metadata = parser.parse(new_file_id)
        
        # 生成文件描述（供 LLM 理解）
        description = parser.generate_description(metadata)
        
        # 保存文件信息
        file_storage[new_file_id] = {
            "path": str(new_file_path),
            "original_name": file_info["original_name"],
            "metadata": metadata,
            "description": description  # 🆕 添加文件描述
        }
        
        return UploadResponse(
            success=True,
            file_id=new_file_id,
            metadata=metadata,
            message="继续处理准备完成"
        )
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"继续处理失败: {str(e)}")


@app.get("/api/file/{file_id}/metadata")
async def get_file_metadata(file_id: str):
    """获取文件元数据"""
    if file_id not in file_storage:
        raise HTTPException(status_code=404, detail="文件不存在")
    
    return file_storage[file_id]["metadata"]


@app.delete("/api/file/{file_id}")
async def delete_file(file_id: str):
    """删除文件"""
    if file_id not in file_storage:
        raise HTTPException(status_code=404, detail="文件不存在")
    
    file_info = file_storage[file_id]
    file_path = Path(file_info["path"])
    
    if file_path.exists():
        file_path.unlink()
    
    del file_storage[file_id]
    return {"success": True, "message": "文件已删除"}


# ============ 静态文件服务 ============

# 挂载静态文件目录
static_dir = Path(__file__).parent / "static"
if static_dir.exists():
    app.mount("/static", StaticFiles(directory=str(static_dir)), name="static")


@app.get("/", response_class=HTMLResponse)
async def index():
    """主页"""
    index_path = static_dir / "index.html"
    if index_path.exists():
        return HTMLResponse(content=index_path.read_text(encoding="utf-8"))
    return HTMLResponse(content="<h1>Excel 智能助手</h1><p>请配置前端页面</p>")


# ============ API 配置管理 ============

@app.get("/api/configs")
async def list_configs():
    """获取所有 API 配置列表"""
    configs = api_manager.list_configs()
    return {"success": True, "configs": configs}


@app.get("/api/configs/{config_id}")
async def get_config(config_id: str):
    """获取指定 API 配置"""
    config = api_manager.get_config(config_id)
    if config:
        # 隐藏 API Key 中间部分
        masked_key = ""
        if config.api_key:
            if len(config.api_key) > 8:
                masked_key = config.api_key[:4] + "****" + config.api_key[-4:]
            else:
                masked_key = "****"
        
        return {
            "success": True,
            "config": {
                "id": config.id,
                "name": config.name,
                "api_key": masked_key,
                "api_key_set": bool(config.api_key),
                "api_base": config.api_base,
                "model": config.model,
                "is_default": config.is_default
            }
        }
    else:
        raise HTTPException(status_code=404, detail="配置不存在")


class AddConfigRequest(BaseModel):
    name: str
    api_key: str
    api_base: str
    model: str
    set_as_default: bool = False


@app.post("/api/configs")
async def add_config(request: AddConfigRequest):
    """添加新的 API 配置"""
    global refiner
    
    result = api_manager.add_config(
        name=request.name,
        api_key=request.api_key,
        api_base=request.api_base,
        model=request.model,
        set_as_default=request.set_as_default
    )
    
    if result['success'] and request.set_as_default:
        # 重置 refiner 以使用新配置
        refiner = None
    
    return result


class UpdateConfigRequest(BaseModel):
    name: Optional[str] = None
    api_key: Optional[str] = None
    api_base: Optional[str] = None
    model: Optional[str] = None
    is_default: Optional[bool] = None


@app.put("/api/configs/{config_id}")
async def update_config(config_id: str, request: UpdateConfigRequest):
    """更新 API 配置"""
    global refiner
    
    result = api_manager.update_config(
        config_id=config_id,
        name=request.name,
        api_key=request.api_key,
        api_base=request.api_base,
        model=request.model,
        is_default=request.is_default
    )
    
    if result['success'] and request.is_default:
        # 重置 refiner 以使用新配置
        refiner = None
    
    return result


@app.delete("/api/configs/{config_id}")
async def delete_config(config_id: str):
    """删除 API 配置"""
    result = api_manager.delete_config(config_id)
    return result


@app.post("/api/configs/{config_id}/set-default")
async def set_default_config(config_id: str):
    """设置默认 API 配置"""
    global refiner
    
    result = api_manager.set_default(config_id)
    
    if result['success']:
        # 重置 refiner 以使用新配置
        refiner = None
    
    return result


@app.post("/api/models")
async def get_models(request: GetModelsRequest):
    """获取可用模型列表"""
    models = api_manager.get_models(
        api_key=request.api_key,
        api_base=request.api_base
    )
    
    if models:
        return {"success": True, "models": models}
    else:
        return {"success": False, "models": [], "message": "获取模型列表失败，请检查 API 配置"}


@app.post("/api/test-connection")
async def test_connection(request: TestConnectionRequest):
    """测试 API 连接"""
    result = api_manager.test_connection(
        api_key=request.api_key,
        api_base=request.api_base,
        model=request.model
    )
    return result


# ============ 健康检查 ============

@app.get("/health")
async def health_check():
    """健康检查"""
    return {
        "status": "healthy",
        "llm_configured": bool(settings.llm_api_key)
    }


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(
        "app.main:app",
        host=settings.host,
        port=settings.port,
        reload=settings.debug
    )
