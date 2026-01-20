"""
配置管理模块
使用 pydantic-settings 管理应用配置
"""

import os
from pathlib import Path
from pydantic_settings import BaseSettings
from pydantic import Field


class Settings(BaseSettings):
    """应用配置类"""
    
    # LLM API 配置
    llm_api_key: str = Field(default="", description="LLM API 密钥")
    llm_api_base: str = Field(
        default="https://api.openai.com/v1",
        description="LLM API 基础地址"
    )
    llm_model: str = Field(default="gpt-5.1", description="使用的模型名称")
    
    # 服务器配置
    host: str = Field(default="0.0.0.0", description="服务器监听地址")
    port: int = Field(default=8000, description="服务器端口")
    debug: bool = Field(default=True, description="调试模式")
    
    # 文件路径配置
    base_dir: Path = Field(
        default_factory=lambda: Path(__file__).parent.parent,
        description="项目根目录"
    )
    
    @property
    def upload_dir(self) -> Path:
        """上传文件目录"""
        path = self.base_dir / "uploads"
        path.mkdir(exist_ok=True)
        return path
    
    @property
    def output_dir(self) -> Path:
        """输出文件目录"""
        path = self.base_dir / "outputs"
        path.mkdir(exist_ok=True)
        return path
    
    class Config:
        env_file = ".env"
        env_file_encoding = "utf-8"
        extra = "ignore"


# 全局配置实例
settings = Settings()
