"""
API 配置管理模块
负责管理 LLM API 配置、获取模型列表、持久化保存
"""

import os
import re
from pathlib import Path
from typing import Dict, List, Optional, Any
from dataclasses import dataclass

from openai import OpenAI


@dataclass
class APIConfig:
    """API 配置"""
    id: str = ""  # 配置唯一标识
    name: str = ""  # 配置名称
    api_key: str = ""
    api_base: str = "https://api.openai.com/v1"
    model: str = "gpt-5.1"
    is_default: bool = False  # 是否为默认配置


class APIManager:
    """
    API 配置管理器
    
    功能：
    1. 管理多个 API 配置（增删改查）
    2. 获取 API 可用的模型列表
    3. 测试 API 连接
    4. 使用 JSON 文件持久化存储
    """
    
    def __init__(self, config_path: Optional[Path] = None):
        """
        初始化 API 管理器
        
        Args:
            config_path: 配置文件路径，默认为项目根目录的 api_configs.json
        """
        if config_path is None:
            # 默认在项目根目录
            self.config_path = Path(__file__).parent.parent.parent / "api_configs.json"
        else:
            self.config_path = Path(config_path)
        
        # 加载配置
        self._configs: Dict[str, APIConfig] = {}
        self._load_configs()
    
    def _load_configs(self):
        """从文件加载配置"""
        if self.config_path.exists():
            try:
                import json
                data = json.loads(self.config_path.read_text(encoding='utf-8'))
                for item in data.get('configs', []):
                    config = APIConfig(
                        id=item.get('id', ''),
                        name=item.get('name', ''),
                        api_key=item.get('api_key', ''),
                        api_base=item.get('api_base', 'https://api.openai.com/v1'),
                        model=item.get('model', 'gpt-5.1'),
                        is_default=item.get('is_default', False)
                    )
                    self._configs[config.id] = config
            except Exception as e:
                print(f"加载API配置失败: {e}")
                self._configs = {}
        
        # 如果没有任何配置，创建一个默认配置
        if not self._configs:
            default_id = 'default'
            self._configs[default_id] = APIConfig(
                id=default_id,
                name='默认配置',
                api_key='',
                api_base='https://api.openai.com/v1',
                model='gpt-5.1',
                is_default=True
            )
    
    def _save_configs(self):
        """保存配置到文件"""
        try:
            import json
            data = {
                'configs': [
                    {
                        'id': cfg.id,
                        'name': cfg.name,
                        'api_key': cfg.api_key,
                        'api_base': cfg.api_base,
                        'model': cfg.model,
                        'is_default': cfg.is_default
                    }
                    for cfg in self._configs.values()
                ]
            }
            self.config_path.write_text(
                json.dumps(data, ensure_ascii=False, indent=2),
                encoding='utf-8'
            )
            return True
        except Exception as e:
            print(f"保存API配置失败: {e}")
            return False
    
    def list_configs(self) -> List[Dict[str, Any]]:
        """
        获取所有配置列表（API Key 脱敏）
        
        Returns:
            List[Dict]: 配置列表
        """
        result = []
        for cfg in self._configs.values():
            # 隐藏 API Key 中间部分
            masked_key = ""
            if cfg.api_key:
                if len(cfg.api_key) > 8:
                    masked_key = cfg.api_key[:4] + "****" + cfg.api_key[-4:]
                else:
                    masked_key = "****"
            
            result.append({
                'id': cfg.id,
                'name': cfg.name,
                'api_key': masked_key,
                'api_key_set': bool(cfg.api_key),
                'api_base': cfg.api_base,
                'model': cfg.model,
                'is_default': cfg.is_default
            })
        return result
    
    def get_config(self, config_id: Optional[str] = None) -> Optional[APIConfig]:
        """
        获取指定配置，如果不指定则返回默认配置
        
        Args:
            config_id: 配置ID，不提供则返回默认配置
            
        Returns:
            APIConfig: API配置对象，不存在则返回None
        """
        if config_id:
            return self._configs.get(config_id)
        
        # 查找默认配置
        for cfg in self._configs.values():
            if cfg.is_default and cfg.api_key:
                return cfg
        
        # 如果没有默认配置，返回第一个有API Key的配置
        for cfg in self._configs.values():
            if cfg.api_key:
                return cfg
        
        # 如果都没有，返回第一个配置
        if self._configs:
            return list(self._configs.values())[0]
        
        return None
    
    def add_config(
        self, 
        name: str,
        api_key: str, 
        api_base: str, 
        model: str,
        set_as_default: bool = False
    ) -> Dict[str, Any]:
        """
        添加新的 API 配置
        
        Args:
            name: 配置名称
            api_key: API 密钥
            api_base: API 基础地址
            model: 模型名称
            set_as_default: 是否设为默认
            
        Returns:
            Dict: 包含 success 和新配置的 id
        """
        import uuid
        config_id = str(uuid.uuid4())
        
        # 如果设为默认，取消其他配置的默认状态
        if set_as_default:
            for cfg in self._configs.values():
                cfg.is_default = False
        
        new_config = APIConfig(
            id=config_id,
            name=name,
            api_key=api_key,
            api_base=api_base,
            model=model,
            is_default=set_as_default
        )
        
        self._configs[config_id] = new_config
        
        if self._save_configs():
            return {'success': True, 'id': config_id, 'message': '配置已添加'}
        else:
            return {'success': False, 'message': '保存配置失败'}
    
    def update_config(
        self,
        config_id: str,
        name: Optional[str] = None,
        api_key: Optional[str] = None,
        api_base: Optional[str] = None,
        model: Optional[str] = None,
        is_default: Optional[bool] = None
    ) -> Dict[str, Any]:
        """
        更新已有配置
        
        Args:
            config_id: 配置ID
            name: 配置名称
            api_key: API密钥
            api_base: API基础地址
            model: 模型名称
            is_default: 是否为默认
            
        Returns:
            Dict: 包含 success 和 message
        """
        if config_id not in self._configs:
            return {'success': False, 'message': '配置不存在'}
        
        config = self._configs[config_id]
        
        if name is not None:
            config.name = name
        if api_key is not None:
            config.api_key = api_key
        if api_base is not None:
            config.api_base = api_base
        if model is not None:
            config.model = model
        if is_default is not None:
            if is_default:
                # 取消其他配置的默认状态
                for cfg in self._configs.values():
                    cfg.is_default = False
            config.is_default = is_default
        
        if self._save_configs():
            return {'success': True, 'message': '配置已更新'}
        else:
            return {'success': False, 'message': '保存配置失败'}
    
    def delete_config(self, config_id: str) -> Dict[str, Any]:
        """
        删除配置
        
        Args:
            config_id: 配置ID
            
        Returns:
            Dict: 包含 success 和 message
        """
        if config_id not in self._configs:
            return {'success': False, 'message': '配置不存在'}
        
        # 如果删除的是默认配置，自动设置另一个为默认
        was_default = self._configs[config_id].is_default
        del self._configs[config_id]
        
        if was_default and self._configs:
            # 将第一个配置设为默认
            list(self._configs.values())[0].is_default = True
        
        if self._save_configs():
            return {'success': True, 'message': '配置已删除'}
        else:
            return {'success': False, 'message': '保存配置失败'}
    
    def set_default(self, config_id: str) -> Dict[str, Any]:
        """
        设置默认配置
        
        Args:
            config_id: 配置ID
            
        Returns:
            Dict: 包含 success 和 message
        """
        if config_id not in self._configs:
            return {'success': False, 'message': '配置不存在'}
        
        # 取消所有默认状态
        for cfg in self._configs.values():
            cfg.is_default = False
        
        # 设置新的默认
        self._configs[config_id].is_default = True
        
        if self._save_configs():
            return {'success': True, 'message': '默认配置已更新'}
        else:
            return {'success': False, 'message': '保存配置失败'}
    
    def get_models(self, api_key: str, api_base: str) -> List[Dict[str, Any]]:
        """
        获取 API 可用的模型列表
        
        Args:
            api_key: API 密钥
            api_base: API 基础地址
            
        Returns:
            List[Dict]: 模型列表，每个模型包含 id 和 name
        """
        try:
            client = OpenAI(api_key=api_key, base_url=api_base)
            models = client.models.list()
            
            # 过滤聊天模型
            chat_models = []
            for model in models.data:
                model_id = model.id
                # 过滤掉嵌入模型等非聊天模型
                if any(x in model_id.lower() for x in ['embed', 'whisper', 'tts', 'dall-e', 'moderation']):
                    continue
                chat_models.append({
                    'id': model_id,
                    'name': model_id
                })
            
            # 按名称排序
            chat_models.sort(key=lambda x: x['id'])
            return chat_models
            
        except Exception as e:
            print(f"获取模型列表失败: {e}")
            return []
    
    def test_connection(self, api_key: str, api_base: str, model: str) -> Dict[str, Any]:
        """
        测试 API 连接
        
        Args:
            api_key: API 密钥
            api_base: API 基础地址
            model: 模型名称
            
        Returns:
            Dict: 包含 success 和 message
        """
        try:
            client = OpenAI(api_key=api_key, base_url=api_base)
            
            # 发送一个简单的测试请求
            response = client.chat.completions.create(
                model=model,
                messages=[{"role": "user", "content": "Hi"}],
                max_tokens=5
            )
            
            return {
                'success': True,
                'message': f'连接成功！模型 {model} 可用。'
            }
            
        except Exception as e:
            error_msg = str(e)
            if 'authentication' in error_msg.lower() or 'api key' in error_msg.lower():
                return {'success': False, 'message': 'API Key 无效'}
            elif 'model' in error_msg.lower():
                return {'success': False, 'message': f'模型 {model} 不可用'}
            else:
                return {'success': False, 'message': f'连接失败: {error_msg[:100]}'}


# 全局实例
api_manager = APIManager()
