"""
智能字体管理器

自动处理各种字体，无需手动修改代码。
支持字体配置文件、自动检测、智能回退等功能。
"""

import json
import re
from pathlib import Path

import logging

logger = logging.getLogger(__name__)


class FontManager:
    """
    智能字体管理器
    
    功能：
    1. 自动检测字体类型（中文、英文、等宽、衬线等）
    2. 智能生成回退字体
    3. 支持配置文件自定义
    4. 自动学习新字体
    """
    
    def __init__(self, config_file: str | None = None):
        """
        初始化字体管理器

        参数：
            config_file: 字体配置文件路径
        """
        self.config_file = config_file or self._get_default_config_path()
        
        config_data = self._load_config_from_file()
        
        self.font_database = self._load_font_database(config_data)
        self.custom_mappings = self._load_custom_mappings(config_data)
        
    def _get_default_config_path(self) -> str:
        """获取默认配置文件路径"""
        config_dir = Path(__file__).parent / "config"
        config_dir.mkdir(exist_ok=True)
        return str(config_dir / "font_config.json")
    
    def _load_config_from_file(self) -> dict:
        """从文件加载配置"""
        try:
            if Path(self.config_file).exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except Exception as e:
            logger.warning(f"加载字体配置文件失败: {e}")
        return {}

    def _load_font_database(self, config_data: dict) -> dict:
        """加载字体数据库"""
        default_database = {
            "chinese_keywords": [
                # 中文字体关键词
                "华文", "微软", "宋体", "黑体", "楷体", "仿宋", "隶书", "行楷",
                "方正", "汉仪", "文泉驿", "思源", "Source Han", "Noto",
                "SimSun", "SimHei", "SimKai", "FangSong", "Microsoft YaHei",
                "PingFang", "Hiragino", "Yu Gothic", "Meiryo"
            ],
            "monospace_keywords": [
                # 等宽字体关键词
                "Courier", "Consolas", "Monaco", "Menlo", "DejaVu Sans Mono",
                "Liberation Mono", "Source Code Pro", "Fira Code", "JetBrains Mono",
                "Cascadia Code", "SF Mono", "Roboto Mono"
            ],
            "serif_keywords": [
                # 衬线字体关键词
                "Times", "Georgia", "Garamond", "Book Antiqua", "Palatino",
                "Baskerville", "Minion", "Caslon", "Trajan", "Optima"
            ],
            "sans_serif_keywords": [
                # 无衬线字体关键词
                "Arial", "Helvetica", "Verdana", "Tahoma", "Calibri",
                "Segoe UI", "Open Sans", "Roboto", "Lato", "Montserrat"
            ],
            "fallback_chains": {
                "chinese": [
                    "Microsoft YaHei", "PingFang SC", "Hiragino Sans GB",
                    "Source Han Sans SC", "Noto Sans CJK SC", "SimHei", "SimSun", "sans-serif"
                ],
                "monospace": [
                    "Consolas", "Monaco", "Courier New", "Liberation Mono",
                    "DejaVu Sans Mono", "monospace"
                ],
                "serif": [
                    "Times New Roman", "Georgia", "Times", "serif"
                ],
                "sans_serif": [
                    "Arial", "Helvetica", "Segoe UI", "sans-serif"
                ]
            }
        }
        
        # 合并默认配置和用户配置
        for key, value in config_data.get('font_database', {}).items():
            if key in default_database:
                if isinstance(value, list):
                    default_database[key].extend(value)
                elif isinstance(value, dict):
                    default_database[key].update(value)
        
        return default_database
    
    def _load_custom_mappings(self, config_data: dict) -> dict[str, str]:
        """加载自定义字体映射"""
        return config_data.get('custom_mappings', {})
    
    def detect_font_type(self, font_name: str) -> str:
        """
        自动检测字体类型
        
        参数:
            font_name: 字体名称
            
        返回:
            字体类型: 'chinese', 'monospace', 'serif', 'sans_serif'
        """
        if not font_name:
            return 'sans_serif'
        
        font_lower = font_name.lower()
        
        # 检查是否为中文字体
        for keyword in self.font_database['chinese_keywords']:
            if keyword.lower() in font_lower:
                return 'chinese'
        
        # 检查是否包含中文字符
        if any(ord(c) > 127 for c in font_name):
            return 'chinese'
        
        # 检查是否为等宽字体
        for keyword in self.font_database['monospace_keywords']:
            if keyword.lower() in font_lower:
                return 'monospace'
        
        # 检查是否为衬线字体
        for keyword in self.font_database['serif_keywords']:
            if keyword.lower() in font_lower:
                return 'serif'
        
        # 默认为无衬线字体
        return 'sans_serif'
    
    def needs_quotes(self, font_name: str) -> bool:
        """
        判断字体名称是否需要引号包围
        
        参数:
            font_name: 字体名称
            
        返回:
            是否需要引号
        """
        if not font_name:
            return False
        
        # 包含空格、中文字符或特殊字符时需要引号
        return (" " in font_name or 
                any(ord(c) > 127 for c in font_name) or
                any(c in font_name for c in [',', ';', '(', ')', '[', ']', '"', "'"]))
    
    def format_font_name(self, font_name: str) -> str:
        """
        格式化字体名称
        
        参数:
            font_name: 原始字体名称
            
        返回:
            格式化后的字体名称
        """
        if not font_name:
            return ""
        
        # 清理字体名称
        cleaned_name = font_name.strip()
        
        # 检查自定义映射
        if cleaned_name in self.custom_mappings:
            cleaned_name = self.custom_mappings[cleaned_name]
        
        # 添加引号（如果需要）
        if self.needs_quotes(cleaned_name):
            return f'"{cleaned_name}"'
        else:
            return cleaned_name
    
    def get_fallback_fonts(self, font_type: str) -> list[str]:
        """
        获取回退字体列表
        
        参数:
            font_type: 字体类型

        返回:
            回退字体列表
        """
        return self.font_database['fallback_chains'].get(font_type, 
                                                        self.font_database['fallback_chains']['sans_serif'])
    
    def generate_font_family(self, font_name: str) -> str:
        """
        生成完整的font-family字符串
        
        参数:
            font_name: 主字体名称

        返回:
            完整的font-family字符串
        """
        if not font_name:
            return ', '.join(self.get_fallback_fonts('sans_serif'))
        
        # 格式化主字体名称
        formatted_name = self.format_font_name(font_name)
        
        # 检测字体类型
        font_type = self.detect_font_type(font_name)
        
        # 获取回退字体
        fallback_fonts = self.get_fallback_fonts(font_type)
        
        # 组合字体列表
        font_list = [formatted_name] + fallback_fonts
        
        return ', '.join(font_list)
    
    def learn_font(self, font_name: str, font_type: str) -> None:
        """
        学习新字体（添加到数据库）
        
        参数:
            font_name: 字体名称
            font_type: 字体类型
        """
        if not font_name or font_type not in self.font_database['fallback_chains']:
            return
        
        # 提取字体关键词
        keywords = self._extract_keywords(font_name)
        
        # 添加到对应类型的关键词列表
        type_key = f"{font_type}_keywords"
        if type_key in self.font_database:
            for keyword in keywords:
                if keyword not in self.font_database[type_key]:
                    self.font_database[type_key].append(keyword)
                    logger.info(f"学习新字体关键词: {keyword} -> {font_type}")
    
    def _extract_keywords(self, font_name: str) -> list[str]:
        """
        从字体名称中提取关键词
        
        参数:
            font_name: 字体名称
            
        返回:
            关键词列表
        """
        keywords = []
        
        # 按空格分割
        words = font_name.split()
        for word in words:
            if len(word) > 1:  # 忽略单字符
                keywords.append(word)
        
        # 提取中文词汇
        chinese_pattern = r'[\u4e00-\u9fff]+'
        chinese_words = re.findall(chinese_pattern, font_name)
        keywords.extend(chinese_words)
        
        return keywords
    
    def save_config(self) -> None:
        """保存配置到文件"""
        try:
            config_data = {
                'font_database': self.font_database,
                'custom_mappings': self.custom_mappings
            }

            # 确保配置目录存在
            config_path = Path(self.config_file)
            config_path.parent.mkdir(parents=True, exist_ok=True)

            # 使用临时文件确保原子性写入
            temp_file = config_path.with_suffix('.tmp')
            try:
                with open(temp_file, 'w', encoding='utf-8') as f:
                    json.dump(config_data, f, ensure_ascii=False, indent=2)
                # 原子性地替换原文件
                temp_file.replace(config_path)
                logger.info(f"字体配置已保存到: {self.config_file}")
            except Exception as e:
                # 清理临时文件
                if temp_file.exists():
                    temp_file.unlink()
                raise e
        except (OSError, TypeError) as e:
            logger.error(f"保存字体配置失败: {e}")
        except Exception as e:
            logger.error(f"保存字体配置时发生未预期的错误: {e}")
    
    def add_custom_mapping(self, original_name: str, mapped_name: str) -> None:
        """
        添加自定义字体映射
        
        参数:
            original_name: 原始字体名称
            mapped_name: 映射后的字体名称
        """
        self.custom_mappings[original_name] = mapped_name
        logger.info(f"添加字体映射: {original_name} -> {mapped_name}")
    
    def get_font_info(self, font_name: str) -> dict:
        """
        获取字体的详细信息
        
        参数:
            font_name: 字体名称

        返回:
            字体信息字典
        """
        font_type = self.detect_font_type(font_name)
        formatted_name = self.format_font_name(font_name)
        font_family = self.generate_font_family(font_name)
        
        return {
            'original_name': font_name,
            'formatted_name': formatted_name,
            'font_type': font_type,
            'needs_quotes': self.needs_quotes(font_name),
            'font_family': font_family,
            'fallback_fonts': self.get_fallback_fonts(font_type)
        }


# 全局字体管理器实例（线程安全）
import threading
_font_manager = None
_font_manager_lock = threading.Lock()

def get_font_manager() -> FontManager:
    """获取全局字体管理器实例（线程安全）"""
    global _font_manager
    if _font_manager is None:
        with _font_manager_lock:
            # 双重检查锁定模式
            if _font_manager is None:
                _font_manager = FontManager()
    return _font_manager
