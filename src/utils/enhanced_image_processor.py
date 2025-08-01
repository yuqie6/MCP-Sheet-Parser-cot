"""
图片处理增强功能

在重构分析中发现的问题修复和改进。
"""

import logging

logger = logging.getLogger(__name__)

# 图片处理常量
class ImageConstants:
    """图片处理相关常量"""
    DEFAULT_MAX_SIZE = 1024 * 1024  # 1MB
    MIN_DATA_SIZE = 10  # 最小有效图片数据大小

class EnhancedImageProcessor:
    """增强的图片处理器"""
    
    def __init__(self):
        # 图像格式签名和最小长度要求
        self.supported_formats = {
            # PNG: 需要完整的8字节签名
            b'\x89PNG\r\n\x1a\n': 'png',
            # JPEG: 需要完整的SOI标记
            b'\xff\xd8\xff': 'jpeg',
            # GIF: 需要完整的6字节签名
            b'GIF87a': 'gif',
            b'GIF89a': 'gif',
            # BMP: 需要完整的2字节签名
            b'BM': 'bmp',
            # TIFF: 需要完整的4字节签名
            b'II*\x00': 'tiff',
            b'MM\x00*': 'tiff',
            # ICO: 需要完整的4字节签名
            b'\x00\x00\x01\x00': 'ico',
            # WebP: 需要RIFF...WEBP签名（至少12字节）
            b'RIFF': 'webp'  # 特殊处理
        }
    
    def detect_image_format(self, img_data: bytes) -> str:
        """增强的图片格式检测，严格验证格式签名"""
        if not img_data:
            return 'unknown'

        # 特殊处理WebP格式
        if img_data.startswith(b'RIFF') and len(img_data) >= 12:
            if img_data[8:12] == b'WEBP':
                return 'webp'

        # 检查其他格式，要求完整的签名匹配
        for signature, format_name in self.supported_formats.items():
            if signature == b'RIFF':  # 跳过WebP，已经处理过了
                continue
            if len(img_data) >= len(signature) and img_data.startswith(signature):
                return format_name

        return 'unknown'
    
    def validate_image_data(self, img_data: bytes) -> bool:
        """验证图片数据的有效性"""
        if not img_data:
            return False

        # 检查是否是支持的格式
        format_name = self.detect_image_format(img_data)
        if format_name == 'unknown':
            logger.warning(f"Unsupported image format detected")
            return False

        # 根据格式检查最小长度要求
        min_lengths = {
            'png': 8,    # PNG签名长度
            'jpeg': 3,   # JPEG SOI标记长度
            'gif': 6,    # GIF签名长度
            'bmp': 2,    # BMP签名长度
            'tiff': 4,   # TIFF签名长度
            'ico': 4,    # ICO签名长度
            'webp': 12   # WebP需要RIFF...WEBP
        }

        min_length = min_lengths.get(format_name, ImageConstants.MIN_DATA_SIZE)
        if len(img_data) < min_length:
            logger.warning(f"Image data too short for {format_name} format: {len(img_data)} < {min_length}")
            return False

        return True
    
    def optimize_image_size(self, img_data: bytes, max_size: int = ImageConstants.DEFAULT_MAX_SIZE) -> bytes:
        """优化图片大小（如果需要）"""
        if len(img_data) <= max_size:
            return img_data
        
        # 这里可以实现图片压缩逻辑
        logger.warning(f"Image size {len(img_data)} exceeds limit {max_size}")
        return img_data
    
    def generate_image_html(self, img_data: bytes, alt_text: str = "Excel Image") -> str:
        """生成图片HTML"""
        if not self.validate_image_data(img_data):
            return self._generate_placeholder_html(alt_text)
        
        format_name = self.detect_image_format(img_data)
        
        # 优化图片大小
        optimized_data = self.optimize_image_size(img_data)
        
        # 生成Base64编码
        import base64
        img_base64 = base64.b64encode(optimized_data).decode('utf-8')
        data_url = f"data:image/{format_name};base64,{img_base64}"
        
        return f'''
        <div class="chart-svg-wrapper" style="background: none; padding: 0; min-height: auto;">
            <img src="{data_url}" alt="{alt_text}" style="max-width: 100%; height: auto; border-radius: 4px; background: transparent;" />
        </div>
        '''
    
    def _generate_placeholder_html(self, alt_text: str) -> str:
        """生成占位符HTML"""
        return f'''
        <div class="chart-svg-wrapper" style="background: #f5f5f5; padding: 20px; text-align: center; border: 1px dashed #ccc;">
            <p style="color: #666; font-style: italic;">图片加载失败: {alt_text}</p>
        </div>
        '''

# 使用示例
def demo_enhanced_image_processing():
    """演示增强的图片处理功能"""
    processor = EnhancedImageProcessor()
    
    # 测试数据
    import base64
    test_png = base64.b64decode(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChAI9fVNrFgAAAABJRU5ErkJggg=="
    )
    
    print("=== 增强图片处理功能演示 ===")
    print(f"1. 格式检测: {processor.detect_image_format(test_png)}")
    print(f"2. 数据验证: {processor.validate_image_data(test_png)}")
    print(f"3. 数据大小: {len(test_png)} bytes")
    
    # 生成HTML
    html = processor.generate_image_html(test_png, "测试图片")
    print(f"4. HTML生成: 成功 ({len(html)} 字符)")
    
    # 测试无效数据
    invalid_data = b"invalid image data"
    html_invalid = processor.generate_image_html(invalid_data, "无效图片")
    print(f"5. 无效数据处理: {'占位符' if '图片加载失败' in html_invalid else '异常'}")

if __name__ == "__main__":
    demo_enhanced_image_processing()