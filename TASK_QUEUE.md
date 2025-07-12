# MCP Sheet Parser - 任务队列

## 🚨 紧急状况说明

**发现重大问题**: MCP工具系统实现错误！

- **错误现状**: 当前实现了`get_sheet_info`、`parse_sheet`、`apply_changes`三个工具
- **正确设计**: 应该实现`convert_to_html`、`parse_sheet`、`apply_changes`三个工具
- **设计文档**: 参考`memory-bank/systemPatterns.md`中的"4.1 三个核心工具设计"
- **影响范围**: 整个MCP工具系统需要重新实现

## 📋 待完成任务队列

### 任务1: 修正MCP工具定义 (高优先级)
**文件**: `src/mcp_server/tools.py`

**当前问题**:
- 工具名称错误：`get_sheet_info` → 应该是 `convert_to_html`
- 工具描述不符合设计文档
- 参数定义需要调整

**需要修改**:
```python
# 错误的当前实现
Tool(name="get_sheet_info", ...)

# 正确的目标实现  
Tool(name="convert_to_html", ...)
```

**具体要求**:
1. 将`get_sheet_info`改为`convert_to_html`
2. 更新工具描述为"完美HTML转换，95%样式保真度"
3. 参数调整：
   - `file_path` (必需): 源文件路径
   - `output_path` (可选): 输出HTML路径
4. 更新工具处理函数：`_handle_get_sheet_info` → `_handle_convert_to_html`

### 任务2: 创建HTMLConverter组件 (高优先级)
**文件**: `src/converters/html_converter.py`

**状态**: 已开始创建但未完成

**需要实现**:
1. `HTMLConverter`类
2. `convert_to_file(sheet, output_path)`方法
3. 样式CSS生成逻辑
4. HTML模板系统
5. 95%样式保真度支持

**技术要求**:
- 支持15个样式属性（字体、颜色、背景、边框、对齐等）
- CSS类复用优化
- 合并单元格处理
- 超链接转换

### 任务3: 完善CoreService (高优先级)
**文件**: `src/core_service.py`

**当前问题**:
- 缺少`convert_to_html`方法
- `apply_changes`方法功能不完整
- 需要集成HTMLConverter

**需要添加**:
```python
def convert_to_html(self, file_path: str, output_path: str = None) -> Dict[str, Any]:
    """将表格文件转换为HTML文件"""
    # 1. 使用ParserFactory解析文件
    # 2. 调用HTMLConverter转换
    # 3. 返回转换结果
```

**需要完善**:
```python
def apply_changes(self, file_path: str, table_model_json: Dict, create_backup: bool = True) -> Dict[str, Any]:
    """将修改写回原文件"""
    # 1. 数据验证
    # 2. 备份原文件
    # 3. 写回数据（复杂功能，可先提供验证）
```

### 任务4: 更新工具处理函数 (中优先级)
**文件**: `src/mcp_server/tools.py`

**需要修改**:
1. `_handle_convert_to_html`函数实现
2. `_handle_apply_changes`函数更新（支持create_backup参数）
3. 错误处理和日志记录

### 任务5: 创建工具测试 (中优先级)
**文件**: `tests/test_mcp_tools.py`

**需要测试**:
1. `convert_to_html`工具的完整流程
2. `parse_sheet`工具的JSON输出
3. `apply_changes`工具的数据验证
4. 错误处理场景

### 任务6: 集成测试和验证 (中优先级)

**需要验证**:
1. 三个工具的端到端流程
2. 与现有解析器系统的集成
3. 样式保真度是否达到95%目标
4. 性能表现

## 📊 项目现状总结

### ✅ 已完成 (优秀成果)
- **解析器系统**: 5种格式解析器全部实现，平均97.64%保真度
- **测试覆盖**: 40个测试用例，100%通过率
- **样式验证**: StyleValidator系统完整实现
- **文档更新**: README和API文档已更新

### ⚠️ 发现问题
- **MCP工具错误**: 实现了错误的工具组合
- **架构不一致**: tools.py与systemPatterns.md设计不符
- **功能缺失**: HTMLConverter组件未完成

### 🔄 影响评估
- **解析器系统**: 无影响，可正常使用
- **MCP接口**: 需要重新实现，影响用户使用
- **整体进度**: 需要额外1-2天完成修正

## 🎯 交接建议

### 优先级排序
1. **P0 (紧急)**: 修正tools.py工具定义
2. **P0 (紧急)**: 完成HTMLConverter实现
3. **P1 (重要)**: 完善CoreService方法
4. **P2 (一般)**: 创建测试和验证

### 技术要点
- 严格按照`memory-bank/systemPatterns.md`设计实现
- 复用现有的解析器系统（ParserFactory等）
- 保持95%样式保真度目标
- 确保与现有测试系统兼容

### 验收标准
- [ ] 三个MCP工具正确实现并可调用
- [ ] convert_to_html能生成高质量HTML文件
- [ ] parse_sheet返回LLM友好的JSON格式
- [ ] apply_changes提供数据验证功能
- [ ] 所有测试通过
- [ ] 文档更新完整

## 📞 联系信息
- **设计文档**: `memory-bank/systemPatterns.md`
- **进度记录**: `memory-bank/progress.md`
- **当前状态**: `memory-bank/activeContext.md`
- **技术架构**: 基于新解析器系统，ParserFactory + StyleValidator

---

**注意**: 解析器系统本身已经完美实现，问题仅在于MCP工具层的接口错误。修正后整个系统将完整可用。
