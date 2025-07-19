# 测试文档

本目录包含 MCP Sheet Parser 项目的所有单元测试。

## 测试结构

```
tests/
├── README.md                    # 本文档
├── test_core_service.py        # 核心服务测试 
├── test_validators.py          # 验证器测试
├── test_font_manager.py        # 字体管理器测试
├── cache/                      # 缓存模块测试
├── converters/                 # 转换器模块测试
├── mcp_server/                 # MCP服务器测试
├── models/                     # 数据模型测试
├── parsers/                    # 解析器测试
├── streaming/                  # 流式处理测试
└── utils/                      # 工具函数测试
```

## 运行测试

### 使用测试运行脚本 (推荐)

项目根目录提供了 `run_tests.py` 脚本，支持多种测试模式：

```bash
# 只测试核心服务 (快速验证)
python run_tests.py --mode core

# 运行所有测试
python run_tests.py --mode all

# 生成覆盖率报告
python run_tests.py --mode coverage

# 详细输出
python run_tests.py --mode core --verbose

# 自定义覆盖率目标
python run_tests.py --mode coverage --target-coverage 85
```

### 直接使用 uv 命令

```bash
# 测试核心服务 
uv run pytest tests/test_core_service.py --cov=src.core_service --cov-report=term-missing

# 运行所有测试
uv run pytest --cov=src --cov-report=term-missing

# 生成HTML覆盖率报告
uv run pytest --cov=src --cov-report=html

# 运行特定测试文件
uv run pytest tests/test_validators.py -v

# 运行特定测试方法
uv run pytest tests/test_core_service.py::TestCoreService::test_parse_sheet_normal -v
```

## 核心服务测试详情

`test_core_service.py` 是最重要的测试文件，包含59个测试用例，覆盖率达到87%：

### 主要测试功能

1. **解析功能测试**
   - 正常文件解析
   - 缓存机制测试
   - 工作表选择
   - 范围解析
   - 流式读取
   - 异常处理

2. **HTML转换测试**
   - 基本转换功能
   - 分页转换
   - 工作表选择
   - 错误处理

3. **数据写回测试**
   - XLSX格式写回
   - CSV格式写回
   - XLS格式写回
   - 备份功能
   - 异常处理

4. **内部方法测试**
   - 数据类型分析
   - 范围提取
   - JSON序列化
   - 大小计算
   - 样式处理

### 测试覆盖的边界情况

- 空文件处理
- 不存在的工作表
- 无效的范围字符串
- 文件权限问题
- 内存优化策略
- 流式处理回退

## 覆盖率目标

- **核心服务**: 87% (目标: 80%) ✅
- **整体项目**: 75% (目标: 70%) ✅

## 测试最佳实践

1. **命名规范**: 测试方法使用中文描述，清晰表达测试意图
2. **Mock使用**: 适当使用mock来隔离外部依赖
3. **边界测试**: 重点测试边界条件和异常情况
4. **数据驱动**: 使用pytest的fixture来提供测试数据
5. **断言清晰**: 使用具体的断言消息

## 添加新测试

当添加新功能时，请确保：

1. 在相应的测试文件中添加测试用例
2. 覆盖正常流程和异常情况
3. 使用描述性的测试方法名
4. 添加适当的注释说明测试目的
5. 运行测试确保通过
