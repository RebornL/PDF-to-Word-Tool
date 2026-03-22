# PDF转Word工具 - 敏感词替换

一款带图形界面的PDF转Word工具，支持敏感词/关键词搜索替换功能。

## 功能特性

- **PDF转Word**: 将PDF文件转换为Word文档(.docx)格式，保留原PDF排版格式
- **敏感词替换**: 在转换后的文档中搜索并替换指定关键词
- **批量替换**: 支持多组关键词批量替换
- **预览确认**: 替换前可预览效果，支持选择性替换
- **匹配选项**: 支持区分大小写、全词匹配
- **现代UI**: 使用CustomTkinter实现现代化界面
- **实时进度**: 转换过程中显示实时进度条

## 软件界面

![PDF转Word工具界面](screenshot.png)

## 使用方法

### 1. 转换PDF为Word

1. 点击 **浏览** 按钮选择要转换的PDF文件
2. 选择输出目录（默认与PDF同目录）
3. 点击 **转换PDF为Word** 按钮
4. 等待转换完成（进度条实时显示进度）

### 2. 搜索关键词

1. 在 **搜索词** 输入框中输入要查找的关键词
2. 可选择匹配选项：
   - **区分大小写**: 精确匹配大小写
   - **全词匹配**: 仅匹配完整单词
3. 点击 **搜索** 按钮
4. 搜索结果将显示在右侧表格中

### 3. 替换关键词

**替换单个关键词:**
1. 在 **替换为** 输入框中输入替换后的文本
2. 点击 **预览替换** 查看替换效果
3. 在表格中勾选要替换的项
4. 点击 **替换选中项** 或 **替换全部**

**批量替换多个关键词:**
1. 在 **批量替换列表** 中输入规则（每行一个）
   ```
   张三=***
   电话=联系方式
   身份证=证件号码
   ```
2. 点击 **批量替换**
3. 确认后执行替换

### 4. 保存文档

点击 **保存文档** 按钮，选择保存位置即可。

## 从源码运行

```bash
# 安装依赖
pip install -r requirements.txt

# 运行程序
python app.py
```

## 打包说明

### 使用PyInstaller打包

**推荐方式 (约77MB):**

```bash
pip install pyinstaller

pyinstaller --noconfirm --clean "PDF转Word工具.spec"
```

优化措施：
- 排除 OpenCV 视频IO模块 (PDF转换不需要)
- 排除 OpenCV 人脸检测模型
- 排除不需要的 Python 标准库模块
- 启用字节码优化

## 系统要求

- 操作系统: Windows 10/11 (64位)
- 无需安装Python环境（直接运行exe）

## 技术栈

| 组件 | 技术 |
|------|------|
| PDF引擎 | pdf2docx (保留格式) |
| UI框架 | CustomTkinter |
| Word处理 | python-docx |
| 打包工具 | PyInstaller |

## 项目结构

```
PdfToWordWithoutSensitive/
├── app.py                        # 主程序源码
├── src/                          # 模块化源码
│   ├── pdf_converter.py          # PDF转换模块
│   ├── search_replace.py         # 搜索替换模块
│   └── gui.py                    # GUI模块
├── PDF转Word工具.spec             # 打包配置文件
├── requirements.txt              # 依赖列表
├── README.md                     # 说明文档
└── dist/                         # 打包输出
    └── PDF转Word工具.exe          # 约77MB
```

## 注意事项

1. 使用pdf2docx引擎，能较好保留PDF排版格式（表格、图片、排版）
2. 建议在替换前先预览确认
3. 替换后请及时保存文档

## 许可证

MIT License