# Markdown 与 Word 互相转换工具

一个轻量级纯前端的 Markdown 与 Word（.docx）互相转换工具，无需后端服务器，可直接在浏览器中使用。

预览这个网页[https://hqjprime.github.io/markdown-docx-change/]

## 功能特性

- 📄 **Markdown 转 Word**：支持通过文件上传或手动输入 Markdown 内容，转换为 .docx 格式
- 📝 **Word 转 Markdown**：支持上传 .docx 文件，转换为 Markdown 格式
- 🎨 **用户友好界面**：简洁直观的操作界面，适合各种用户
- 🔄 **实时转换进度**：显示转换过程中的进度百分比
- 📱 **响应式设计**：适配不同屏幕尺寸，支持移动端
- 📦 **纯前端实现**：无需后端服务器，所有转换在浏览器中完成
- 🚀 **快速转换**：使用专业库实现高效转换

## 支持的格式

### Markdown 转 Word 支持

- 6 级标题（# 到 ######）
- 无序列表（- 和 *）
- 有序列表（1. 2. 3.）
- 引用块（>）
- 分隔线（--- 和 ***）
- 普通段落
- 适当的间距和格式

### Word 转 Markdown 支持

- 标题级别
- 列表（有序和无序）
- 段落
- 引用
- 分隔线
- 基本文本格式

## 技术实现

- **前端框架**：纯原生 JavaScript
- **UI 设计**：HTML5 + CSS3
- **Markdown 解析**：[marked.js](https://marked.js.org/)
- **Word 文档生成**：[docx.js](https://docx.js.org/)
- **Word 文档解析**：[mammoth.js](https://github.com/mwilliamson/mammoth.js)
- **文件下载**：[FileSaver.js](https://github.com/eligrey/FileSaver.js/)

## 快速开始

### 方法一：直接打开

1. 克隆或下载本项目到本地
2. 直接在浏览器中打开 `index.html` 文件
3. 开始使用转换功能

### 方法二：本地服务器

1. 克隆或下载本项目到本地
2. 在项目目录下启动本地服务器
   
   ```bash
   # 使用 Python 3
   python -m http.server 3000
   
   # 或使用 Node.js
   npx http-server -p 3000
   
   # 或使用 PHP
   php -S localhost:3000
   ```
3. 在浏览器中访问 `http://localhost:3000`
4. 开始使用转换功能

## 使用指南

### Markdown 转 Word

1. 在转换类型中选择 "Markdown 转 Word"
2. 选择输入方式：
   - **上传文件**：点击 "上传文件" 按钮，选择 .md 文件
   - **手动输入**：点击 "手动输入" 按钮，在文本框中输入 Markdown 内容
3. 点击 "开始转换" 按钮
4. 转换完成后，点击 "下载结果" 按钮保存 .docx 文件

### Word 转 Markdown

1. 在转换类型中选择 "Word 转 Markdown"
2. 选择输入方式：
   - **上传文件**：点击 "上传文件" 按钮，选择 .docx 文件
   - **手动输入**：点击 "手动输入" 按钮，在文本框中输入 Word 内容
3. 点击 "开始转换" 按钮
4. 转换完成后，可以：
   - 查看转换结果预览
   - 点击 "下载结果" 按钮保存 .md 文件

## 项目结构

```
├── index.html          # 主页面
├── style.css           # 样式文件
├── script.js           # 核心功能实现
├── README.md           # 项目说明文档
└── test.md             # 测试用 Markdown 文件
```

## 浏览器兼容性

- Chrome 60+
- Firefox 55+
- Safari 12+
- Edge 79+

## 注意事项

1. 由于是纯前端实现，转换大型文件可能会受到浏览器性能限制
2. 对于复杂的 Word 文档，某些高级格式可能无法完全转换
3. 建议使用现代浏览器以获得最佳体验
4. 本工具仅用于个人或小型项目，不适合处理敏感或机密文档

## 本地开发

### 克隆项目

```bash
git clone <repository-url>
cd <project-directory>
```

### 启动开发服务器

```bash
# 使用 Python 3
python -m http.server 3000

# 或使用 Node.js
npx http-server -p 3000
```

### 访问应用

在浏览器中打开 `http://localhost:3000`

## 贡献

欢迎提交 Issue 和 Pull Request 来帮助改进这个项目！

## 许可证

本项目采用 MIT 许可证 - 详见 [LICENSE](LICENSE) 文件

## 致谢

- [marked.js](https://marked.js.org/) - Markdown 解析库
- [docx.js](https://docx.js.org/) - Word 文档生成库
- [mammoth.js](https://github.com/mwilliamson/mammoth.js) - Word 文档解析库
- [FileSaver.js](https://github.com/eligrey/FileSaver.js/) - 文件下载库

## 联系

如果您有任何问题或建议，欢迎通过以下方式联系：

- GitHub Issues：在本项目仓库中提交 Issue

---

**享受 Markdown 与 Word 之间的无缝转换！** ✨
