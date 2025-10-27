# Document Converter Plugin for Dify

## 简介

这是一个用于 Dify 平台的文档转换插件，支持将 Markdown 文本转换为 Word 文档（.docx），并提供丰富的格式样式自定义功能。

## 功能特性

- **Markdown 转 Word**：将 Markdown 文本转换为格式化的 Word 文档
- **样式模板**：提供学术论文、商务报告、技术文档等预定义样式模板
- **自定义样式**：支持字体、段落、页面布局等全方位样式自定义
- **代码块支持**：保持代码块的格式和语法高亮样式
- **标题层级**：支持多级标题的自动格式化

## 安装方法

### 1. 本地安装插件

根据 [Dify 本地插件安装文档](https://blog.csdn.net/nalanxiaoxiao2011/article/details/147230916)，需要先修改 Dify 的配置文件：

1. 在 Dify 项目的 `.env` 文件中添加以下配置：
   ```
   FORCE_VERIFYING_SIGNATURE=false
   PLUGIN_MAX_PACKAGE_SIZE=524288000
   NGINX_CLIENT_MAX_BODY_SIZE=500M
   ```

2. 重启 Dify 服务

3. 将本插件目录打包为 `.zip` 文件

4. 在 Dify 管理界面上传并安装插件

### 2. 依赖安装

插件依赖以下 Python 包：
```
python-docx>=0.8.11
markdown>=3.4.0
beautifulsoup4>=4.11.0
dify-plugin
```

## 使用方法

### 在 Dify 工作流中使用

1. 在工作流中添加"Document Converter"工具
2. 配置以下参数：
   - **Markdown Content**：要转换的 Markdown 文本
   - **Style Template**：选择样式模板（学术论文/商务报告/技术文档）
   - **Custom Styles**：可选的自定义样式配置（JSON 格式）

### 样式模板

#### 学术论文
- 标题：宋体 16pt 加粗居中
- 一级标题：宋体 14pt 加粗
- 正文：宋体 12pt，首行缩进 24pt，行距 1.5

#### 商务报告
- 标题：微软雅黑 18pt 加粗居中
- 一级标题：微软雅黑 14pt 加粗
- 正文：微软雅黑 11pt，行距 1.3

#### 技术文档
- 标题：Arial 16pt 加粗左对齐
- 一级标题：Arial 14pt 加粗
- 正文：Arial 10pt，行距 1.2
- 代码：Consolas 9pt，左缩进 15pt

### 自定义样式示例

```json
{
  "title": {
    "font": {
      "family": "微软雅黑",
      "size": 20,
      "bold": true,
      "color": "#2c3e50"
    },
    "paragraph": {
      "alignment": "center",
      "space_after": 20
    }
  },
  "normal": {
    "font": {
      "family": "微软雅黑",
      "size": 12
    },
    "paragraph": {
      "line_spacing": 1.6,
      "indent_first_line": 32
    }
  }
}
```

## 支持的 Markdown 语法

- 标题（# ## ### 等）
- 段落文本
- 代码块（```）
- 基本文本格式

## 输出格式

- 文件格式：Microsoft Word (.docx)
- 兼容性：支持 Word 2007 及以上版本
- 编码：UTF-8

## 注意事项

1. 确保 Dify 环境已正确配置插件支持
2. 大文件转换可能需要较长时间
3. 自定义样式需要使用正确的 JSON 格式
4. 插件重启后可能需要重新配置

## 技术支持

如有问题或建议，请联系开发团队。