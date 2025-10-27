"""
文档样式定义模块
定义各种文档元素的样式配置
"""

from dataclasses import dataclass, asdict
from typing import Dict, Any, Optional
import json

@dataclass
class FontStyle:
    """字体样式配置"""
    family: str = "宋体"
    size: int = 12
    bold: bool = False
    italic: bool = False
    color: str = "#000000"

@dataclass
class ParagraphStyle:
    """段落样式配置"""
    alignment: str = "left"  # left, center, right, justify
    line_spacing: float = 1.5
    space_before: int = 0
    space_after: int = 0
    indent_first_line: int = 0
    indent_left: int = 0
    indent_right: int = 0

@dataclass
class PageStyle:
    """页面样式配置"""
    width: int = 210  # A4宽度 mm
    height: int = 297  # A4高度 mm
    margin_top: int = 25
    margin_bottom: int = 25
    margin_left: int = 25
    margin_right: int = 25
    orientation: str = "portrait"  # portrait, landscape

@dataclass
class HeaderFooterStyle:
    """页眉页脚样式配置"""
    enabled: bool = False
    content: str = ""
    font: FontStyle = None
    alignment: str = "center"
    distance_from_edge: int = 12

    def __post_init__(self):
        if self.font is None:
            self.font = FontStyle(size=10)

@dataclass
class TableStyle:
    """表格样式配置"""
    border_width: int = 1
    border_color: str = "#000000"
    cell_padding: int = 5
    header_background: str = "#f0f0f0"
    alternate_row_color: str = "#f9f9f9"
    font: FontStyle = None

    def __post_init__(self):
        if self.font is None:
            self.font = FontStyle()

class DocumentStyleManager:
    """文档样式管理器"""
    
    def __init__(self):
        # 默认样式配置
        self.styles = {
            "title": {
                "font": FontStyle(family="宋体", size=16, bold=True),
                "paragraph": ParagraphStyle(alignment="center", space_after=18)
            },
            "heading1": {
                "font": FontStyle(family="宋体", size=14, bold=True),
                "paragraph": ParagraphStyle(space_before=12, space_after=6)
            },
            "heading2": {
                "font": FontStyle(family="宋体", size=13, bold=True),
                "paragraph": ParagraphStyle(space_before=10, space_after=5)
            },
            "heading3": {
                "font": FontStyle(family="宋体", size=12, bold=True),
                "paragraph": ParagraphStyle(space_before=8, space_after=4)
            },
            "normal": {
                "font": FontStyle(family="宋体", size=12),
                "paragraph": ParagraphStyle(indent_first_line=24, line_spacing=1.5)
            },
            "code": {
                "font": FontStyle(family="Consolas", size=10),
                "paragraph": ParagraphStyle(indent_left=15, space_before=6, space_after=6)
            }
        }
        
        # 页面样式
        self.page_style = PageStyle()
        
        # 页眉页脚样式
        self.header_style = HeaderFooterStyle()
        self.footer_style = HeaderFooterStyle()
        
        # 表格样式
        self.table_style = TableStyle()

    def get_style(self, style_name: str) -> Dict[str, Any]:
        """获取指定样式"""
        return self.styles.get(style_name, self.styles["normal"])

    def get_all_styles(self) -> Dict[str, Any]:
        """获取所有样式配置"""
        return {
            "styles": self.styles,
            "page_style": self.page_style,
            "header_style": self.header_style,
            "footer_style": self.footer_style,
            "table_style": self.table_style
        }

    def update_style(self, style_name: str, font_style: FontStyle = None, paragraph_style: ParagraphStyle = None):
        """更新样式"""
        if style_name not in self.styles:
            self.styles[style_name] = {"font": FontStyle(), "paragraph": ParagraphStyle()}
        
        if font_style:
            self.styles[style_name]["font"] = font_style
        if paragraph_style:
            self.styles[style_name]["paragraph"] = paragraph_style

    def set_page_style(self, page_style: PageStyle):
        """设置页面样式"""
        self.page_style = page_style

    def set_header_footer(self, header: HeaderFooterStyle = None, footer: HeaderFooterStyle = None):
        """设置页眉页脚"""
        if header:
            self.header_style = header
        if footer:
            self.footer_style = footer

    def set_table_style(self, table_style: TableStyle):
        """设置表格样式"""
        self.table_style = table_style

    def export_config(self) -> str:
        """导出配置为JSON"""
        config = {
            "styles": {
                name: {
                    "font": asdict(style["font"]),
                    "paragraph": asdict(style["paragraph"])
                }
                for name, style in self.styles.items()
            },
            "page_style": asdict(self.page_style),
            "header_style": asdict(self.header_style),
            "footer_style": asdict(self.footer_style),
            "table_style": asdict(self.table_style)
        }
        return json.dumps(config, ensure_ascii=False, indent=2)

    def import_config(self, config_json: str):
        """从JSON导入配置"""
        try:
            config = json.loads(config_json)
            
            # 导入样式
            if "styles" in config:
                for name, style_config in config["styles"].items():
                    font_config = style_config.get("font", {})
                    paragraph_config = style_config.get("paragraph", {})
                    
                    self.styles[name] = {
                        "font": FontStyle(**font_config),
                        "paragraph": ParagraphStyle(**paragraph_config)
                    }
            
            # 导入页面样式
            if "page_style" in config:
                self.page_style = PageStyle(**config["page_style"])
            
            # 导入页眉页脚样式
            if "header_style" in config:
                header_config = config["header_style"]
                if "font" in header_config:
                    header_config["font"] = FontStyle(**header_config["font"])
                self.header_style = HeaderFooterStyle(**header_config)
            
            if "footer_style" in config:
                footer_config = config["footer_style"]
                if "font" in footer_config:
                    footer_config["font"] = FontStyle(**footer_config["font"])
                self.footer_style = HeaderFooterStyle(**footer_config)
            
            # 导入表格样式
            if "table_style" in config:
                table_config = config["table_style"]
                if "font" in table_config:
                    table_config["font"] = FontStyle(**table_config["font"])
                self.table_style = TableStyle(**table_config)
                
        except Exception as e:
            raise ValueError(f"配置导入失败: {str(e)}")

# 预定义样式模板
STYLE_TEMPLATES = {
    "学术论文": {
        "title": {"font": FontStyle(family="宋体", size=16, bold=True), "paragraph": ParagraphStyle(alignment="center", space_after=18)},
        "heading1": {"font": FontStyle(family="宋体", size=14, bold=True), "paragraph": ParagraphStyle(space_before=12, space_after=6)},
        "normal": {"font": FontStyle(family="宋体", size=12), "paragraph": ParagraphStyle(indent_first_line=24, line_spacing=1.5)},
        "page": PageStyle(margin_top=30, margin_bottom=25, margin_left=30, margin_right=25)
    },
    "商务报告": {
        "title": {"font": FontStyle(family="微软雅黑", size=18, bold=True), "paragraph": ParagraphStyle(alignment="center", space_after=15)},
        "heading1": {"font": FontStyle(family="微软雅黑", size=14, bold=True), "paragraph": ParagraphStyle(space_before=10, space_after=5)},
        "normal": {"font": FontStyle(family="微软雅黑", size=11), "paragraph": ParagraphStyle(line_spacing=1.3)},
        "page": PageStyle(margin_top=25, margin_bottom=25, margin_left=25, margin_right=25)
    },
    "技术文档": {
        "title": {"font": FontStyle(family="Arial", size=16, bold=True), "paragraph": ParagraphStyle(alignment="left", space_after=12)},
        "heading1": {"font": FontStyle(family="Arial", size=14, bold=True), "paragraph": ParagraphStyle(space_before=8, space_after=4)},
        "normal": {"font": FontStyle(family="Arial", size=10), "paragraph": ParagraphStyle(line_spacing=1.2)},
        "code": {"font": FontStyle(family="Consolas", size=9), "paragraph": ParagraphStyle(indent_left=15)},
        "page": PageStyle(margin_top=20, margin_bottom=20, margin_left=20, margin_right=20)
    }
}