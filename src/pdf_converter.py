"""
PDF转Word转换模块
支持将PDF文件转换为Word文档格式(.docx)
"""

import os
from typing import Optional, Callable
from pathlib import Path

try:
    from pdf2docx import Converter
    HAS_PDF2DOCX = True
except ImportError:
    HAS_PDF2DOCX = False

try:
    import pdfplumber
    from docx import Document
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False


class PDFConverter:
    """PDF转Word转换器"""

    def __init__(self):
        self.cancelled = False

    def convert(
        self,
        pdf_path: str,
        output_path: str,
        progress_callback: Optional[Callable[[int, int], None]] = None
    ) -> bool:
        """
        将PDF转换为Word文档

        Args:
            pdf_path: PDF文件路径
            output_path: 输出Word文件路径
            progress_callback: 进度回调函数，参数为(当前页码, 总页数)

        Returns:
            bool: 转换是否成功
        """
        self.cancelled = False

        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"PDF文件不存在: {pdf_path}")

        # 确保输出目录存在
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)

        # 优先使用pdf2docx（保留格式更好）
        if HAS_PDF2DOCX:
            return self._convert_with_pdf2docx(pdf_path, output_path, progress_callback)
        elif HAS_PDFPLUMBER:
            return self._convert_with_pdfplumber(pdf_path, output_path, progress_callback)
        else:
            raise RuntimeError("没有可用的PDF转换库，请安装 pdf2docx 或 pdfplumber")

    def _convert_with_pdf2docx(
        self,
        pdf_path: str,
        output_path: str,
        progress_callback: Optional[Callable[[int, int], None]] = None
    ) -> bool:
        """使用pdf2docx进行转换（保留格式）"""
        try:
            cv = Converter(pdf_path)

            # 获取总页数
            total_pages = len(cv.pages)

            # 自定义进度回调
            def internal_progress(page_num):
                if progress_callback:
                    progress_callback(page_num, total_pages)
                if self.cancelled:
                    raise InterruptedError("转换已取消")

            cv.convert(output_path, progress=internal_progress)
            cv.close()
            return True

        except InterruptedError:
            return False
        except Exception as e:
            raise RuntimeError(f"PDF转换失败: {str(e)}")

    def _convert_with_pdfplumber(
        self,
        pdf_path: str,
        output_path: str,
        progress_callback: Optional[Callable[[int, int], None]] = None
    ) -> bool:
        """使用pdfplumber进行转换（纯文本模式）"""
        try:
            doc = Document()

            with pdfplumber.open(pdf_path) as pdf:
                total_pages = len(pdf.pages)

                for i, page in enumerate(pdf.pages):
                    if self.cancelled:
                        return False

                    if progress_callback:
                        progress_callback(i + 1, total_pages)

                    # 提取文本
                    text = page.extract_text()
                    if text:
                        # 按段落添加
                        for para in text.split('\n'):
                            if para.strip():
                                doc.add_paragraph(para)

                    # 每页之后添加分页符（除了最后一页）
                    if i < total_pages - 1:
                        doc.add_page_break()

            doc.save(output_path)
            return True

        except Exception as e:
            raise RuntimeError(f"PDF转换失败: {str(e)}")

    def cancel(self):
        """取消当前转换"""
        self.cancelled = True

    @staticmethod
    def get_page_count(pdf_path: str) -> int:
        """获取PDF页数"""
        if HAS_PDF2DOCX:
            try:
                cv = Converter(pdf_path)
                count = len(cv.pages)
                cv.close()
                return count
            except:
                pass

        if HAS_PDFPLUMBER:
            try:
                with pdfplumber.open(pdf_path) as pdf:
                    return len(pdf.pages)
            except:
                pass

        return 0


if __name__ == "__main__":
    # 测试代码
    converter = PDFConverter()

    def progress(current, total):
        print(f"转换进度: {current}/{total}")

    # 测试转换
    # converter.convert("test.pdf", "output.docx", progress)
