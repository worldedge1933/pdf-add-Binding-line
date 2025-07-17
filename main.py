import os
from PyPDF2 import PdfReader, PdfWriter, Transformation
import flet as ft


def shift_pdf_pages(
    input_path, output_path, shift_cm=1.0, start_page=1, end_page=None, first_right=True
):
    """
    对PDF文件的奇偶页进行平移处理，为胶装排版做准备。
    
    Args:
        input_path (str): 输入PDF文件的路径
        output_path (str): 输出PDF文件的路径
        shift_cm (float, optional): 平移距离，单位为厘米。默认为1.0厘米
        start_page (int, optional): 开始处理的页码（从1开始）。默认为1
        end_page (int, optional): 结束处理的页码，如果为None则处理到最后一页。默认为None
        first_right (bool, optional): 第一页是否向右平移。默认为True
    
    Returns:
        None: 处理后的PDF文件直接保存到指定路径
    """
    reader = PdfReader(input_path)
    writer = PdfWriter()
    shift = shift_cm * 28.35
    total_pages = len(reader.pages)
    if end_page is None or end_page > total_pages:
        end_page = total_pages
    for i, page in enumerate(reader.pages):
        page_num = i + 1
        if page_num < start_page or page_num > end_page:
            writer.add_page(page)
            continue
        # 判断方向
        if ((page_num - start_page) % 2 == 0 and first_right) or (
            (page_num - start_page) % 2 == 1 and not first_right
        ):
            t = Transformation().translate(tx=shift, ty=0)
        else:
            t = Transformation().translate(tx=-shift, ty=0)
        page.add_transformation(t)
        writer.add_page(page)
    with open(output_path, "wb") as f:
        writer.write(f)


def main(page: ft.Page):
    """
    主函数，创建PDF奇偶页平移工具的用户界面。
    
    Args:
        page (ft.Page): Flet页面对象，用于构建用户界面
    
    Returns:
        None: 函数在页面上添加UI组件并设置事件处理器
    """
    page.title = "PDF 奇偶页平移工具"
    page.window.width = 450
    page.window.height = 650

    page.theme = ft.Theme(font_family="Microsoft YaHei")

    input_file = ft.TextField(label="选择PDF文件", read_only=True, expand=True)
    output_file = ft.TextField(label="输出文件名", value="output.pdf", expand=True)
    status = ft.Text("", font_family="Microsoft YaHei")
    shift_value = ft.TextField(label="平移距离（厘米）", value="1", expand=True)

    # 新增参数输入框
    range_start = ft.TextField(label="开始页码（从1开始）", value="1", expand=True)
    range_end = ft.TextField(label="结束页码（留空为最后一页）", value="", expand=True)
    direction_switch = ft.Switch(label="第一页向右", value=True)

    def pick_file_result(e: ft.FilePickerResultEvent):
        """
        文件选择器的回调函数，处理用户选择PDF文件的事件。
        
        Args:
            e (ft.FilePickerResultEvent): 文件选择器事件对象
        
        Returns:
            None: 更新输入文件路径和输出文件名
        """
        if e.files:
            input_file.value = e.files[0].path
            # 自动设置输出文件名
            import os

            src_name = os.path.splitext(os.path.basename(e.files[0].path))[0]
            output_file.value = f"{src_name}(胶装排版).pdf"
            page.update()

    file_picker = ft.FilePicker(on_result=pick_file_result)
    page.overlay.append(file_picker)

    def process_pdf(e):
        """
        处理PDF文件的事件处理函数，执行实际的PDF平移操作。
        
        Args:
            e: 按钮点击事件对象
        
        Returns:
            None: 执行PDF处理并更新状态显示
        """
        if not input_file.value or not os.path.isfile(input_file.value):
            status.value = "请选择有效的PDF文件"
            page.update()
            return
        try:
            shift_cm = float(shift_value.value)
            start_page = int(range_start.value)
            end_page = int(range_end.value) if range_end.value.strip() else None
            first_right = direction_switch.value
        except Exception:
            status.value = "请输入有效参数"
            page.update()
            return
        status.value = "正在处理..."
        page.update()
        try:
            shift_pdf_pages(
                input_file.value,
                output_file.value,
                shift_cm,
                start_page,
                end_page,
                first_right,
            )
            status.value = f"处理完成"
        except Exception as ex:
            status.value = f"处理失败: {ex}"
        page.update()

    page.add(
        ft.Row(
            [
                input_file,
                ft.IconButton(
                    icon="folder_open",
                    on_click=lambda e: file_picker.pick_files(
                        allowed_extensions=["pdf"]
                    ),
                ),
            ]
        ),
        output_file,
        shift_value,
        ft.Row([range_start, range_end]),
        direction_switch,
        ft.ElevatedButton("开始处理", on_click=process_pdf),
        status,
    )


if __name__ == "__main__":
    ft.app(target=main)
