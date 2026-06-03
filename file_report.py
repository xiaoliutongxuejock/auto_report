
#==================配置区==================
#1.文件路径
import hashlib
import shutil
from PIL import Image
import time
import win32com.client
from datetime import datetime
from pathlib import Path
import sys
import logging
import requests
import base64


base_path = Path(r'S:\04.FA\03.CIM\09.个人文件\刘欢\auto_reporter') #根目录
source_file = base_path / "FileServer空间增长记录" / "fileserver_temp_20260531.csv" #源文件
target_file = base_path / "效能表,温湿度表,fileserver表" / "fileserver.xlsx" #待处理的目标文件

title_rows = [18, 19]   #需要导出的标题行，可以根据实际情况调整，如果没有标题行可以设置为 None 或空列表

#日志路径
log_path = base_path / "logs"
log_path.mkdir(exist_ok=True)

#2.脚本路径

#3.截图配置  提前创建输出根目录，避免后续路径报错
output_root = base_path / 'screenshots'
output_root.mkdir(parents=True, exist_ok=True)
output_path = output_root / "report_file.png"  #图片输出目录
#测试
# output_root = Path(__file__).parent / 'screenshots'
# output_root.mkdir(exist_ok=True)
output_path = output_root / "report_file.png"  #图片输出目录
#4.企业微信配置
WEBHOOK_URL = "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=22823db5-535b-427c-b61f-502d5aa9e35c"
IMAGE_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp'}

#5.日志配置
log_file_name = log_path / f'report_{datetime.now().strftime("%Y-%m-%d_%H-%M-%S")}.log'
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file_name, encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

#=================配置区结束==================


#读取源文件最后一行数据，导入到目标文件最后一行
def process_excels(excel):
    #检查文件是否存在
    if not source_file.exists():
        logger.error(f"源文件 {source_file} 不存在，请检查路径是否正确。")
        sys.exit()
    if not target_file.exists():
        logger.error(f"目标文件{target_file}不存在，请检查路径是否正确。")
        sys.exit()

    #获取source_excel最后一行数据
    wb_source = None
    try:
        wb_source = excel.Workbooks.Open(source_file)
        ws_source = wb_source.ActiveSheet

        # 只读取实际使用的列数（而非整行），避免取到大量尾部None
        last_row_num = ws_source.Cells(ws_source.Rows.Count, 1).End(-4162).Row  # xlUp
        last_col_num = ws_source.Cells(last_row_num, ws_source.Columns.Count).End(-4159).Column  # xlToLeft

        source_range = ws_source.Range(
            ws_source.Cells(last_row_num, 1),
            ws_source.Cells(last_row_num, last_col_num)
        )
        row_data = source_range.Value
        clean_data = [v for v in row_data[0] if v is not None]
        logger.info(
            f"成功读取{source_file}, 行号：{last_row_num}, 列数：{last_col_num}, "
            f"有效数据：{clean_data}"
        )
    except Exception as e:
        logger.error(f"[ERROR]获取 Excel 最后一行失败: {e}")
        return None
    finally:
        if wb_source:
            wb_source.Close(SaveChanges=False)

    #导入到目标文件
    wb_target = None
    success = False
    try:
        wb_target = excel.Workbooks.Open(target_file)
        ws_target = wb_target.ActiveSheet
        last_row_target = ws_target.Cells(ws_target.Rows.Count, 1).End(-4162).Row  # xlUp
        new_row = last_row_target + 1
        logger.info(f"目标文件最后一行：{last_row_target}，准备写入第{new_row}行（共{last_col_num}列）")

        # 一次性写入整行数据（1次COM调用）
        dest_range = ws_target.Range(
            ws_target.Cells(new_row, 1),
            ws_target.Cells(new_row, last_col_num)
        )
        dest_range.Value = row_data[0]

        # 复制上一行的格式到新行
        ws_target.Rows(last_row_target).Copy()
        ws_target.Rows(new_row).PasteSpecial(Paste=-4122)  # xlPasteFormats
        excel.CutCopyMode = False  # 清除剪贴板，避免状态冲突

        success = True
        logger.info(f"成功将数据导入到{target_file}，行号：{new_row}")
    except Exception as e:
        logger.error(f"[ERROR]导入 Excel 数据失败: {e}")
    finally:
        if wb_target:
            if success:
                wb_target.Save()
            wb_target.Close()

#截图函数（Excel COM 直接导出，不依赖屏幕截图，RDP 最小化也能用）
#将指定单元格区域导出为裁剪后的 PNG，返回 (路径, 宽, 高)
def capture_range_as_png(ws,wb,start_row,end_row,end_col_letter,suffix,temp_dir):
    """将指定单元格区域导出为裁剪后的 PNG，返回 (路径, 宽, 高)"""
    rng = ws.Range(f"A{start_row}:{end_col_letter}{end_row}")
    # Chart 尺寸匹配 Range 实际宽高，避免画布留白造成图片间有间隙
    rng_width = rng.Width
    rng_height = rng.Height
    tmp_ws = wb.Worksheets.Add()
    chart_obj = tmp_ws.ChartObjects().Add(0, 0, rng_width, rng_height)
    chat = chart_obj.Chart
    # CopyPicture 必须在 Add 工作表之后、Paste 之前调用，否则切换工作表会清空剪贴板
    rng.CopyPicture(Appearance=1, Format=2)
    chat.Paste()
    data_png_path = temp_dir / f"capture_{suffix}.png"
    chat.Export(str(data_png_path), FilterName="PNG")
    time.sleep(10)  # 增加短暂延迟，确保文件完全写入
    chart_obj.Delete()
    tmp_ws.Delete()

    img =Image.open(data_png_path)
    gray_img = img.convert('L')
    bbox = gray_img.point(lambda x: 0 if x > 250 else 255).getbbox()
    if bbox:
        img = img.crop(bbox)
        img.save(data_png_path, "PNG")
    dw, dh = img.size
    img.close()
    return data_png_path, dw, dh
#导出工作表：冻结区域的图表 + 标题行图 + 最后7行数据图
def export_excel_view(excel,output_path,title_rows=None):
    
    try:
        logger.info(f"处理: {target_file}，准备导出 Excel 视图为图片...")
        wb = excel.Workbooks.Open(target_file)
        ws = wb.ActiveSheet
        # 滚到底部，确保最后几行数据在可见范围内（xlScreen模式需要）
        last_row_num= ws.Cells(ws.Rows.Count, 1).End(-4162).Row
        ws.Range(f"A{max(1, last_row_num - 30)}").Select()
        time.sleep(1)  # 等待滚动完成
        logger.info(f"滚动至 A{max(1, last_row_num - 30)} (最后数据行: {last_row_num})，确保数据在可见范围内。")

        time.sleep(1.5)  # 等待 Excel 稳定，避免文件锁定或未完全加载导致的错误
        excel.ScreenUpdating = True # 允许界面更新，确保截图能捕捉到最新内容
        time.sleep(1.5)  # 确保界面更新完成

        temp_dir = output_path.parent/"_temp_parts"
        temp_dir.mkdir(exist_ok=True)

        charts = ws.ChartObjects()
        parts = []

        #1.导出所有图表（位于冻结区域）
        for chart in charts:
            # 导出图表为图片
            tp = temp_dir / f"chart_{chart.Name}.png"
            chart.Chart.Export(str(tp), FilterName="PNG")
            parts.append({
                "path": tp,
                "width": chart.Width,
                "height": chart.Height,
                "left": chart.Left,
                "top": chart.Top
            })

        #2.导出标题行（18、19行）
        title_start = title_end = None
        scan_ranges = []
        if title_rows:
            title_start = min(title_rows)
            title_end = max(title_rows)
            scan_ranges = [range(title_start, title_end + 1)]
        #3.导出最后7行数据
        data_start_row_num = max(1, last_row_num - 6)
        scan_ranges.append(range(data_start_row_num, last_row_num + 1))
        max_col = 1
        for col in range(1, 26):
            has_data = False
            for row_range in scan_ranges:
                for row in row_range:
                    try:
                        val = ws.Cells(row, col).Value
                    except Exception:
                        continue
                    if val is not None and val != "":
                        has_data = True
                        break
                if has_data:
                    break
            if has_data:
                max_col = col 
        end_col_letter = chr(64 + max_col) if max_col <= 26 else None  # 将列号转换为字母（1->A, 2->B,...）
        logger.info(f"检测到数据区域最大列为 {end_col_letter}，将导出 A{title_start if title_start else data_start_row_num} 到 {end_col_letter}{last_row_num} 的内容。") 
        #4.将数据转换成chrt对象截图（因为直接截图单元格区域在RDP最小化时会有问题，改为复制范围内容到Chart再导出图片的方式，稳定且不受显示环境影响）
        if end_col_letter:
            data_top = max((p["top"] + p["height"] for p in parts), default=0) + 30
            data_left = min((p["left"] for p in parts), default=0)       
            #(1).标题行（先抓，放在顶部）
            if title_start:
                png_path, dw, dh = capture_range_as_png(
                ws, wb, title_start, title_end, end_col_letter, "title", temp_dir)
                parts.append({"path": png_path, "left": data_left, "top": data_top, "width": dw, "height": dh})
                data_top += dh

            #(2).最后 7 行数据行（后抓，放在标题下方）
            png_path, dw, dh = capture_range_as_png(
            ws, wb, data_start_row_num, last_row_num, end_col_letter, "data", temp_dir)
            parts.append({"path": png_path, "left": data_left, "top": data_top, "width": dw, "height": dh})
           
            #(3).拼接
            if not parts:
                logger.warning("没有内容可导出")
                return
            max_width = int(max(p["left"] + p["width"] for p in parts))
            max_height = int(max(p["top"] + p["height"] for p in parts))
            canvas = Image.new("RGB", (max_width + 20, max_height + 20), "white")
            for p in parts:
                try:
                    img = Image.open(p["path"])
                    w, h = int(p["width"]), int(p["height"])
                    if img.size != (w, h):
                        img = img.resize((w, h), Image.LANCZOS)
                    canvas.paste(img, (int(p["left"]), int(p["top"])))
                    img.close()
            
                except Exception as e:
                    logger.error(f"[ERROR]拼接处理图片 {p['path']} 时出错: {e}")
            canvas.save(output_path, "PNG")
            logger.info(f"导出完成: {output_path} ({canvas.size[0]}x{canvas.size[1]})")
    except Exception as e:
        logger.error(f"[ERROR]导出 Excel 视图失败: {e}")
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
        time.sleep(1)
        try:
            if 'wb' in locals():
                wb.Close(SaveChanges=False)
                logger.info(f"已关闭: {target_file}")
        except Exception:
            logger.warning(f"关闭 Excel 文件 {target_file} 时发生异常，可能已被关闭或未正确打开(不影响结果）")

#===========企业微信推送===============
def get_report_time():
    """获取报告时间字符串，格式为 YYYY-MM-DD HH:MM:SS"""
    return datetime.now().strftime("%Y-%m-%d "+"10:00:00")

def send_text_to_wechat(text,webhook_url):
    """发送文本消息到企业微信"""
    payload = {
        "msgtype":"text",
        "text":{
            "content":text
        }
    }
    try:
        response = requests.post(webhook_url,json = payload)
        result =response.json()
        if result.get('errcode') == 0:
            logger.info("文本消息发送成功")
        else:
            logger.error(f"文本消息发送失败：{result.get('errmsg')}")
    except Exception as e:
        logger.error(f"文本消息请求出错：{e}")

def send_image_to_wechat(image_path,webhook_url):
    """发送单张图片到企业微信"""
    success_count = 0
    if not image_path.exists():
        logger.error(f"图片文件 {image_path} 不存在，无法发送到企业微信。")
        return False
    try:
        with open(image_path,'rb') as f:
            image_data = f.read()
        
        if len(image_data) > 2 * 1024 * 1024:
            logger.error(f"图片文件 {image_path} 大小超过企业微信限制（2MB），无法发送。")
            return False
        
        base64_data = base64.b64encode(image_data).decode('utf-8')
        md5_hash = hashlib.md5(image_data).hexdigest()
        payload = {
            "msgtype": "image",
            "image": {
                "base64": base64_data,
                "md5": md5_hash
            }
        }

        response = requests.post(webhook_url, json=payload)
        result = response.json()

        if result.get('errcode') == 0:
            logger.info("图片消息发送成功")
            return True
        else:
            logger.error(f"图片消息发送失败：{result.get('errmsg')}")
            return False
    except Exception as e:
        logger.error(f"图片消息请求出错：{e}")
        return False
    
   
    


if __name__ == "__main__":
    #1.处理 Excel 数据并导出视图
    excel = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        #读取新增数据并导入目标文件
        process_excels(excel)
        #截图
        try:
            export_excel_view(excel,output_path, title_rows=title_rows)
        except Exception as e:
            logger.error(f"[ERROR]导出 Excel 视图失败: {e}")
    except Exception as e:
        logger.error(f"[ERROR]脚本运行异常: {e}", exc_info=True)
    finally:
        if excel:
            excel.Quit()
            del excel
            logger.info("Excel 应用已关闭。")
    #2.推送到企业微信
    report_time = get_report_time()
    report_text = f"呈长官：{report_time} Fileserver剩余空间及使用率如下"
    send_text_to_wechat(report_text, WEBHOOK_URL)
    send_image_to_wechat(output_path, WEBHOOK_URL)