import cv2
import numpy as np
import pytesseract
import easyocr
import torch
from pdf2image import convert_from_path
from PIL import Image
import os
import gc
import argparse
import json
from tqdm import tqdm
import datetime
import logging
import time
import psutil
from scipy import stats
from sklearn.cluster import KMeans
import matplotlib.pyplot as plt

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class ColumnDetector:
    """高级分栏检测器，专为单栏/双栏文档优化"""

    def __init__(self, min_column_width_ratio=0.15):
        self.min_column_width_ratio = min_column_width_ratio
        self.debug_dir = None
        self.debug_enabled = False

    def enable_debug(self, output_dir):
        """启用调试模式，保存中间结果"""
        self.debug_dir = os.path.join(output_dir, "debug")
        os.makedirs(self.debug_dir, exist_ok=True)
        self.debug_enabled = True

    def save_debug_image(self, image, name, page_num):
        """保存调试图像"""
        if not self.debug_enabled:
            return
        debug_path = os.path.join(self.debug_dir, f"page_{page_num}_{name}.png")
        cv2.imwrite(debug_path, image)

    def detect_columns(self, image, page_num):
        """
        检测文档分栏结构
        返回: (column_count, split_points)
        """
        # 转换为OpenCV格式
        img = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
        height, width = img.shape[:2]

        # 方法1: 文本行中心点聚类分析
        column_count, split_point = self.detect_by_textline_clustering(img, page_num)
        if column_count == 2:
            return 2, [split_point]

        # 方法2: 垂直投影分析（增强版）
        column_count, split_points = self.detect_by_vertical_projection(img, page_num)
        if column_count == 2:
            return 2, split_points

        # 方法3: 轮廓分析（备用方法）
        column_count, split_points = self.detect_by_contours(img, page_num)
        if column_count == 2:
            return 2, split_points

        # 默认返回单栏
        return 1, []

    def detect_by_textline_clustering(self, img, page_num):
        """基于文本行中心点聚类的分栏检测"""
        # 预处理
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
        self.save_debug_image(binary, "01_binary", page_num)

        # 水平投影找文本行
        horizontal_proj = np.sum(binary, axis=1)
        threshold = np.max(horizontal_proj) * 0.1
        text_lines = []
        in_text = False
        start_y = 0

        for y, value in enumerate(horizontal_proj):
            if value > threshold and not in_text:
                in_text = True
                start_y = y
            elif value <= threshold and in_text:
                in_text = False
                end_y = y
                # 只考虑高度合理的行（排除噪点）
                if (end_y - start_y) > 5:
                    text_lines.append((start_y, end_y))

        # 提取每行文本的中心点
        centers = []
        for start_y, end_y in text_lines:
            # 提取当前行区域
            line_roi = binary[start_y:end_y, :]

            # 垂直投影找文本边界
            vertical_proj = np.sum(line_roi, axis=0)
            threshold = np.max(vertical_proj) * 0.1
            in_text = False
            start_x = 0

            for x, value in enumerate(vertical_proj):
                if value > threshold and not in_text:
                    in_text = True
                    start_x = x
                elif value <= threshold and in_text:
                    in_text = False
                    end_x = x
                    # 计算文本块中心
                    center_x = (start_x + end_x) // 2
                    centers.append(center_x)

        # 如果文本行太少，无法有效聚类
        if len(centers) < 10:
            return 1, 0

        # K-Means聚类分析
        centers_np = np.array(centers).reshape(-1, 1)
        kmeans = KMeans(n_clusters=2, random_state=0).fit(centers_np)
        labels = kmeans.labels_
        cluster_centers = kmeans.cluster_centers_.flatten()

        # 分析聚类结果
        cluster1_count = np.sum(labels == 0)
        cluster2_count = np.sum(labels == 1)
        min_count = min(cluster1_count, cluster2_count)
        total_count = len(centers)

        # 双栏条件：两个簇都有显著数量的点
        if min_count / total_count > 0.3 and abs(cluster_centers[0] - cluster_centers[1]) > img.shape[
            1] * self.min_column_width_ratio:
            # 取两个簇中心的中点作为分栏点
            split_point = int((cluster_centers[0] + cluster_centers[1]) / 2)
            return 2, split_point

        return 1, 0

    def detect_by_vertical_projection(self, img, page_num):
        """基于垂直投影的分栏检测（增强版）"""
        height, width = img.shape[:2]

        # 预处理
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
        self.save_debug_image(binary, "02_binary", page_num)

        # 垂直投影
        vertical_proj = np.sum(binary, axis=0) / 255
        self.save_debug_image(self.plot_projection(vertical_proj, width, height), "03_vertical_proj", page_num)

        # 平滑投影曲线
        kernel_size = max(5, int(width * 0.01))
        kernel = np.ones(kernel_size) / kernel_size
        smoothed_proj = np.convolve(vertical_proj, kernel, 'same')
        self.save_debug_image(self.plot_projection(smoothed_proj, width, height), "04_smoothed_proj", page_num)

        # 自适应阈值
        threshold = np.percentile(smoothed_proj, 30) * 1.5
        gap_mask = smoothed_proj < threshold

        # 寻找连续空白区域
        gap_starts = []
        gap_ends = []
        in_gap = False

        for i, is_gap in enumerate(gap_mask):
            if is_gap and not in_gap:
                in_gap = True
                gap_starts.append(i)
            elif not is_gap and in_gap:
                in_gap = False
                gap_ends.append(i)

        # 如果最后处于gap中，添加结束点
        if in_gap:
            gap_ends.append(len(gap_mask) - 1)

        # 计算空白区域宽度
        gap_widths = [end - start for start, end in zip(gap_starts, gap_ends)]

        # 如果没有空白区域，返回单栏
        if not gap_widths:
            return 1, []

        # 找到最宽的空白区域
        max_width_index = np.argmax(gap_widths)
        max_start = gap_starts[max_width_index]
        max_end = gap_ends[max_width_index]
        max_width = max_end - max_start

        # 检查是否满足双栏条件
        if max_width > width * self.min_column_width_ratio:
            # 空白区域中心作为分栏点
            split_point = (max_start + max_end) // 2

            # 确保分栏点在合理区域 (30%-70%)
            if width * 0.3 < split_point < width * 0.7:
                return 2, [split_point]

        return 1, []

    def detect_by_contours(self, img, page_num):
        """基于轮廓分析的分栏检测（备用方法）"""
        # 预处理
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)

        # 形态学操作连接文本
        kernel = np.ones((5, 5), np.uint8)
        dilated = cv2.dilate(binary, kernel, iterations=2)
        self.save_debug_image(dilated, "05_dilated", page_num)

        # 查找轮廓
        contours, _ = cv2.findContours(dilated, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

        # 分析轮廓位置
        x_positions = []
        for cnt in contours:
            x, y, w, h = cv2.boundingRect(cnt)
            if w > 20 and h > 20:  # 过滤小噪点
                center_x = x + w // 2
                x_positions.append(center_x)

        if not x_positions:
            return 1, []

        # 使用核密度估计检测分布峰值
        kde = stats.gaussian_kde(x_positions)
        x_grid = np.linspace(0, img.shape[1], 1000)
        density = kde(x_grid)
        peaks = np.where((density[1:-1] > density[:-2]) & (density[1:-1] > density[2:]))[0] + 1

        # 如果有两个显著峰值，则为双栏
        if len(peaks) >= 2:
            # 按密度值排序峰值
            sorted_peaks = sorted(peaks, key=lambda i: density[i], reverse=True)[:2]
            sorted_peaks.sort()

            # 取两个峰值的中点
            split_point = int((x_grid[sorted_peaks[0]] + x_grid[sorted_peaks[1]]) / 2)
            return 2, [split_point]

        return 1, []

    def plot_projection(self, projection, width, height):
        """可视化投影曲线"""
        img = np.zeros((height, width, 3), dtype=np.uint8)
        max_val = np.max(projection)
        if max_val == 0:
            return img

        # 归一化到图像高度
        normalized = (projection / max_val * height).astype(int)

        # 绘制投影曲线
        for x, h in enumerate(normalized):
            cv2.line(img, (x, height), (x, height - h), (0, 255, 0), 1)

        return img


class DocumentProcessor:
    def __init__(self, dpi=300, lang='chi_sim+eng', use_easyocr=True, gpu_mem_ratio=0.8,
                 crop_top_ratio=0.1, crop_bottom_ratio=0.1, crop_margin=0.05):
        self.dpi = dpi
        self.lang = lang
        self.layout_analysis = True
        self.use_easyocr = use_easyocr and torch.cuda.is_available()
        self.gpu_mem_ratio = gpu_mem_ratio
        self.crop_top_ratio = crop_top_ratio
        self.crop_bottom_ratio = crop_bottom_ratio
        self.crop_margin = crop_margin  # 内容检测的安全边界

        # 初始化OCR引擎和分栏检测器
        self.ocr_engine = None
        self.column_detector = ColumnDetector(min_column_width_ratio=0.15)

        if self.use_easyocr:
            self.init_ocr_engine()

    def init_ocr_engine(self):
        """初始化OCR引擎"""
        # 将语言代码转换为EasyOCR格式
        lang_map = {
            'chi_sim': 'ch_sim',
            'chi_tra': 'ch_tra',
            'eng': 'en',
            'jpn': 'ja',
            'kor': 'ko'
        }

        # 解析语言列表
        ocr_langs = []
        for lang_code in self.lang.split('+'):
            if lang_code in lang_map:
                ocr_langs.append(lang_map[lang_code])
            else:
                ocr_langs.append(lang_code)

        # 创建EasyOCR阅读器
        if torch.cuda.is_available():
            torch.cuda.set_per_process_memory_fraction(self.gpu_mem_ratio)
            self.ocr_engine = easyocr.Reader(
                ocr_langs,
                gpu=True,
                quantize=True,
                model_storage_directory='easyocr_models',
                download_enabled=True,
                recog_network='standard'
            )

    def preprocess_image(self, image):
        """图像预处理增强OCR精度"""
        img = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

        # 自适应阈值二值化
        binary = cv2.adaptiveThreshold(
            gray, 255,
            cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY, 11, 2
        )

        # 降噪
        denoised = cv2.fastNlMeansDenoising(binary, h=10)

        return Image.fromarray(denoised)

    def smart_crop_page(self, image):
        """
        智能裁剪页面，移除页头和页尾
        返回: 裁剪后的图像
        """
        # 转换为OpenCV格式
        img = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
        height, width = img.shape[:2]

        # 转换为灰度图
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

        # 二值化处理
        _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)

        # 水平投影分析
        horizontal_proj = np.sum(binary, axis=1) / 255

        # 计算裁剪边界
        top_bound = 0
        bottom_bound = height

        # 1. 确定顶部边界
        for y in range(height):
            if horizontal_proj[y] > width * 0.05:  # 有内容的行
                top_bound = max(0, int(y - height * self.crop_margin))
                break

        # 2. 确定底部边界
        for y in range(height - 1, -1, -1):
            if horizontal_proj[y] > width * 0.05:  # 有内容的行
                bottom_bound = min(height, int(y + height * self.crop_margin))
                break

        # 3. 应用固定比例裁剪作为后备
        fixed_top = int(height * self.crop_top_ratio)
        fixed_bottom = height - int(height * self.crop_bottom_ratio)

        # 使用智能检测边界，但确保不会裁剪过多
        top_bound = min(top_bound, fixed_top)
        bottom_bound = max(bottom_bound, fixed_bottom)

        # 确保边界合理
        if bottom_bound - top_bound < height * 0.5:
            # 如果裁剪过多，回退到固定比例
            top_bound = fixed_top
            bottom_bound = fixed_bottom

        # 执行裁剪
        cropped_img = img[top_bound:bottom_bound, :]
        return Image.fromarray(cropped_img)

    def split_image_columns(self, image, split_points):
        """根据分栏点拆分图像"""
        img = np.array(image)
        width = img.shape[1]
        sections = []

        # 添加起始点
        all_points = [0] + split_points + [width]

        for i in range(len(all_points) - 1):
            left = all_points[i]
            right = all_points[i + 1]
            section = img[:, left:right]
            sections.append(Image.fromarray(section))

        return sections

    def extract_text_with_easyocr(self, image):
        """使用EasyOCR提取文本"""
        try:
            # 转换为OpenCV格式
            img_np = np.array(image.convert('RGB'))

            # 执行OCR
            results = self.ocr_engine.readtext(
                img_np,
                detail=0,  # 只返回文本
                paragraph=True,  # 按段落分组
                batch_size=1  # 单张处理减少峰值显存
            )
            return "\n".join(results)
        except Exception as e:
            logger.error(f"EasyOCR处理失败: {e}, 回退到Tesseract")
            return self.extract_text_with_tesseract(image)

    def extract_text_with_tesseract(self, image):
        """使用Tesseract提取文本（CPU）"""
        # 预处理提升OCR精度
        processed_img = self.preprocess_image(image)

        # OCR配置
        config = r'--oem 3 --psm 6 -c preserve_interword_spaces=1'

        # 执行OCR
        text = pytesseract.image_to_string(
            processed_img,
            lang=self.lang,
            config=config
        )

        return text.strip()

    def extract_text(self, image):
        """根据配置选择OCR引擎"""
        if self.use_easyocr:
            return self.extract_text_with_easyocr(image)
        else:
            return self.extract_text_with_tesseract(image)

    def release_resources(self):
        """释放资源"""
        # 清理GPU缓存
        if torch.cuda.is_available():
            torch.cuda.empty_cache()

        # 强制垃圾回收
        gc.collect()
        time.sleep(0.1)

    def process_pdf_page(self, page_image, page_num):
        """处理单页PDF图像"""
        # 裁剪页面（移除页头和页尾）
        cropped_image = self.smart_crop_page(page_image)

        results = []

        if self.layout_analysis:
            # 检测分栏
            num_columns, split_points = self.column_detector.detect_columns(cropped_image, page_num)

            # 双栏处理
            if num_columns == 2:
                # 拆分图像
                column_images = self.split_image_columns(cropped_image, split_points)

                # 对各栏分别OCR
                for col_idx, col_img in enumerate(column_images):
                    text = self.extract_text(col_img)
                    results.append({
                        "page": page_num,
                        "column": col_idx + 1,
                        "text": text,
                        "split_point": split_points[0] if col_idx == 0 else None
                    })
                return results

        # 单栏处理
        text = self.extract_text(cropped_image)
        results.append({
            "page": page_num,
            "column": 0,  # 0表示整页
            "text": text,
            "split_point": None
        })

        return results

    def process_pdf(self, pdf_path, output_dir, enable_debug=False):
        """处理整个PDF文档"""
        start_time = datetime.datetime.now()
        logger.info(f"开始处理文档: {start_time.strftime('%Y-%m-%d %H:%M:%S')}")

        # 创建输出目录
        os.makedirs(output_dir, exist_ok=True)

        # 启用调试模式
        if enable_debug:
            self.column_detector.enable_debug(output_dir)

        # PDF转图像
        logger.info(f"正在转换PDF为图像 (DPI={self.dpi})...")
        images = convert_from_path(pdf_path, dpi=self.dpi)
        logger.info(f"已转换 {len(images)} 页图像")

        # 处理所有页面
        all_results = []
        for i, image in enumerate(tqdm(images, desc="Processing PDF")):
            page_start = time.time()

            # 处理当前页
            page_results = self.process_pdf_page(image, i + 1)
            all_results.extend(page_results)

        # 保存结果
        self.save_results(all_results, output_dir, start_time)
        self.release_resources()
        return all_results

    def save_results(self, results, output_dir, start_time):
        """保存处理结果 - 添加时间戳到文件名"""
        end_time = datetime.datetime.now()
        duration = end_time - start_time

        # 生成时间戳字符串 (格式: _HHMMSS)
        timestamp = end_time.strftime("_%H%M%S")

        # 创建元数据
        metadata = {
            "start_time": start_time.strftime('%Y-%m-%d %H:%M:%S'),
            "end_time": end_time.strftime('%Y-%m-%d %H:%M:%S'),
            "duration_seconds": duration.total_seconds(),
            "page_count": len(set(res['page'] for res in results)),
            "ocr_engine": "EasyOCR" if self.use_easyocr else "Tesseract",
            "column_stats": self.calculate_column_stats(results),
            "crop_top_ratio": self.crop_top_ratio,
            "crop_bottom_ratio": self.crop_bottom_ratio
        }

        # 保存为JSON - 文件名添加时间戳
        json_filename = f"ocr_results{timestamp}.json"
        json_path = os.path.join(output_dir, json_filename)
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump({
                "metadata": metadata,
                "pages": results
            }, f, ensure_ascii=False, indent=2)

        # 保存为文本文件 - 文件名添加时间戳
        txt_filename = f"full_text{timestamp}.txt"
        txt_path = os.path.join(output_dir, txt_filename)

        # 按页码分组文本
        page_texts = {}
        for res in results:
            page = res['page']
            if page not in page_texts:
                page_texts[page] = []
            page_texts[page].append(res['text'])

        with open(txt_path, "w", encoding="utf-8") as f:
            # 添加处理信息头
            f.write(f"OCR处理报告\n")
            f.write(f"开始时间: {metadata['start_time']}\n")
            f.write(f"结束时间: {metadata['end_time']}\n")
            f.write(f"处理时长: {duration} ({duration.total_seconds():.2f}秒)\n")
            f.write(f"页面数量: {metadata['page_count']}\n")
            f.write(f"OCR引擎: {metadata['ocr_engine']}\n")
            f.write(f"裁剪设置: 顶部={self.crop_top_ratio * 100}%, 底部={self.crop_bottom_ratio * 100}%\n")
            f.write(
                f"分栏统计: 单栏页={metadata['column_stats']['single']}, 双栏页={metadata['column_stats']['double']}\n")
            f.write("-" * 50 + "\n\n")

            # 添加OCR结果 - 每页一个文本块
            for page in sorted(page_texts.keys()):
                f.write(f"\n\n--- 第 {page} 页 ---\n\n")
                # 合并同一页的所有文本（不分栏）
                full_text = "\n".join(page_texts[page])
                f.write(full_text.replace(" ","") + "\n")

        logger.info(f"处理完成! 结果已保存至: {output_dir}")
        logger.info(f"JSON文件: {json_filename}")
        logger.info(f"文本文件: {txt_filename}")
        logger.info(f"总处理时长: {duration}")

    def calculate_column_stats(self, results):
        """计算分栏统计信息"""
        page_stats = {}
        for res in results:
            page = res['page']
            if page not in page_stats:
                page_stats[page] = {"columns": set(), "split_points": []}

            if res['column'] > 0:
                page_stats[page]["columns"].add(res['column'])
            if res['split_point']:
                page_stats[page]["split_points"].append(res['split_point'])

        single_count = 0
        double_count = 0

        for page, stats in page_stats.items():
            if len(stats["columns"]) == 2:
                double_count += 1
            else:
                single_count += 1

        return {
            "single": single_count,
            "double": double_count,
            "total_pages": len(page_stats)
        }


def main():
    parser = argparse.ArgumentParser(description="智能分栏文档OCR处理器")
    parser.add_argument("input_pdf", help="输入PDF文件路径")
    parser.add_argument("output_dir", help="输出目录路径")
    parser.add_argument("--dpi", type=int, default=300, help="扫描分辨率(默认:300)")
    parser.add_argument("--lang", default="chi_sim+eng",
                        help="OCR语言(默认:chi_sim+eng)")
    parser.add_argument("--no-layout", action="store_true",
                        help="禁用布局分析(不分栏)")
    parser.add_argument("--no-easyocr", action="store_true",
                        help="禁用EasyOCR(强制使用Tesseract)")
    parser.add_argument("--gpu-mem", type=float, default=0.8,
                        help="GPU内存使用比例(0.1-0.9, 默认:0.8)")
    parser.add_argument("--crop-top", type=float, default=0.1,
                        help="顶部裁剪比例(0-0.5, 默认:0.1)")
    parser.add_argument("--crop-bottom", type=float, default=0.1,
                        help="底部裁剪比例(0-0.5, 默认:0.1)")
    parser.add_argument("--debug", action="store_true",
                        help="启用调试模式（保存中间图像）")

    args = parser.parse_args()

    processor = DocumentProcessor(
        dpi=args.dpi,
        lang=args.lang,
        use_easyocr=not args.no_easyocr,
        gpu_mem_ratio=args.gpu_mem,
        crop_top_ratio=args.crop_top,
        crop_bottom_ratio=args.crop_bottom
    )
    processor.layout_analysis = not args.no_layout

    processor.process_pdf(args.input_pdf, args.output_dir, enable_debug=args.debug)


if __name__ == "__main__":
    main()
