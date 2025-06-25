import requests
import json
import re
import os
import shutil
from datetime import datetime

try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except ImportError as e:
    PANDAS_AVAILABLE = False
    print("错误: pandas 库未找到或导入失败。无法从Excel读取配置或生成模板。")
    print(f"详细错误: {e}")
    print("请尝试运行 'pip install pandas openpyxl' 来安装它以启用此功能。")
try:
    from premailer import Premailer
    # 强制导入相关依赖
    import premailer.premailer
    import cssutils
    import cssselect
    PREMAILER_AVAILABLE = True
except ImportError as e: 
    PREMAILER_AVAILABLE = False
    print(f"Premailer导入失败: {e}")

try:
    from bs4 import BeautifulSoup
    # 强制导入lxml解析器
    import lxml
    import lxml.etree
    import lxml.html
    import bs4.builder._lxml
    BS4_AVAILABLE = True
except ImportError as e: 
    BS4_AVAILABLE = False
    print(f"BeautifulSoup/lxml导入失败: {e}")


# ===================== GUI 相关代码 =====================
try:
    from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                                QHBoxLayout, QPushButton, QTextEdit, QPlainTextEdit, QLabel, 
                                QFileDialog, QProgressBar, QTableWidget, QTableWidgetItem,
                                QTabWidget, QGroupBox, QMessageBox)
    from PyQt6.QtCore import QThread, pyqtSignal, Qt
    from PyQt6.QtGui import QFont
    import sys
    PYQT6_AVAILABLE = True
except ImportError:
    PYQT6_AVAILABLE = False

# --- 全局配置 (这些可以被Excel中的数据覆盖) ---
BASE_URL = "https://api.weixin.qq.com/cgi-bin"
WECHAT_IMG_DOMAINS = ["mmbiz.qlogo.cn", "mmbiz.qpic.cn"] 
ARCHIVED_FOLDER_NAME = "已发内容" # 移动已处理文件的子文件夹名
EXCEL_TEMPLATE_NAME = "wechat_config_template.xlsx"
STATISTICS_FILE = "wechat_statistics.json"  # 统计数据保存文件
# --- 全局配置结束 ---

# ===================== 统计数据管理 =====================
class StatisticsManager:
    """统计数据管理器，负责保存和加载历史统计数据"""
    
    def __init__(self, stats_file=STATISTICS_FILE):
        self.stats_file = stats_file
        self.ensure_stats_file()
    
    def ensure_stats_file(self):
        """确保统计文件存在"""
        if not os.path.exists(self.stats_file):
            self.save_statistics([])
    
    def load_statistics(self):
        """加载历史统计数据"""
        try:
            with open(self.stats_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                return data.get('history', [])
        except Exception as e:
            log_message(f"加载统计数据失败: {e}")
            return []
    
    def save_statistics(self, history_data):
        """保存统计数据"""
        try:
            data = {'history': history_data}
            with open(self.stats_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            log_message(f"保存统计数据失败: {e}")
            return False
    
    def add_record(self, account_name, stats, message_type, processing_time=None):
        """添加新的处理记录"""
        if processing_time is None:
            processing_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        record = {
            'timestamp': processing_time,
            'account_name': account_name,
            'message_type': message_type,
            'success_count': stats.get('success_count', 0),
            'fail_count': stats.get('fail_count', 0),
            'total_processed': stats.get('success_count', 0) + stats.get('fail_count', 0),
            'failed_items': stats.get('failed_items', [])
        }
        
        history = self.load_statistics()
        history.append(record)
        self.save_statistics(history)
        return record
    
    def clear_statistics(self):
        """清除所有统计数据"""
        return self.save_statistics([])
    
    def export_to_csv(self, csv_file):
        """导出统计数据到CSV文件"""
        try:
            import csv
            history = self.load_statistics()
            
            with open(csv_file, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                # 写入标题行
                writer.writerow(['处理时间', '账号名称', '消息类型', '成功数量', '失败数量', '总处理数', '失败详情'])
                
                # 写入数据行
                for record in history:
                    failed_items_str = '; '.join(record.get('failed_items', []))
                    writer.writerow([
                        record.get('timestamp', ''),
                        record.get('account_name', ''),
                        record.get('message_type', ''),
                        record.get('success_count', 0),
                        record.get('fail_count', 0),
                        record.get('total_processed', 0),
                        failed_items_str
                    ])
            return True
        except Exception as e:
            log_message(f"导出CSV失败: {e}")
            return False

# 全局日志函数
def log_message(message):
    """统一的日志输出函数"""
    if hasattr(log_message, 'callback') and log_message.callback:
        log_message.callback(message)
    else:
        print(message)

# 设置日志回调函数
log_message.callback = None

def set_log_callback(callback):
    """设置日志回调函数"""
    log_message.callback = callback

def _make_request(method, url, **kwargs):
    """统一处理 requests 请求，加入 proxies 参数"""
    # proxies 参数应形如: {'http': 'http://user:pass@host:port', 'https': 'http://user:pass@host:port'}
    # 或者 {'http': 'http://host:port', 'https': 'http://host:port'}
    # kwargs 中可以包含 proxies, timeout, stream, files, data, headers 等
    
    # 确保超时设置
    if 'timeout' not in kwargs:
        if method.upper() == 'POST': # 上传文件可能需要更长时间
            kwargs['timeout'] = kwargs.get('files', None) and 120 or 60 
        else:
            kwargs['timeout'] = 30
            
    return requests.request(method, url, **kwargs)

def get_access_token(appid, appsecret, proxies=None):
    url = f"{BASE_URL}/token?grant_type=client_credential&appid={appid}&secret={appsecret}"
    try:
        response = _make_request("get", url, proxies=proxies)
        response.raise_for_status()
        data = response.json()
        if "access_token" in data:
            return data["access_token"]
        else:
            log_message("  获取access_token失败 (AppID: " + str(appid) + "): " + str(data))
            return None
    except requests.exceptions.RequestException as e:
        log_message("  请求access_token时发生错误 (AppID: " + str(appid) + "): " + str(e))
        return None
    except json.JSONDecodeError as e:
        log_message("  解析access_token响应JSON时出错 (AppID: " + str(appid) + "): " + str(e))
        return None

def download_image_from_url(image_url, local_filename, proxies=None):
    try:
        log_message("    下载图片从: " + str(image_url) + " -> " + str(local_filename))
        
        # 设置请求头，模拟浏览器访问，避免防盗链
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
            'Referer': 'https://www.baidu.com/',
            'Accept': 'image/webp,image/apng,image/*,*/*;q=0.8',
            'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
        }
        
        response = _make_request("get", image_url, stream=True, proxies=proxies, headers=headers)
        response.raise_for_status()
        
        # 检查响应内容类型
        content_type = response.headers.get('content-type', '').lower()
        if not any(img_type in content_type for img_type in ['image/', 'application/octet-stream']):
            log_message(f"    警告: 响应内容类型不是图片 ({content_type})")
        
        with open(local_filename, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)
        
        # 验证下载的文件大小
        file_size = os.path.getsize(local_filename)
        if file_size < 100:  # 如果文件太小，可能是错误页面
            log_message(f"    警告: 下载的文件太小 ({file_size} bytes)，可能下载失败")
            return False
            
        log_message(f"    图片下载成功，文件大小: {file_size} bytes")
        return True
        
    except requests.exceptions.RequestException as e:
        log_message("    下载图片失败 (" + str(image_url) + "): " + str(e))
        return False
    except IOError as e:
        log_message("    保存图片失败 (" + str(local_filename) + "): " + str(e))
        return False

def upload_permanent_material(access_token, file_path, material_type='image', appid_for_log="", proxies=None):
    url = f"{BASE_URL}/material/add_material?access_token={access_token}&type={material_type}"
    log_prefix = f"(AppID: {appid_for_log}) " if appid_for_log else ""
    response = None 
    try:
        log_message("    " + log_prefix + "上传本地素材 " + str(file_path) + " (类型: " + str(material_type) + ")...")
        with open(file_path, 'rb') as f:
            file_name_for_upload = os.path.basename(file_path)
            mime_type = 'image/jpeg' # 默认，可根据需要扩展
            if file_name_for_upload.lower().endswith('.png'):
                mime_type = 'image/png'
            elif file_name_for_upload.lower().endswith('.gif'):
                mime_type = 'image/gif'
            
            files = {'media': (file_name_for_upload, f, mime_type if material_type == 'image' else 'application/octet-stream')}
            response = _make_request("post", url, files=files, proxies=proxies)
            response.raise_for_status()
        result = response.json()
        if "media_id" in result:
            wx_media_id = result["media_id"]
            wx_image_url = result.get("url") 
            log_message("    " + log_prefix + "素材上传成功！Media ID: " + str(wx_media_id) + (", URL: " + str(wx_image_url) if wx_image_url else ""))
            return {"media_id": wx_media_id, "url": wx_image_url}
        else:
            log_message("    " + log_prefix + "上传永久素材失败: " + str(result))
            return None
    except requests.exceptions.RequestException as e:
        log_message("    " + log_prefix + "请求上传永久素材错误: " + str(e))
        return None
    except IOError as e:
        log_message("    " + log_prefix + "读取本地文件错误 " + str(file_path) + ": " + str(e))
        return None
    except json.JSONDecodeError:
        response_text = response.text if response is not None and hasattr(response, 'text') else 'No response object or text attribute'
        log_message("    " + log_prefix + "无法解析上传素材响应: " + str(response_text))
        return None

def optimize_html_with_inline_styles(html_string):
    if not PREMAILER_AVAILABLE:
        log_message("    Premailer 库不可用，跳过HTML样式内联优化。")
        return html_string
    try:
        # 清理可能导致PyQt6警告的CSS属性
        cleaned_html = html_string
        # 移除可能导致警告的CSS属性
        import re
        cleaned_html = re.sub(r'word-wrap\s*:[^;]+;?', '', cleaned_html, flags=re.IGNORECASE)
        cleaned_html = re.sub(r'break-word\s*:[^;]+;?', '', cleaned_html, flags=re.IGNORECASE)
        cleaned_html = re.sub(r'width\s*:\s*fit-content\s*;?', 'width: auto;', cleaned_html, flags=re.IGNORECASE)
        cleaned_html = re.sub(r'height\s*:\s*fit-content\s*;?', 'height: auto;', cleaned_html, flags=re.IGNORECASE)
        
        p = Premailer(cleaned_html, remove_classes=False, keep_style_tags=True, strip_important=False)
        inlined_html = p.transform()
        return inlined_html
    except Exception as e:
        log_message("    Premailer优化错误: " + str(e) + "。使用原始HTML。")
        return html_string

def replace_external_images_in_html(html_content, access_token, appid_for_log="", current_html_file_path="", proxies=None):
    if not BS4_AVAILABLE:
        log_message("    BeautifulSoup4 库不可用，跳过正文图片链接替换。")
        return html_content
    if not access_token:
        return html_content

    try:
        soup = BeautifulSoup(html_content, 'lxml')
    except Exception as e:
        log_message("    BeautifulSoup解析HTML失败: " + str(e) + "。跳过图片替换。")
        return html_content
        
    img_tags = soup.find_all('img')
    image_counter = 0
    processed_image_count = 0

    for i, img in enumerate(img_tags):
        original_src = img.get('src')
        if not original_src:
            continue
        is_external = original_src.startswith(('http://', 'https://'))
        is_wechat_domain = False
        if is_external:
            try:
                domain = original_src.split('//')[1].split('/')[0].lower()
                if any(wx_domain in domain for wx_domain in WECHAT_IMG_DOMAINS):
                    is_wechat_domain = True
            except IndexError:
                pass 
        
        if is_external and not is_wechat_domain:
            image_counter += 1
            log_message("      处理第" + str(image_counter) + "个外部图片: " + original_src[:70] + ('...' if len(original_src)>70 else ''))
            base_html_filename = os.path.splitext(os.path.basename(current_html_file_path))[0]
            temp_img_filename = f"temp_body_img_{appid_for_log.replace('.', '_')}_{base_html_filename}_{i}.jpg"
            
            try:
                if download_image_from_url(original_src, temp_img_filename, proxies=proxies):
                    upload_result = upload_permanent_material(access_token, temp_img_filename, 'image', appid_for_log, proxies=proxies)
                    if upload_result and upload_result.get("url"):
                        wx_image_url = upload_result["url"]
                        img['src'] = wx_image_url 
                        log_message("        ✓ 成功替换为微信图片URL: " + str(wx_image_url))
                        processed_image_count +=1
                    else:
                        log_message("        ✗ 上传失败或未返回URL，保留原始src")
                else:
                    log_message("        ✗ 下载失败，保留原始src")
            except Exception as e:
                log_message(f"        ✗ 处理图片时发生异常: {str(e)}，保留原始src")
            finally:
                # 确保清理临时文件
                if os.path.exists(temp_img_filename):
                    try: 
                        os.remove(temp_img_filename)
                    except OSError as e: 
                        log_message("        删除临时文件 " + str(temp_img_filename) + " 失败: " + str(e))

    if image_counter > 0:
        log_message("    共找到" + str(image_counter) + "个外部图片链接，成功处理了" + str(processed_image_count) + "个。")
    
    # 修复BeautifulSoup的HTML输出格式问题
    # 使用soup.body.decode_contents()保留body内的内容，避免添加html/body标签
    try:
        if soup.body:
            result_html = soup.body.decode_contents()
        else:
            # 如果没有body标签，尝试获取所有内容
            result_html = str(soup)
            # 移除lxml自动添加的html和body标签
            if result_html.startswith('<html><body>') and result_html.endswith('</body></html>'):
                result_html = result_html[12:-14]  # 移除<html><body>和</body></html>
        
        log_message(f"    HTML处理完成，输出长度: {len(result_html)}")
        return result_html
    except Exception as e:
        log_message(f"    HTML格式化失败: {e}，使用原始HTML")
        return html_content

def create_draft_api(access_token, articles_data, appid_for_log="", proxies=None, show_content=True):
    url = f"{BASE_URL}/draft/add?access_token={access_token}"
    headers = {"Content-Type": "application/json"}
    log_prefix = f"(AppID: {appid_for_log}) " if appid_for_log else ""
    
    # 记录发送的数据用于调试
    log_message("    准备发送草稿数据...")
    article = articles_data.get("articles", [{}])[0]
    log_message(f"    标题: {article.get('title', 'N/A')}")
    log_message(f"    作者: {article.get('author', 'N/A')}")
    log_message(f"    封面Media ID: {article.get('thumb_media_id', 'N/A')}")
    log_message(f"    内容长度: {len(article.get('content', ''))}")
    
    # 调试：输出部分HTML内容用于分析（仅图文消息）
    if show_content:
        content = article.get('content', '')
        if content:
            log_message(f"    内容开头200字符: {content[:200]}...")
            log_message(f"    内容结尾200字符: ...{content[-200:]}")
            
            # 检查是否包含可能有问题的标签
            problematic_patterns = [
                r'<script[^>]*>',
                r'<style[^>]*>',
                r'<iframe[^>]*>',
                r'<object[^>]*>',
                r'<embed[^>]*>',
                r'<form[^>]*>',
                r'<input[^>]*>',
                r'<button[^>]*>',
                r'<link[^>]*>',
                r'<meta[^>]*>',
                r'on\w+\s*=',
                r'javascript:',
                r'<mp-[^>]*>',
                r'data-miniprogram-[^=]*='
            ]
            
            for pattern in problematic_patterns:
                matches = re.findall(pattern, content, re.IGNORECASE)
                if matches:
                    log_message(f"    发现可能有问题的标签/属性: {pattern} -> {matches[:3]}")
    
    try:
        response = _make_request("post", url, headers=headers, data=json.dumps(articles_data, ensure_ascii=False).encode('utf-8'), proxies=proxies)
        response.raise_for_status()
        result = response.json()
        
        if result.get("errcode") == 0 or "media_id" in result:
            log_message("  " + log_prefix + "草稿API响应成功: " + str(result)) 
            return result.get("media_id") if "media_id" in result else True
        else:
            errcode = result.get("errcode", "未知")
            errmsg = result.get("errmsg", "未知错误")
            log_message("  " + log_prefix + f"创建草稿失败 (错误码: {errcode}): {errmsg}")
            
            # 提供具体的错误解释
            error_explanations = {
                40001: "AppSecret错误或者AppSecret不属于这个公众号",
                40007: "不合法的媒体文件id (封面图片media_id无效)",
                40008: "不合法的消息类型",
                40009: "不合法的图片文件大小",
                40013: "不合法的AppID",
                41001: "缺少access_token参数",
                41002: "缺少appid参数",
                41003: "缺少refresh_token参数",
                41004: "缺少secret参数",
                41005: "缺少多媒体文件数据",
                41006: "缺少media_id参数",
                42001: "access_token超时",
                45001: "多媒体文件大小超过限制",
                45002: "消息内容超过限制",
                45003: "标题字段超过限制",
                45004: "描述字段超过限制",
                45007: "语音播放时间超过限制",
                45008: "图文消息超过限制",
                45009: "接口调用超过限制",
                45166: "内容格式无效 - 可能包含不支持的HTML标签、图片格式问题或内容审核不通过"
            }
            
            if errcode in error_explanations:
                log_message("  " + log_prefix + f"错误解释: {error_explanations[errcode]}")
                
            if errcode == 40007:
                log_message("  " + log_prefix + "建议: 检查封面图片是否成功上传，或尝试使用其他图片")
            elif errcode == 42001:
                log_message("  " + log_prefix + "建议: access_token已过期，需要重新获取")
            elif errcode == 45166:
                log_message("  " + log_prefix + "建议: 检查以下问题:")
                log_message("  " + log_prefix + "  1) HTML内容是否包含不支持的标签(如script、style等)")
                log_message("  " + log_prefix + "  2) 图片链接是否都来自微信域名")
                log_message("  " + log_prefix + "  3) 图片宽高比例是否合适")
                log_message("  " + log_prefix + "  4) 内容是否包含敏感词汇")
                log_message("  " + log_prefix + "  5) 检查是否有小程序相关标签格式错误")
            
            return None
            
    except requests.exceptions.RequestException as e:
        log_message("  " + log_prefix + "请求创建草稿时发生错误: " + str(e))
        return None
    except json.JSONDecodeError:
        response_text = response.text if 'response' in locals() and hasattr(response, 'text') else 'No response text available'
        log_message("  " + log_prefix + "无法解析草稿创建响应: " + str(response_text))
        return None


def clean_and_normalize_text(text_content):
    """严格清理和规范化文本内容，移除所有可能导致45166错误的字符"""
    if not text_content:
        return ""
    
    import unicodedata
    import re
    
    # 步骤1: 基础清理
    cleaned_text = text_content.strip()
    
    # 移除所有制表符和特殊空白字符（更全面）
    cleaned_text = re.sub(r'[\t\r\f\v\u00A0]', ' ', cleaned_text)  # 将tab和不间断空格转换为普通空格
    cleaned_text = re.sub(r' +', ' ', cleaned_text)  # 合并多个空格为单个空格
    
    # 步骤2: Unicode规范化
    cleaned_text = unicodedata.normalize('NFC', cleaned_text)
    
    # 步骤3: 移除变体选择符和零宽字符
    cleaned_text = re.sub(r'[\uFE00-\uFE0F\u200B-\u200D\uFEFF\u2060]', '', cleaned_text)
    
    # 步骤4: 替换危险的Unicode字符
    char_replacements = {
        # 空格和分隔符
        '\u2028': '\n',      # 行分隔符
        '\u2029': '\n\n',    # 段落分隔符
        '\u00A0': ' ',       # 不间断空格
        '\u2060': '',        # 零宽无断空格
        
        # 箭头和符号（确定导致45166错误的字符）
        '\u2192': ' -> ',    # 右箭头 →
        '\u2190': ' <- ',    # 左箭头 ←
        '\u2191': ' ^ ',     # 上箭头 ↑
        '\u2193': ' v ',     # 下箭头 ↓
        
        # 破折号（确定导致45166错误的字符）
        '\u2014': '-',       # EM DASH —
        '\u2013': '-',       # EN DASH –
        '\u2026': '...',     # 省略号 …
        
        # 引号规范化
        '\u201C': '"',       # 左双引号 "
        '\u201D': '"',       # 右双引号 "
        '\u2018': "'",       # 左单引号 '
        '\u2019': "'",       # 右单引号 '
        
        # 其他符号
        '\u2022': '•',       # 项目符号 •
        '\u2023': '►',       # 三角项目符号
        '\u25B6': '►',       # 播放符号 ▶
        
        # 数学符号
        '\u00D7': 'x',       # 乘号 ×
        '\u00F7': '/',       # 除号 ÷
        '\u2212': '-',       # 减号 −
        
        # 从文件分析中发现的额外问题字符
        '\u26A0': '⚠',       # 警告符号 ⚠ 
        '\u2705': '✓',       # 白色重复选中标记 ✅
        '\u2728': '✨',       # 闪亮 ✨
        '\u274C': '✗',       # 叉号 ❌
        '\uFE0F': '',        # 变体选择符-16（移除）
    }
    
    for old_char, new_char in char_replacements.items():
        cleaned_text = cleaned_text.replace(old_char, new_char)
    
    return cleaned_text

def convert_text_to_plain_for_pic_message(text_content):
    """将纯文本内容转换为图片消息格式（纯文本，不支持HTML）"""
    if not text_content:
        return ""
    
    # 清理和规范化文本内容
    cleaned_text = clean_and_normalize_text(text_content)
    
    # 规范化确实有问题的emoji
    emoji_map = {
        '1⃣️': '1️⃣', '2⃣️': '2️⃣', '3⃣️': '3️⃣', '4⃣️': '4️⃣', '5⃣️': '5️⃣',
        '6⃣️': '6️⃣', '7⃣️': '7️⃣', '8⃣️': '8️⃣', '9⃣️': '9️⃣',
        '[赞R]': '👍', '[强]': '💪', '[握手]': '🤝',
        '➡️': '->', '⬅️': '<-', '⬆️': '^', '⬇️': 'v',
    }
    
    # 应用emoji映射
    for old_emoji, new_emoji in emoji_map.items():
        cleaned_text = cleaned_text.replace(old_emoji, new_emoji)
    
    # 图片消息使用纯文本格式
    plain_content = cleaned_text.strip()
    
    return plain_content

def process_single_picture_folder(folder_path, article_config, access_token, proxies=None):
    """处理单个图片消息文件夹"""
    appid_for_log = article_config.get('appid', 'N/A')
    log_prefix = f"(AppID: {appid_for_log}) "
    
    log_message(f"  处理图片消息文件夹: {folder_path}")
    
    # 获取文件夹名作为标题
    folder_name = os.path.basename(folder_path)
    
    # 查找txt文件
    txt_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.txt')]
    if not txt_files:
        log_message("    错误：未找到txt文件，跳过此文件夹")
        return False
    
    if len(txt_files) > 1:
        log_message(f"    警告：找到多个txt文件，使用第一个：{txt_files[0]}")
    
    txt_file_path = os.path.join(folder_path, txt_files[0])
    
    # 读取txt文件内容
    try:
        with open(txt_file_path, 'r', encoding='utf-8') as f:
            text_content = f.read().strip()
    except Exception as e:
        log_message(f"    错误：读取txt文件失败 {txt_file_path}: {e}")
        return False
    
    # 将txt内容转换为图片消息纯文本格式
    content = convert_text_to_plain_for_pic_message(text_content)
    
    # 查找所有图片文件，按文件名数字排序
    image_files = []
    for f in os.listdir(folder_path):
        if f.lower().endswith(('.jpg', '.jpeg', '.png', '.gif')):
            image_files.append(f)
    
    if not image_files:
        log_message("    错误：未找到图片文件，跳过此文件夹")
        return False
    
    # 按文件名中的数字排序
    def extract_number(filename):
        numbers = re.findall(r'\d+', filename)
        return int(numbers[0]) if numbers else float('inf')
    
    image_files.sort(key=extract_number)
    log_message(f"    找到 {len(image_files)} 个图片文件，按顺序处理...")
    
    # 上传所有图片
    uploaded_images = []
    failed_images = []
    
    for i, image_file in enumerate(image_files):
        image_path = os.path.join(folder_path, image_file)
        log_message(f"      上传第 {i+1} 个图片: {image_file}")
        
        upload_result = upload_permanent_material(access_token, image_path, 'image', appid_for_log, proxies=proxies)
        if upload_result and upload_result.get("media_id"):
            uploaded_images.append({
                "image_media_id": upload_result["media_id"]
            })
            log_message(f"        ✓ 上传成功，Media ID: {upload_result['media_id']}")
        else:
            log_message(f"        ✗ 上传失败: {image_file}")
            failed_images.append(image_file)
    
    # 检查上传结果
    if not uploaded_images:
        log_message("    错误：没有成功上传任何图片")
        return False
    elif failed_images:
        log_message(f"    警告：{len(failed_images)} 张图片上传失败，已跳过: {', '.join(failed_images)}")
        log_message(f"    继续处理剩余的 {len(uploaded_images)} 张图片")
    
    # 构建图片消息API数据
    is_comment_enabled = article_config.get('is_comment_enabled', False)
    comment_permission = article_config.get('comment_permission', '所有人')
    need_open_comment = int(1 if is_comment_enabled else 0)
    only_fans_can_comment = int(1 if comment_permission == '仅粉丝' else 0)
    
    # 获取txt文件名作为标题（去掉扩展名）
    title = os.path.splitext(txt_files[0])[0]
    
    articles_data = {
        "articles": [{
            "article_type": "newspic",
            "title": title,
            "content": content,
            "need_open_comment": need_open_comment,
            "only_fans_can_comment": only_fans_can_comment,
            "image_info": {
                "image_list": uploaded_images
            }
        }]
    }
    
    log_message("    创建图片消息草稿...")
    success = create_draft_api(access_token, articles_data, appid_for_log, proxies=proxies, show_content=False)
    return success

def process_picture_message_folders(articles_folder_path, article_config, access_token, num_to_publish, proxies=None):
    """处理图片消息文件夹"""
    log_message(f"  开始处理图片消息文件夹模式...")
    
    # 获取所有子文件夹
    try:
        subfolders = [f for f in os.listdir(articles_folder_path) 
                     if os.path.isdir(os.path.join(articles_folder_path, f)) 
                     and f != ARCHIVED_FOLDER_NAME]
        subfolders.sort()  # 按名称排序
    except Exception as e:
        log_message(f"    错误：读取图片消息目录失败: {e}")
        return 0
    
    if not subfolders:
        log_message("    未找到图片消息子文件夹")
        return 0
    
    processed_count = 0
    
    for i, subfolder in enumerate(subfolders):
        if processed_count >= num_to_publish:
            log_message(f"    已达存稿上限 ({num_to_publish})")
            break
            
        subfolder_path = os.path.join(articles_folder_path, subfolder)
        log_message(f"\n    [{i+1}/{len(subfolders)}] 处理子文件夹: {subfolder}")
        
        if process_single_picture_folder(subfolder_path, article_config, access_token, proxies):
            log_message(f"      图片消息 '{subfolder}' 创建成功")
            
            # 移动整个文件夹到已发内容
            archived_dir = os.path.join(articles_folder_path, ARCHIVED_FOLDER_NAME)
            if not os.path.exists(archived_dir):
                try:
                    os.makedirs(archived_dir)
                    log_message(f"      创建文件夹: {archived_dir}")
                except OSError as e:
                    log_message(f"      错误：创建已发内容文件夹失败: {e}")
                    continue
            
            destination_path = os.path.join(archived_dir, subfolder)
            log_message(f"      准备移动文件夹: '{subfolder_path}' -> '{destination_path}'")
            try:
                shutil.move(subfolder_path, destination_path)
                log_message("      文件夹已移动")
                processed_count += 1
            except Exception as e:
                log_message(f"      移动文件夹失败: {e}")
        else:
            log_message(f"      处理图片消息失败: {subfolder}")
    
    return processed_count

def process_single_article(article_config, access_token, proxies=None):
    log_message("  处理文章: " + str(article_config['html_file_full_path']))
    appid_for_log = article_config.get('appid', 'N/A')
    current_html_file_path = article_config['html_file_full_path']
    raw_html_content = ""
    try:
        with open(current_html_file_path, 'r', encoding='utf-8') as f: raw_html_content = f.read()
    except FileNotFoundError: 
        log_message("    错误：找不到文件 " + str(current_html_file_path) + "。跳过。") 
        return False
    except Exception as e: 
        log_message("    读取文件 " + str(current_html_file_path) + " 错误: " + str(e) + "。跳过。")
        return False

    log_message("    步骤1: Premailer CSS内联优化...")
    optimized_html_content = optimize_html_with_inline_styles(raw_html_content)
    log_message(f"    优化后HTML长度: {len(optimized_html_content)}")
    
    log_message("    步骤2: 替换正文外部图片链接...")
    html_with_wechat_images = replace_external_images_in_html(optimized_html_content, access_token, appid_for_log, current_html_file_path, proxies=proxies)
    log_message(f"    图片处理后HTML长度: {len(html_with_wechat_images)}")
    
    # 检查图片处理后是否还有img标签
    img_count_after = len(re.findall(r'<img[^>]*>', html_with_wechat_images, re.IGNORECASE))
    log_message(f"    图片处理后剩余图片数量: {img_count_after}")
    
    log_message("    步骤3: 正则清理HTML...")
    cleaned_html = html_with_wechat_images
    
    # 移除危险的HTML标签
    cleaned_html = re.sub(r'<script[^>]*>.*?</script>', '', cleaned_html, flags=re.IGNORECASE | re.DOTALL)
    cleaned_html = re.sub(r'<style[^>]*>.*?</style>', '', cleaned_html, flags=re.IGNORECASE | re.DOTALL)
    cleaned_html = re.sub(r'<link[^>]*>', '', cleaned_html, flags=re.IGNORECASE)
    cleaned_html = re.sub(r'<meta[^>]*>', '', cleaned_html, flags=re.IGNORECASE)
    cleaned_html = re.sub(r'<iframe[^>]*>.*?</iframe>', '', cleaned_html, flags=re.IGNORECASE | re.DOTALL)
    cleaned_html = re.sub(r'<object[^>]*>.*?</object>', '', cleaned_html, flags=re.IGNORECASE | re.DOTALL)
    cleaned_html = re.sub(r'<embed[^>]*>', '', cleaned_html, flags=re.IGNORECASE)
    cleaned_html = re.sub(r'<form[^>]*>.*?</form>', '', cleaned_html, flags=re.IGNORECASE | re.DOTALL)
    cleaned_html = re.sub(r'<input[^>]*>', '', cleaned_html, flags=re.IGNORECASE)
    cleaned_html = re.sub(r'<button[^>]*>.*?</button>', '', cleaned_html, flags=re.IGNORECASE | re.DOTALL)
    cleaned_html = re.sub(r'<select[^>]*>.*?</select>', '', cleaned_html, flags=re.IGNORECASE | re.DOTALL)
    cleaned_html = re.sub(r'<textarea[^>]*>.*?</textarea>', '', cleaned_html, flags=re.IGNORECASE | re.DOTALL)
    
    # 移除小程序相关标签
    cleaned_html = re.sub(r'<mp-[^>]*>.*?</mp-[^>]*>', '', cleaned_html, flags=re.IGNORECASE | re.DOTALL)
    cleaned_html = re.sub(r'<mp-[^>]*>', '', cleaned_html, flags=re.IGNORECASE)
    
    # 清理小程序相关属性
    cleaned_html = re.sub(r'\s*data-miniprogram-[^=]*=["\'][^"\']*["\']', '', cleaned_html, flags=re.IGNORECASE)
    
    # 清理事件处理器
    cleaned_html = re.sub(r'\s*on\w+\s*=\s*["\'][^"\']*["\']', '', cleaned_html, flags=re.IGNORECASE)
    
    # 清理javascript:链接
    cleaned_html = re.sub(r'href\s*=\s*["\']javascript:[^"\']*["\']', 'href="#"', cleaned_html, flags=re.IGNORECASE)
    
    # 清理其他可能有问题的属性
    cleaned_html = re.sub(r'\s*contenteditable\s*=\s*["\'][^"\']*["\']', '', cleaned_html, flags=re.IGNORECASE)
    cleaned_html = re.sub(r'\s*draggable\s*=\s*["\'][^"\']*["\']', '', cleaned_html, flags=re.IGNORECASE)
    
    # 只清理明显的空段落，保持原有换行格式
    cleaned_html = re.sub(r'<p\b[^>]*>\s*(?:&nbsp;|\s)*\s*</p>', '', cleaned_html, flags=re.IGNORECASE | re.DOTALL)
    # 移除过于激进的换行清理，保持原有HTML格式
    
    # 最终安全检查：移除任何剩余的危险元素
    dangerous_tags = ['script', 'style', 'iframe', 'object', 'embed', 'form', 'input', 'button', 'select', 'textarea', 'link', 'meta']
    for tag in dangerous_tags:
        cleaned_html = re.sub(f'<{tag}[^>]*>.*?</{tag}>', '', cleaned_html, flags=re.IGNORECASE | re.DOTALL)
        cleaned_html = re.sub(f'<{tag}[^>]*/?>', '', cleaned_html, flags=re.IGNORECASE)
    
    # 移除HTML5特殊属性和data-*属性（除了微信图片）
    cleaned_html = re.sub(r'\s*data-(?!src)[^=]*=["\'][^"\']*["\']', '', cleaned_html, flags=re.IGNORECASE)
    
    # Emoji和特殊字符规范化
    log_message("    步骤4: 规范化emoji和特殊字符...")
    
    # 定义emoji映射表（图文消息保守处理）
    emoji_map = {
        # 只处理确认有问题的数字emoji变体
        '1⃣️': '1️⃣',
        '2⃣️': '2️⃣', 
        '3⃣️': '3️⃣',
        '4⃣️': '4️⃣',
        '5⃣️': '5️⃣',
        '6⃣️': '6️⃣',
        '7⃣️': '7️⃣',
        '8⃣️': '8️⃣',
        '9⃣️': '9️⃣',
        
        # 文本符号替换
        '[赞R]': '👍',
        '[强]': '💪',
        '[握手]': '🤝',
        
        # 图文消息中保持大部分emoji原样，只替换确认有问题的
        # 移除过度的emoji转换，保持原有样式
    }
    
    # 记录处理前后的示例（在规范化前）
    log_message(f"    规范化前示例: {repr(cleaned_html[:100])}")
    
    # 应用emoji映射
    changes_made = []
    for old_emoji, new_emoji in emoji_map.items():
        if old_emoji in cleaned_html:
            cleaned_html = cleaned_html.replace(old_emoji, new_emoji)
            changes_made.append(f"{old_emoji} -> {new_emoji}")
    
    if changes_made:
        log_message(f"    发现并替换的emoji: {', '.join(changes_made[:5])}")
    else:
        log_message("    未发现需要替换的emoji")
    
    # 将复杂的emoji转换为标准Unicode
    import unicodedata
    
    def normalize_emoji(text):
        # 规范化Unicode字符
        normalized = unicodedata.normalize('NFC', text)
        
        # 移除变体选择符和零宽字符
        normalized = re.sub(r'[\uFE00-\uFE0F\u200B-\u200D\uFEFF\u2060]', '', normalized)
        
        # 将问题字符转换为安全替代（图文消息保持HTML格式）
        char_map = {
            # 只处理确认会导致问题的字符，保持HTML格式
            '\u00A0': '&nbsp;',  # 不间断空格 → HTML实体
            '\u2060': '',        # 零宽无断空格
            
            # 箭头和符号
            '\u2192': ' → ',     # 右箭头保持原样（HTML中一般没问题）
            '\u2014': '—',       # EM DASH保持原样
            '\u2013': '–',       # EN DASH保持原样
            '\u2026': '…',       # 省略号保持原样
            
            # 引号保持原样（HTML中一般没问题）
            '\u201C': '"',       # 左双引号
            '\u201D': '"',       # 右双引号
            '\u2018': "'",       # 左单引号
            '\u2019': "'",       # 右单引号
            
            # 其他符号保持原样
            '\u2022': '•',       # 项目符号
            '\u2023': '▸',       # 三角项目符号
            '\u25B6': '▶',       # 播放符号
        }
        
        for old_char, new_char in char_map.items():
            normalized = normalized.replace(old_char, new_char)
            
        return normalized
    
    cleaned_html = normalize_emoji(cleaned_html)
    
    # 记录处理后的示例
    log_message(f"    规范化后示例: {repr(cleaned_html[:100])}")
    
    final_html_content_for_api = cleaned_html
    
    # 最终检查图片数量
    final_img_count = len(re.findall(r'<img[^>]*>', final_html_content_for_api, re.IGNORECASE))
    log_message(f"    最终HTML长度: {len(final_html_content_for_api)}, 图片数量: {final_img_count}")
    log_message("    步骤4: 准备封面图...")
    
    # 查找封面图片URL（优先使用原始HTML中的图片URL，避免微信防盗链问题）
    image_matches = re.findall(r'<img [^>]*src="([^"]+)"', raw_html_content, re.IGNORECASE)
    
    # 如果原始HTML中没有图片，再尝试使用处理过的HTML
    if not image_matches:
        image_matches = re.findall(r'<img [^>]*src="([^"]+)"', html_with_wechat_images, re.IGNORECASE)
    if not image_matches:
        log_message("    HTML中未找到任何图片")
        log_message("    警告: 微信API要求图文消息必须有封面图片。")
        log_message("    建议: 在HTML中添加至少一张图片")
        return False
    
    log_message(f"    在HTML中找到 {len(image_matches)} 张图片，将依次尝试作为封面")
    
    actual_thumb_media_id = None
    base_cover_html_filename = os.path.splitext(os.path.basename(current_html_file_path))[0]
    
    # 依次尝试每张图片作为封面
    for i, cover_image_url in enumerate(image_matches):
        log_message(f"    尝试第 {i+1} 张图片作为封面: {cover_image_url[:80]}{'...' if len(cover_image_url) > 80 else ''}")
        
        # 检查是否已经是微信域名的图片
        is_wechat_image = False
        try:
            domain = cover_image_url.split('//')[1].split('/')[0].lower() if '//' in cover_image_url else ''
            if any(wx_domain in domain for wx_domain in WECHAT_IMG_DOMAINS):
                is_wechat_image = True
                log_message("    这是微信域名的图片，直接上传获取media_id...")
        except (IndexError, AttributeError):
            pass
        
        temp_cover_filename = f"temp_cover_{appid_for_log.replace('.', '_')}_{base_cover_html_filename}_{i}.jpg"
        
        try:
            # 尝试下载图片
            if download_image_from_url(cover_image_url, temp_cover_filename, proxies=proxies):
                log_message("    图片下载成功，开始上传到微信...")
                
                # 尝试上传到微信
                upload_result = upload_permanent_material(access_token, temp_cover_filename, 'image', appid_for_log, proxies=proxies)
                if upload_result and upload_result.get("media_id"): 
                    actual_thumb_media_id = upload_result["media_id"]
                    log_message(f"    ✓ 封面图片上传成功！Media ID: {actual_thumb_media_id}")
                    
                    # 清理临时文件
                    if os.path.exists(temp_cover_filename): 
                        try: 
                            os.remove(temp_cover_filename)
                            log_message("    临时封面文件已删除")
                        except OSError as e: 
                            log_message("    删除临时封面文件失败: " + str(e))
                    
                    # 成功获取封面，跳出循环
                    break
                else:
                    log_message("    图片上传失败，尝试下一张图片...")
            else:
                log_message("    图片下载失败，尝试下一张图片...")
                
        except Exception as e:
            log_message(f"    处理第 {i+1} 张图片时发生异常: {str(e)}")
            
        finally:
            # 清理可能存在的临时文件
            if os.path.exists(temp_cover_filename):
                try:
                    os.remove(temp_cover_filename)
                except OSError:
                    pass
    
    # 检查是否成功获取封面图片
    if not actual_thumb_media_id: 
        log_message("    ✗ 所有图片都无法作为封面使用")
        log_message("    可能的原因:")
        log_message("      1) 所有图片链接都无法访问或有防盗链保护")
        log_message("      2) 图片格式不被微信支持")
        log_message("      3) 图片文件太大或太小")
        log_message("      4) 网络连接问题")
        log_message("    建议:")
        log_message("      1) 使用稳定的图床服务（如微信公众平台、七牛云等）")
        log_message("      2) 检查图片链接是否可以在浏览器中正常打开")
        log_message("      3) 确保图片格式为 JPG、PNG 或 GIF")
        log_message("      4) 图片大小控制在 5MB 以内")
        return False  # 没有封面图片，直接返回失败
    else:
        log_message(f"    ✓ 成功设置封面图片，Media ID: {actual_thumb_media_id}")
    
    is_comment_enabled = article_config.get('is_comment_enabled', False)
    comment_permission = article_config.get('comment_permission', '所有人')
    need_open_comment = int(1 if is_comment_enabled else 0)
    only_fans_can_comment = int(1 if comment_permission == '仅粉丝' else 0)

    log_message("    步骤5: 创建草稿...")
    article_title = os.path.splitext(os.path.basename(current_html_file_path))[0]
    
    # 生成摘要（取正文前54个字符，移除HTML标签）
    plain_text = re.sub(r'<[^>]+>', '', final_html_content_for_api)
    digest = plain_text[:54].strip() if plain_text else ""
    
    # 检查内容长度限制
    content_byte_size = len(final_html_content_for_api.encode('utf-8'))
    if content_byte_size > 1024 * 1024:  # 1MB
        log_message(f"    警告: HTML内容大小 {content_byte_size} 字节，超过1MB限制")
        
    if len(final_html_content_for_api) > 20000:  # 2万字符
        log_message(f"    警告: HTML内容长度 {len(final_html_content_for_api)} 字符，超过2万字符限制")
    
    articles_data = {
        "articles": [{
            "article_type": "news",  # 明确指定为图文消息
            "title": article_title,
            "author": article_config.get('author', '佚名'),
            "digest": digest,
            "content": final_html_content_for_api,
            "thumb_media_id": actual_thumb_media_id,
            "need_open_comment": need_open_comment,
            "only_fans_can_comment": only_fans_can_comment
        }]
    }
    
    log_message(f"    草稿数据摘要:")
    log_message(f"      标题: {article_title}")
    log_message(f"      作者: {article_config.get('author', '佚名')}")
    log_message(f"      摘要: {digest[:30]}...")
    log_message(f"      内容长度: {len(final_html_content_for_api)} 字符")
    log_message(f"      内容大小: {content_byte_size} 字节")
    log_message(f"      封面图片ID: {actual_thumb_media_id}")
    success = create_draft_api(access_token, articles_data, appid_for_log, proxies=proxies)
    return success

def generate_excel_template_if_not_exists(filename=EXCEL_TEMPLATE_NAME):
    if os.path.exists(filename):
        log_message("配置文件模板 '" + str(filename) + "' 已存在。如需重新生成，请先删除它。")
        return
    if not PANDAS_AVAILABLE:
        log_message("Pandas库不可用，无法生成Excel模板 '" + str(filename) + "'")
        return
    
    template_data = {
        '账号名称': ['示例公众号账号1', '我的测试服务号'],
        'appID': ['wx1234567890abcdef', 'wx0987654321fedcba'],
        'app secret': ['abcdef1234567890abcdef1234567890', 'fedcba0987654321fedcba0987654321'],
        '作者名称': ['示例作者张三', '测试小编'],
        '存稿文件路径': ['/path/to/your/account1/articles', 'C:\\Users\\YourName\\Documents\\Account2Articles'],
        '存稿数量': [2, 1],
        '消息类型': ['图文消息', '图片消息'],
        '是否开始原创': ['是', '否'],
        '是否开启评论': ['是', '否'],
        '评论权限': ['所有人', '仅粉丝'],
        '代理IP': ['', '127.0.0.1'],
        '代理端口': ['', '1080'],
        '代理用户名': ['', 'proxyuser'],
        '代理密码': ['', 'proxypass']
    }
    try:
        import pandas as pd
        df_template = pd.DataFrame(template_data)
        df_template.to_excel(filename, index=False)
        log_message("已生成Excel配置文件模板: '" + str(filename) + "'")
        log_message(f"请根据实际情况修改此文件中的内容，然后重新运行脚本并输入此文件名。")
        log_message("注意：")
        log_message("  - 消息类型：填写'图文消息'或'图片消息'")
        log_message("  - 图文消息：存稿路径中放置.html或.txt文件")
        log_message("  - 图片消息：存稿路径中放置包含图片和txt文件的子文件夹")
    except Exception as e:
        log_message("生成Excel模板 '" + str(filename) + "' 失败: " + str(e))

# ===================== GUI 相关代码 =====================

class ProcessingThread(QThread):
    """处理线程，避免阻塞UI"""
    log_signal = pyqtSignal(str)
    progress_signal = pyqtSignal(int, int)  # current, total
    account_stats_signal = pyqtSignal(str, dict)  # account_name, stats
    finished_signal = pyqtSignal(bool)  # success
    
    def __init__(self, excel_file_path):
        super().__init__()
        self.excel_file_path = excel_file_path
        self.account_stats = {}
        self.stats_manager = StatisticsManager()
        self.processing_start_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
    def emit_log(self, message):
        from datetime import datetime
        timestamp = datetime.now().strftime("[%H:%M:%S] ")
        self.log_signal.emit(timestamp + message)
        
    def run(self):
        try:
            self.process_accounts()
            self.finished_signal.emit(True)
        except Exception as e:
            self.emit_log(f"处理过程中发生错误: {str(e)}")
            self.finished_signal.emit(False)
    
    def process_accounts(self):
        # 设置日志回调
        set_log_callback(self.emit_log)
        
        self.emit_log("开始读取Excel配置文件...")
        
        # 检查pandas是否可用
        if not PANDAS_AVAILABLE:
            self.emit_log("错误: pandas库不可用，无法读取Excel文件")
            return
        
        try:
            import pandas as pd
            df = pd.read_excel(self.excel_file_path, sheet_name=0, dtype=str).fillna('')
            self.emit_log(f"成功读取 {len(df)} 条账号配置")
        except Exception as e:
            self.emit_log(f"读取Excel文件失败: {str(e)}")
            return
            
        required_columns = ['appID', 'app secret', '作者名称', '存稿文件路径', '存稿数量', '消息类型', 
                           '是否开始原创', '是否开启评论', '评论权限', '代理IP', '代理端口', '代理用户名', '代理密码']
        missing_cols = [col for col in required_columns if col not in df.columns]
        if missing_cols:
            self.emit_log(f"Excel文件缺少必需列: {', '.join(missing_cols)}")
            return
            
        total_accounts = len(df)
        
        for index, row in df.iterrows():
            account_name = str(row.get('账号名称', f'账号{index+1}')).strip()
            self.emit_log(f"\n{'='*20} 开始处理 {account_name} {'='*20}")
            
            # 初始化账号统计
            stats = {
                'success_count': 0,
                'fail_count': 0,
                'failed_items': []
            }
            
            try:
                message_type = str(row.get('消息类型', '图文消息')).strip()
                self.process_single_account(row, account_name, stats)
                self.emit_log(f"{account_name} 处理完成: 成功 {stats['success_count']} 个，失败 {stats['fail_count']} 个")
                
                # 保存统计数据到历史记录
                self.stats_manager.add_record(account_name, stats, message_type, self.processing_start_time)
                
            except Exception as e:
                self.emit_log(f"{account_name} 处理时发生错误: {str(e)}")
                stats['fail_count'] += 1
                stats['failed_items'].append(f"账号处理异常: {str(e)}")
                
                # 即使出错也要保存记录
                message_type = str(row.get('消息类型', '图文消息')).strip()
                self.stats_manager.add_record(account_name, stats, message_type, self.processing_start_time)
                
            self.account_stats[account_name] = stats
            self.account_stats_signal.emit(account_name, stats)
            self.progress_signal.emit(index + 1, total_accounts)
            
    def process_single_account(self, row, account_name, stats):
        # 解析配置参数
        appid = str(row['appID']).strip()
        appsecret = str(row['app secret']).strip()
        author_name = str(row['作者名称']).strip()
        articles_folder_path = str(row['存稿文件路径']).strip()
        
        try:
            num_to_publish = int(float(str(row['存稿数量']))) if str(row['存稿数量']).strip() else 0
        except ValueError:
            num_to_publish = 0
            
        message_type = str(row.get('消息类型', '图文消息')).strip()
        
        # 其他配置解析
        is_original_bool = str(row['是否开始原创']).strip().lower() in ['是', 'true', '1', 'yes']
        is_comment_bool = str(row['是否开启评论']).strip().lower() in ['是', 'true', '1', 'yes']
        comment_permission = str(row['评论权限']).strip()
        
        # 代理配置
        current_proxies = self.setup_proxy(row)
        
        self.emit_log(f"账号信息: {author_name}, 类型: {message_type}, 数量: {num_to_publish}")
        
        if not os.path.isdir(articles_folder_path):
            error_msg = f"路径无效: {articles_folder_path}"
            self.emit_log(error_msg)
            stats['fail_count'] += 1
            stats['failed_items'].append(error_msg)
            return 0
            
        # 获取Access Token
        access_token = self.get_access_token_with_log(appid, appsecret, current_proxies, stats)
        if not access_token:
            return 0
            
        if num_to_publish == 0:
            self.emit_log("存稿数量为0，跳过处理")
            return 0
            
        # 根据消息类型处理
        article_config = {
            'appid': appid,
            'author': author_name,
            'is_original': is_original_bool,
            'is_comment_enabled': is_comment_bool,
            'comment_permission': comment_permission,
        }
        
        if message_type == '图片消息':
            return self.process_picture_messages_with_stats(
                articles_folder_path, article_config, access_token, 
                num_to_publish, current_proxies, stats
            )
        else:
            return self.process_text_messages_with_stats(
                articles_folder_path, article_config, access_token,
                num_to_publish, current_proxies, stats
            )
    
    def setup_proxy(self, row):
        """设置代理配置"""
        proxy_ip = str(row.get('代理IP', '')).strip()
        proxy_port = str(row.get('代理端口', '')).strip()
        proxy_user = str(row.get('代理用户名', '')).strip()
        proxy_pass = str(row.get('代理密码', '')).strip()
        
        if proxy_ip and proxy_port:
            try:
                int_proxy_port = int(proxy_port)
                proxy_auth = f"{proxy_user}:{proxy_pass}@" if proxy_user and proxy_pass else ""
                return {
                    "http": f"http://{proxy_auth}{proxy_ip}:{int_proxy_port}",
                    "https": f"http://{proxy_auth}{proxy_ip}:{int_proxy_port}"
                }
            except ValueError:
                self.emit_log(f"代理端口配置错误: {proxy_port}")
        return None
    
    def get_access_token_with_log(self, appid, appsecret, proxies, stats):
        """获取访问令牌并记录统计"""
        try:
            token = get_access_token(appid, appsecret, proxies)
            if token:
                self.emit_log("成功获取Access Token")
                return token
            else:
                error_msg = f"获取Access Token失败"
                self.emit_log(error_msg)
                stats['fail_count'] += 1
                stats['failed_items'].append(error_msg)
                return None
        except Exception as e:
            error_msg = f"获取Access Token异常: {str(e)}"
            self.emit_log(error_msg)
            stats['fail_count'] += 1
            stats['failed_items'].append(error_msg)
            return None
    
    def process_text_messages_with_stats(self, articles_folder_path, article_config, 
                                       access_token, num_to_publish, proxies, stats):
        """处理图文消息并统计结果"""
        try:
            article_files = sorted([f for f in os.listdir(articles_folder_path) 
                                  if os.path.isfile(os.path.join(articles_folder_path, f)) 
                                  and (f.lower().endswith('.html') or f.lower().endswith('.txt'))])
        except Exception as e:
            error_msg = f"读取文章目录失败: {str(e)}"
            self.emit_log(error_msg)
            stats['fail_count'] += 1
            stats['failed_items'].append(error_msg)
            return 0
            
        if not article_files:
            self.emit_log("未找到可处理的文章文件")
            return 0
            
        processed_count = 0
        for i, file_name in enumerate(article_files):
            if processed_count >= num_to_publish:
                self.emit_log(f"已达到存稿上限 ({num_to_publish})")
                break
                
            full_file_path = os.path.join(articles_folder_path, file_name)
            self.emit_log(f"[{i+1}/{len(article_files)}] 开始处理文件: {file_name}")
            self.emit_log("-" * 40)
            
            article_config['html_file_full_path'] = full_file_path
            
            try:
                if process_single_article(article_config, access_token, proxies):
                    stats['success_count'] += 1
                    processed_count += 1
                    self.emit_log(f"✓ {file_name} 处理成功")
                    
                    # 移动文件到已发内容
                    if self.move_processed_file(articles_folder_path, file_name):
                        self.emit_log(f"文件已移动到已发内容文件夹")
                    
                    # 每个成功项目后添加分隔线
                    self.emit_log("=" * 60)
                else:
                    stats['fail_count'] += 1
                    stats['failed_items'].append(file_name)
                    self.emit_log(f"✗ {file_name} 处理失败")
                    self.emit_log("=" * 60)
            except Exception as e:
                error_msg = f"{file_name} 处理异常: {str(e)}"
                self.emit_log(error_msg)
                stats['fail_count'] += 1
                stats['failed_items'].append(error_msg)
                self.emit_log("=" * 60)
                
        return processed_count
    
    def process_picture_messages_with_stats(self, articles_folder_path, article_config, 
                                          access_token, num_to_publish, proxies, stats):
        """处理图片消息并统计结果"""
        try:
            subfolders = [f for f in os.listdir(articles_folder_path) 
                         if os.path.isdir(os.path.join(articles_folder_path, f)) 
                         and f != ARCHIVED_FOLDER_NAME]
            subfolders.sort()
        except Exception as e:
            error_msg = f"读取图片消息目录失败: {str(e)}"
            self.emit_log(error_msg)
            stats['fail_count'] += 1
            stats['failed_items'].append(error_msg)
            return 0
            
        if not subfolders:
            self.emit_log("未找到图片消息子文件夹")
            return 0
            
        processed_count = 0
        for i, subfolder in enumerate(subfolders):
            if processed_count >= num_to_publish:
                self.emit_log(f"已达到存稿上限 ({num_to_publish})")
                break
                
            subfolder_path = os.path.join(articles_folder_path, subfolder)
            self.emit_log(f"[{i+1}/{len(subfolders)}] 开始处理子文件夹: {subfolder}")
            self.emit_log("-" * 40)
            
            try:
                if process_single_picture_folder(subfolder_path, article_config, access_token, proxies):
                    stats['success_count'] += 1
                    processed_count += 1
                    self.emit_log(f"✓ {subfolder} 处理成功")
                    
                    # 移动文件夹到已发内容
                    if self.move_processed_folder(articles_folder_path, subfolder):
                        self.emit_log(f"文件夹已移动到已发内容文件夹")
                    
                    # 每个成功项目后添加分隔线
                    self.emit_log("=" * 60)
                else:
                    stats['fail_count'] += 1
                    stats['failed_items'].append(subfolder)
                    self.emit_log(f"✗ {subfolder} 处理失败")
                    self.emit_log("=" * 60)
            except Exception as e:
                error_msg = f"{subfolder} 处理异常: {str(e)}"
                self.emit_log(error_msg)
                stats['fail_count'] += 1
                stats['failed_items'].append(error_msg)
                self.emit_log("=" * 60)
                
        return processed_count
    
    def move_processed_file(self, articles_folder_path, file_name):
        """移动已处理的文件"""
        try:
            archived_dir = os.path.join(articles_folder_path, ARCHIVED_FOLDER_NAME)
            if not os.path.exists(archived_dir):
                os.makedirs(archived_dir)
                
            source_path = os.path.join(articles_folder_path, file_name)
            destination_path = os.path.join(archived_dir, file_name)
            shutil.move(source_path, destination_path)
            return True
        except Exception as e:
            self.emit_log(f"移动文件失败: {str(e)}")
            return False
    
    def move_processed_folder(self, articles_folder_path, folder_name):
        """移动已处理的文件夹"""
        try:
            archived_dir = os.path.join(articles_folder_path, ARCHIVED_FOLDER_NAME)
            if not os.path.exists(archived_dir):
                os.makedirs(archived_dir)
                
            source_path = os.path.join(articles_folder_path, folder_name)
            destination_path = os.path.join(archived_dir, folder_name)
            shutil.move(source_path, destination_path)
            return True
        except Exception as e:
            self.emit_log(f"移动文件夹失败: {str(e)}")
            return False

class WeChatDraftGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.processing_thread = None
        self.stats_manager = StatisticsManager()
        self.init_ui()
        self.load_historical_data()
        
    def init_ui(self):
        self.setWindowTitle("微信公众号批量存稿工具")
        self.setGeometry(100, 100, 900, 600)
        
        # 设置中央widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 主布局
        main_layout = QVBoxLayout(central_widget)
        
        # 文件选择区域
        self.setup_file_selection(main_layout)
        
        # 创建选项卡
        self.setup_tabs(main_layout)
        
        # 控制按钮
        self.setup_control_buttons(main_layout)
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)
        
        # 状态栏
        self.statusBar().showMessage("就绪")
        
    def setup_file_selection(self, main_layout):
        """设置文件选择区域"""
        file_group = QGroupBox("配置文件选择")
        file_layout = QHBoxLayout(file_group)
        
        self.file_path_label = QLabel("请选择Excel配置文件")
        self.file_path_label.setStyleSheet("QLabel { padding: 5px; border: 1px solid #ccc; }")
        
        self.browse_button = QPushButton("浏览")
        self.browse_button.clicked.connect(self.browse_file)
        
        file_layout.addWidget(self.file_path_label, 1)
        file_layout.addWidget(self.browse_button)
        
        main_layout.addWidget(file_group)
        
    def setup_tabs(self, main_layout):
        """设置选项卡"""
        self.tab_widget = QTabWidget()
        
        # 日志选项卡 - 使用QPlainTextEdit替代QTextEdit避免光标问题
        self.log_text = QPlainTextEdit()
        # 不设置字体，使用系统默认字体
        self.log_text.setReadOnly(True)
        # 设置文本编辑器属性以减少警告
        self.log_text.setLineWrapMode(QPlainTextEdit.LineWrapMode.WidgetWidth)
        # 设置滚动条策略
        self.log_text.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        # 禁用撤销/重做功能以减少光标操作
        self.log_text.setUndoRedoEnabled(False)
        # 允许选择和复制文本，但不允许编辑
        self.log_text.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse | Qt.TextInteractionFlag.TextSelectableByKeyboard)
        # 设置最大块数限制，避免内存问题
        self.log_text.setMaximumBlockCount(1000)
        self.tab_widget.addTab(self.log_text, "处理日志")
        
        # 统计选项卡
        self.stats_table = QTableWidget()
        self.setup_stats_table()
        self.tab_widget.addTab(self.stats_table, "处理统计")
        
        main_layout.addWidget(self.tab_widget)
        
    def setup_stats_table(self):
        """设置统计表格"""
        self.stats_table.setColumnCount(7)
        self.stats_table.setHorizontalHeaderLabels([
            "处理时间", "账号名称", "消息类型", "成功数量", "失败数量", "总处理数", "失败详情"
        ])
        
        # 设置列宽
        header = self.stats_table.horizontalHeader()
        header.setStretchLastSection(True)
        self.stats_table.setColumnWidth(0, 150)  # 处理时间
        self.stats_table.setColumnWidth(1, 150)  # 账号名称
        self.stats_table.setColumnWidth(2, 100)  # 消息类型
        self.stats_table.setColumnWidth(3, 80)   # 成功数量
        self.stats_table.setColumnWidth(4, 80)   # 失败数量
        self.stats_table.setColumnWidth(5, 80)   # 总处理数
        
        # 设置表格属性以减少警告
        self.stats_table.setAlternatingRowColors(True)
        self.stats_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.stats_table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.stats_table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        
    def setup_control_buttons(self, main_layout):
        """设置控制按钮"""
        button_layout = QHBoxLayout()
        
        self.start_button = QPushButton("开始处理")
        self.start_button.clicked.connect(self.start_processing)
        self.start_button.setEnabled(False)
        
        self.stop_button = QPushButton("停止处理")
        self.stop_button.clicked.connect(self.stop_processing)
        self.stop_button.setEnabled(False)
        
        self.clear_log_button = QPushButton("清空日志")
        self.clear_log_button.clicked.connect(self.clear_log)
        
        self.generate_template_button = QPushButton("生成配置模板")
        self.generate_template_button.clicked.connect(self.generate_template)
        
        self.clear_stats_button = QPushButton("清除统计历史")
        self.clear_stats_button.clicked.connect(self.clear_statistics)
        
        self.export_stats_button = QPushButton("导出统计数据")
        self.export_stats_button.clicked.connect(self.export_statistics)
        
        button_layout.addWidget(self.start_button)
        button_layout.addWidget(self.stop_button)
        button_layout.addStretch()
        button_layout.addWidget(self.clear_log_button)
        button_layout.addWidget(self.clear_stats_button)
        button_layout.addWidget(self.export_stats_button)
        button_layout.addWidget(self.generate_template_button)
        
        main_layout.addLayout(button_layout)
        
    def browse_file(self):
        """浏览选择文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择Excel配置文件", "", "Excel文件 (*.xlsx *.xls)"
        )
        if file_path:
            self.file_path_label.setText(file_path)
            self.start_button.setEnabled(True)
            self.log_message(f"已选择配置文件: {file_path}")
            
    def start_processing(self):
        """开始处理"""
        excel_file_path = self.file_path_label.text()
        if not excel_file_path or excel_file_path == "请选择Excel配置文件":
            QMessageBox.warning(self, "警告", "请先选择Excel配置文件")
            return
            
        if not os.path.exists(excel_file_path):
            QMessageBox.warning(self, "警告", "配置文件不存在")
            return
            
        # 重新加载历史数据（不清空，保持历史记录）
        self.load_historical_data()
        
        # 启动处理线程
        self.processing_thread = ProcessingThread(excel_file_path)
        self.processing_thread.log_signal.connect(self.log_message)
        self.processing_thread.progress_signal.connect(self.update_progress)
        self.processing_thread.account_stats_signal.connect(self.update_account_stats)
        self.processing_thread.finished_signal.connect(self.processing_finished)
        
        self.processing_thread.start()
        
        # 更新UI状态
        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(True)
        self.progress_bar.setVisible(True)
        self.statusBar().showMessage("正在处理...")
        
    def stop_processing(self):
        """停止处理"""
        if self.processing_thread and self.processing_thread.isRunning():
            self.processing_thread.terminate()
            self.processing_thread.wait()
            self.log_message("处理已被用户停止")
            self.processing_finished(False)
            
    def processing_finished(self, success):
        """处理完成"""
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        self.progress_bar.setVisible(False)
        
        if success:
            self.statusBar().showMessage("处理完成")
            self.log_message("所有账号处理完成！")
            QMessageBox.information(self, "完成", "所有账号处理完成！")
        else:
            self.statusBar().showMessage("处理失败或被中断")
            
    def update_progress(self, current, total):
        """更新进度条"""
        if total > 0:
            self.progress_bar.setMaximum(total)
            self.progress_bar.setValue(current)
            
    def update_account_stats(self, account_name, stats):
        """更新账号统计"""
        # 重新加载所有历史数据以显示最新记录
        self.load_historical_data()
        
    def log_message(self, message):
        """记录日志消息"""
        # 使用QPlainTextEdit的appendPlainText方法，它不会有光标问题
        self.log_text.appendPlainText(message)
        
        # 自动滚动到底部
        try:
            scrollbar = self.log_text.verticalScrollBar()
            scrollbar.setValue(scrollbar.maximum())
        except Exception:
            pass
        
    def clear_log(self):
        """清空日志"""
        self.log_text.clear()
        
    def generate_template(self):
        """生成配置模板"""
        try:
            generate_excel_template_if_not_exists()
            self.log_message("Excel配置模板生成完成")
            QMessageBox.information(self, "完成", "Excel配置模板生成完成")
        except Exception as e:
            error_msg = f"生成模板失败: {str(e)}"
            self.log_message(error_msg)
            QMessageBox.warning(self, "错误", error_msg)
    
    def load_historical_data(self):
        """加载历史统计数据到表格"""
        try:
            history = self.stats_manager.load_statistics()
            
            # 清空表格
            self.stats_table.setRowCount(0)
            
            # 按时间倒序显示（最新的在上面）
            history_sorted = sorted(history, key=lambda x: x.get('timestamp', ''), reverse=True)
            
            for record in history_sorted:
                row = self.stats_table.rowCount()
                self.stats_table.insertRow(row)
                
                # 按新的列结构填充数据
                self.stats_table.setItem(row, 0, QTableWidgetItem(record.get('timestamp', '')))
                self.stats_table.setItem(row, 1, QTableWidgetItem(record.get('account_name', '')))
                self.stats_table.setItem(row, 2, QTableWidgetItem(record.get('message_type', '')))
                self.stats_table.setItem(row, 3, QTableWidgetItem(str(record.get('success_count', 0))))
                self.stats_table.setItem(row, 4, QTableWidgetItem(str(record.get('fail_count', 0))))
                self.stats_table.setItem(row, 5, QTableWidgetItem(str(record.get('total_processed', 0))))
                
                # 失败详情
                failed_items = record.get('failed_items', [])
                failed_text = '\n'.join(failed_items) if failed_items else "无"
                self.stats_table.setItem(row, 6, QTableWidgetItem(failed_text))
                
        except Exception as e:
            self.log_message(f"加载历史数据失败: {str(e)}")
    
    def clear_statistics(self):
        """清除统计历史"""
        reply = QMessageBox.question(
            self, '确认清除', 
            '确定要清除所有统计历史数据吗？此操作不可撤销。',
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            if self.stats_manager.clear_statistics():
                self.load_historical_data()
                self.log_message("统计历史已清除")
                QMessageBox.information(self, "完成", "统计历史已成功清除")
            else:
                QMessageBox.warning(self, "错误", "清除统计历史失败")
    
    def export_statistics(self):
        """导出统计数据"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "导出统计数据", f"微信存稿统计_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv", 
            "CSV文件 (*.csv)"
        )
        
        if file_path:
            if self.stats_manager.export_to_csv(file_path):
                self.log_message(f"统计数据已导出到: {file_path}")
                QMessageBox.information(self, "完成", f"统计数据已成功导出到:\n{file_path}")
            else:
                QMessageBox.warning(self, "错误", "导出统计数据失败")

def run_gui():
    """运行GUI模式"""
    if not PYQT6_AVAILABLE:
        log_message("PyQt6库不可用，无法启动GUI模式。请安装: pip install PyQt6")
        return False
        
    import sys
    import os
    
    # 抑制Qt的一些警告输出
    os.environ['QT_LOGGING_RULES'] = '*.debug=false;qt.qpa.fonts=false;qt.text.font.db=false;*.warning=false'
    os.environ['QT_ASSUME_STDERR_HAS_CONSOLE'] = '1'
    # 抑制CSS和光标相关的警告
    os.environ['QT_QUICK_CONTROLS_STYLE'] = 'Basic'
    # 完全禁用Qt日志输出
    os.environ['QT_LOGGING_RULES'] = '*=false'
    
    app = QApplication(sys.argv)
    app.setApplicationName("微信公众号批量存稿工具")
    
    # 设置应用程序属性
    app.setAttribute(Qt.ApplicationAttribute.AA_DontShowIconsInMenus, False)
    
    window = WeChatDraftGUI()
    window.show()
    
    sys.exit(app.exec())

# ===================== 主程序入口 =====================

def main():
    """主程序入口，直接启动GUI模式"""
    if not PYQT6_AVAILABLE:
        log_message("错误：PyQt6库不可用，无法启动GUI界面。")
        log_message("请安装PyQt6: pip install PyQt6")
        log_message("然后重新运行程序。")
        input("按回车键退出...")
        return
    
    # 直接启动GUI
    run_gui()

# 命令行模式已删除，程序现在只支持GUI模式

if __name__ == "__main__":
    main() 