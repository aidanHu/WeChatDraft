import requests
import json
import re
import os
import shutil

try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except ImportError:
    PANDAS_AVAILABLE = False
    print("警告: pandas 库未找到。无法从Excel读取配置或生成模板。")
    print("请尝试运行 'pip install pandas openpyxl' 来安装它以启用此功能。")
try:
    from premailer import Premailer
    PREMAILER_AVAILABLE = True
except ImportError: PREMAILER_AVAILABLE = False
try:
    from bs4 import BeautifulSoup
    BS4_AVAILABLE = True
except ImportError: BS4_AVAILABLE = False
try:
    import socks
    SOCKS_AVAILABLE = True
except ImportError: SOCKS_AVAILABLE = False

# --- 全局配置 (这些可以被Excel中的数据覆盖) ---
BASE_URL = "https://api.weixin.qq.com/cgi-bin"
WECHAT_IMG_DOMAINS = ["mmbiz.qlogo.cn", "mmbiz.qpic.cn"] 
ARCHIVED_FOLDER_NAME = "已发内容" # 移动已处理文件的子文件夹名
EXCEL_TEMPLATE_NAME = "wechat_config_template.xlsx"
# --- 全局配置结束 ---

def _make_request(method, url, **kwargs):
    """统一处理 requests 请求，加入 proxies 参数"""
    # proxies 参数应形如: {'http': 'socks5h://user:pass@host:port', 'https': 'socks5h://user:pass@host:port'}
    # 或者 {'http': 'socks5h://host:port', 'https': 'socks5h://host:port'}
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
            print("  获取access_token失败 (AppID: " + str(appid) + "): " + str(data))
            return None
    except requests.exceptions.RequestException as e:
        print("  请求access_token时发生错误 (AppID: " + str(appid) + "): " + str(e))
        return None
    except json.JSONDecodeError as e:
        print("  解析access_token响应JSON时出错 (AppID: " + str(appid) + "): " + str(e))
        return None

def download_image_from_url(image_url, local_filename, proxies=None):
    try:
        print("    下载图片从: " + str(image_url) + " -> " + str(local_filename))
        response = _make_request("get", image_url, stream=True, proxies=proxies)
        response.raise_for_status()
        with open(local_filename, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)
        return True
    except requests.exceptions.RequestException as e:
        print("    下载图片失败 (" + str(image_url) + "): " + str(e))
        return False
    except IOError as e:
        print("    保存图片失败 (" + str(local_filename) + "): " + str(e))
        return False

def upload_permanent_material(access_token, file_path, material_type='image', appid_for_log="", proxies=None):
    url = f"{BASE_URL}/material/add_material?access_token={access_token}&type={material_type}"
    log_prefix = f"(AppID: {appid_for_log}) " if appid_for_log else ""
    response = None 
    try:
        print("    " + log_prefix + "上传本地素材 " + str(file_path) + " (类型: " + str(material_type) + ")...")
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
            print("    " + log_prefix + "素材上传成功！Media ID: " + str(wx_media_id) + (", URL: " + str(wx_image_url) if wx_image_url else ""))
            return {"media_id": wx_media_id, "url": wx_image_url}
        else:
            print("    " + log_prefix + "上传永久素材失败: " + str(result))
            return None
    except requests.exceptions.RequestException as e:
        print("    " + log_prefix + "请求上传永久素材错误: " + str(e))
        return None
    except IOError as e:
        print("    " + log_prefix + "读取本地文件错误 " + str(file_path) + ": " + str(e))
        return None
    except json.JSONDecodeError:
        response_text = response.text if response is not None and hasattr(response, 'text') else 'No response object or text attribute'
        print("    " + log_prefix + "无法解析上传素材响应: " + str(response_text))
        return None

def optimize_html_with_inline_styles(html_string):
    if not PREMAILER_AVAILABLE:
        print("    Premailer 库不可用，跳过HTML样式内联优化。")
        return html_string
    # print("    开始 Premailer HTML样式内联优化...") # 减少重复打印
    try:
        p = Premailer(html_string, remove_classes=False, keep_style_tags=True, strip_important=False)
        inlined_html = p.transform()
        # print("    Premailer HTML样式内联优化完成。")
        return inlined_html
    except Exception as e:
        print("    Premailer优化错误: " + str(e) + "。使用原始HTML。")
        return html_string

def replace_external_images_in_html(html_content, access_token, appid_for_log="", current_html_file_path="", proxies=None):
    if not BS4_AVAILABLE:
        print("    BeautifulSoup4 库不可用，跳过正文图片链接替换。")
        return html_content
    if not access_token:
        # print("Access Token无效，无法替换正文图片链接。") # 由调用者处理
        return html_content

    # print("    开始处理HTML中的外部图片链接...") # 减少重复打印
    try:
        soup = BeautifulSoup(html_content, 'lxml')
    except Exception as e:
        print("    BeautifulSoup解析HTML失败: " + str(e) + "。跳过图片替换。")
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
            print("      处理第" + str(image_counter) + "个外部图片: " + original_src[:70] + ('...' if len(original_src)>70 else ''))
            # 为每个账号和文章图片生成更独特的临时文件名
            base_html_filename = os.path.splitext(os.path.basename(current_html_file_path))[0]
            temp_img_filename = f"temp_body_img_{appid_for_log.replace('.', '_')}_{base_html_filename}_{i}.jpg"
            
            if download_image_from_url(original_src, temp_img_filename, proxies=proxies):
                upload_result = upload_permanent_material(access_token, temp_img_filename, 'image', appid_for_log, proxies=proxies)
                if upload_result and upload_result.get("url"):
                    wx_image_url = upload_result["url"]
                    img['src'] = wx_image_url 
                    print("        成功替换为微信图片URL: " + str(wx_image_url))
                    processed_image_count +=1
                else:
                    print("        上传失败或未返回URL，保留原始src.")
                if os.path.exists(temp_img_filename):
                    try: os.remove(temp_img_filename)
                    except OSError as e: print("        删除临时文件 " + str(temp_img_filename) + " 失败: " + str(e))
            else:
                print("        下载失败，保留原始src.")
        # elif is_wechat_domain:
            # print(f"  跳过微信内部图片链接: {original_src}")

    if image_counter > 0:
        print("    共找到" + str(image_counter) + "个外部图片链接，成功处理了" + str(processed_image_count) + "个。")
    # else:
        # print("    未在HTML中找到需要处理的外部图片链接。")
    return str(soup)

def create_draft_api(access_token, articles_data, appid_for_log="", proxies=None):
    url = f"{BASE_URL}/draft/add?access_token={access_token}"
    headers = {"Content-Type": "application/json"}
    log_prefix = f"(AppID: {appid_for_log}) " if appid_for_log else ""
    try:
        response = _make_request("post", url, headers=headers, data=json.dumps(articles_data, ensure_ascii=False).encode('utf-8'), proxies=proxies)
        response.raise_for_status()
        result = response.json()
        if result.get("errcode") == 0 or "media_id" in result :
            print("  " + log_prefix + "草稿API响应成功: " + str(result)) 
            return result.get("media_id") if "media_id" in result else True
        else:
            print("  " + log_prefix + "创建草稿失败: " + str(result))
            return None
    except requests.exceptions.RequestException as e:
        print("  " + log_prefix + "请求创建草稿时发生错误: " + str(e))
        return None
    except json.JSONDecodeError:
        response_text = response.text if 'response' in locals() and hasattr(response, 'text') else 'No response text available'
        print("  " + log_prefix + "无法解析草稿创建响应: " + str(response_text))
        return None

def process_single_article(article_config, access_token, proxies=None):
    print("  处理文章: " + str(article_config['html_file_full_path']))
    appid_for_log = article_config.get('appid', 'N/A')
    current_html_file_path = article_config['html_file_full_path']
    raw_html_content = ""
    try:
        with open(current_html_file_path, 'r', encoding='utf-8') as f: raw_html_content = f.read()
    except FileNotFoundError: 
        print("    错误：找不到文件 " + str(current_html_file_path) + "。跳过。") 
        return False
    except Exception as e: 
        print("    读取文件 " + str(current_html_file_path) + " 错误: " + str(e) + "。跳过。")
        return False

    print("    步骤1: Premailer CSS内联优化...")
    optimized_html_content = optimize_html_with_inline_styles(raw_html_content)
    print("    步骤2: 替换正文外部图片链接...")
    html_with_wechat_images = replace_external_images_in_html(optimized_html_content, access_token, appid_for_log, current_html_file_path, proxies=proxies)
    print("    步骤3: 正则清理HTML...")
    cleaned_html = html_with_wechat_images
    cleaned_html = re.sub(r'<p\b[^>]*>\s*(?:&nbsp;|<br\s*/?>|\s)*\s*</p>', '', cleaned_html, flags=re.IGNORECASE | re.DOTALL)
    cleaned_html = re.sub(r'<p\b[^>]*>\s*<span\b[^>]*>\s*<br\b[^>]*>\s*</span>\s*</p>', '', cleaned_html, flags=re.IGNORECASE | re.DOTALL) 
    cleaned_html = re.sub(r'(\s*<br\s*/?>\s*){2,}', '<br>\n', cleaned_html, flags=re.IGNORECASE)
    cleaned_html = re.sub(r'>\s+<', '><', cleaned_html)
    final_html_content_for_api = cleaned_html
    print("    步骤4: 准备封面图...")
    cover_image_url_from_html = None
    match_cover = re.search(r'<img [^>]*src="([^"]+)"', raw_html_content, re.IGNORECASE)
    if match_cover: cover_image_url_from_html = match_cover.group(1)
    actual_thumb_media_id = None
    if cover_image_url_from_html:
        base_cover_html_filename = os.path.splitext(os.path.basename(current_html_file_path))[0]
        temp_cover_filename = f"temp_cover_{appid_for_log.replace('.', '_')}_{base_cover_html_filename}.jpg"
        if download_image_from_url(cover_image_url_from_html, temp_cover_filename, proxies=proxies):
            upload_result = upload_permanent_material(access_token, temp_cover_filename, 'image', appid_for_log, proxies=proxies)
            if upload_result and upload_result.get("media_id"): actual_thumb_media_id = upload_result["media_id"]
            if os.path.exists(temp_cover_filename): 
                try: os.remove(temp_cover_filename)
                except OSError as e: 
                    print("    删除临时封面文件 " + str(temp_cover_filename) + " 失败: " + str(e))
    if not actual_thumb_media_id: print("    警告: 未能准备封面图。API可能拒绝无封面草稿。")
    
    is_original_for_api = 1 if article_config.get('is_original', False) else 0

    is_comment_enabled = article_config.get('is_comment_enabled', False)
    comment_permission = article_config.get('comment_permission', '所有人')
    need_open_comment = int(1 if is_comment_enabled else 0)
    only_fans_can_comment = int(1 if comment_permission == '仅粉丝' else 0)

    print("    步骤5: 创建草稿...")
    article_title = os.path.splitext(os.path.basename(current_html_file_path))[0]
    articles_data = {
        "articles": [{
            "article_type": "news",  # 明确指定为图文消息
            "title": article_title,
            "author": article_config.get('author', '佚名'),
            "content": final_html_content_for_api,
            "thumb_media_id": actual_thumb_media_id,
            "need_open_comment": need_open_comment,
            "only_fans_can_comment": only_fans_can_comment,
            "is_original": is_original_for_api
        }]
    }
    success = create_draft_api(access_token, articles_data, appid_for_log, proxies=proxies)
    return success

def generate_excel_template_if_not_exists(filename=EXCEL_TEMPLATE_NAME):
    if os.path.exists(filename):
        print("配置文件模板 '" + str(filename) + "' 已存在。如需重新生成，请先删除它。")
        return
    if not PANDAS_AVAILABLE:
        print("Pandas库不可用，无法生成Excel模板 '" + str(filename) + "'")
        return
    
    template_data = {
        '账号名称': ['示例公众号账号1', '我的测试服务号'],
        'appID': ['wx1234567890abcdef', 'wx0987654321fedcba'],
        'app secret': ['abcdef1234567890abcdef1234567890', 'fedcba0987654321fedcba0987654321'],
        '作者名称': ['示例作者张三', '测试小编'],
        '存稿文件路径': ['/path/to/your/account1/articles', 'C:\\Users\\YourName\\Documents\\Account2Articles'],
        '存稿数量': [2, 1],
        '是否开始原创': ['是', '否'],
        '是否开启评论': ['是', '否'],
        '评论权限': ['所有人', '仅粉丝'],
        '代理IP': ['', '127.0.0.1'],
        '代理端口': ['', '1080'],
        '代理用户名': ['', 'proxyuser'],
        '代理密码': ['', 'proxypass']
    }
    df_template = pd.DataFrame(template_data)
    try:
        df_template.to_excel(filename, index=False)
        print("已生成Excel配置文件模板: '" + str(filename) + "'")
        print(f"请根据实际情况修改此文件中的内容，然后重新运行脚本并输入此文件名。")
    except Exception as e:
        print("生成Excel模板 '" + str(filename) + "' 失败: " + str(e))

def main():
    print("微信公众号批量存稿脚本启动...")
    if not PANDAS_AVAILABLE: print("Pandas库缺失，无法读取Excel配置。请安装: pip install pandas openpyxl"); return
    if not PREMAILER_AVAILABLE: print("Premailer库缺失，HTML样式内联将跳过。建议安装: pip install premailer")
    if not BS4_AVAILABLE: print("BeautifulSoup4库缺失，正文图片处理将跳过。建议安装: pip install beautifulsoup4 lxml")
    if not SOCKS_AVAILABLE: print("SOCKS支持库缺失，代理功能可能无法使用。建议安装: pip install \"requests[socks]\" Pysocks")

    excel_file_path = input("请输入Excel配置文件的路径: ").strip()
    # 不再询问 sheet_name，默认读取第一个工作表
    # excel_sheet_name = input("请输入Excel中的工作表名称 (默认为 'Sheet1'): ").strip()
    # if not excel_sheet_name: excel_sheet_name = 'Sheet1'

    try:
        # pd.read_excel 若 sheet_name=0 则读取第一个sheet
        df = pd.read_excel(excel_file_path, sheet_name=0, dtype=str).fillna('') 
        print("\n成功从 '" + str(excel_file_path) + "' (默认读取第一个工作表) 读取 " + str(len(df)) + " 条账号配置。")
    except FileNotFoundError: 
        print("错误：找不到Excel文件 '" + str(excel_file_path) + "'")
        return
    except Exception as e: 
        print("读取Excel文件时发生错误: " + str(e))
        return

    required_columns = ['appID', 'app secret', '作者名称', '存稿文件路径', '存稿数量', '是否开始原创',
                        '是否开启评论', '评论权限', '代理IP', '代理端口', '代理用户名', '代理密码'] 
    missing_cols = [col for col in required_columns if col not in df.columns]
    if missing_cols:
        print("错误：Excel文件中缺少以下列: " + ", ".join(missing_cols) + "。请检查Excel列名。")
        print("提示：必需列名包括: " + ", ".join(required_columns) + " (建议完全复制)")
        return

    for index, row in df.iterrows():
        account_name = str(row.get('账号名称', f'账号{index+1}')).strip()
        appid = str(row['appID']).strip()
        appsecret = str(row['app secret']).strip()
        author_name = str(row['作者名称']).strip()
        articles_folder_path = str(row['存稿文件路径']).strip()
        num_to_publish = 0
        try:
            num_to_publish_val = row['存稿数量']
            if pd.isna(num_to_publish_val) or str(num_to_publish_val).strip() == '' :
                 print("  警告: 账号 '" + str(account_name) + "' 的存稿数量为空，将视为0处理。")
                 num_to_publish = 0
            else:
                num_to_publish = int(float(str(num_to_publish_val)))
                if num_to_publish < 0: 
                    print("  警告: 账号 '" + str(account_name) + "' 的存稿数量配置为负数，将视为0处理。")
                    num_to_publish = 0
        except ValueError:
            print("  警告: 账号 '" + str(account_name) + "' 的存稿数量 '" + str(row['存稿数量']) + "' 不是有效数字，将视为0处理。")
            num_to_publish = 0
            
        is_original_str_from_excel = str(row['是否开始原创']).strip().lower()
        is_original_bool = is_original_str_from_excel in ['是', 'true', '1', 'yes']

        comment_str_from_excel = str(row['是否开启评论']).strip().lower()
        is_comment_bool = comment_str_from_excel in ['是', 'true', '1', 'yes']

        comment_permission = str(row['评论权限']).strip()

        proxy_ip = str(row.get('代理IP', '')).strip()
        proxy_port = str(row.get('代理端口', '')).strip()
        proxy_user = str(row.get('代理用户名', '')).strip()
        proxy_pass = str(row.get('代理密码', '')).strip()
        current_proxies = None
        if proxy_ip and proxy_port:
            if not SOCKS_AVAILABLE:
                print("  警告：SOCKS库未加载，无法为此账号启用代理。请安装 \"requests[socks]\"。")
            else:
                try:
                    int_proxy_port = int(proxy_port)
                    proxy_auth = ""
                    if proxy_user and proxy_pass:
                        proxy_auth = str(proxy_user) + ":" + str(proxy_pass) + "@"
                    current_proxies = { 
                        "http": "socks5h://" + proxy_auth + str(proxy_ip) + ":" + str(int_proxy_port),
                        "https": "socks5h://" + proxy_auth + str(proxy_ip) + ":" + str(int_proxy_port)
                    }
                    print("  已为此账号配置SOCKS5代理: " + str(proxy_ip) + ":" + str(int_proxy_port) + (" (带认证)" if proxy_user else ""))
                except ValueError:
                    print("  警告: 代理端口 '" + str(proxy_port) + "' 不是有效数字，此账号代理将不被使用。")
        
        print("\n================== " + str(account_name) + " (AppID: " + str(appid) + ") ==================")
        print("  作者: " + str(author_name) + ", 原创: " + str(is_original_bool) + ", 评论: " + str(is_comment_bool) + ", 代理: " + ('启用' if current_proxies else '禁用'))
        print("  路径: " + str(articles_folder_path) + ", 处理数量: " + str(num_to_publish))
        print("====================================================================")

        if not os.path.isdir(articles_folder_path): 
            print("  错误：路径 '" + str(articles_folder_path) + "' 无效。跳过。") 
            continue
        access_token = get_access_token(appid, appsecret, proxies=current_proxies)
        if not access_token: 
            print("  无法获取 " + str(account_name) + " 的Token。跳过。") 
            continue
        print("  成功获取 " + str(account_name) + " 的Token。")

        try: article_files = sorted([f for f in os.listdir(articles_folder_path) if os.path.isfile(os.path.join(articles_folder_path, f)) and (f.lower().endswith('.html') or f.lower().endswith('.txt'))])
        except Exception as e: 
            print("  错误：读取存稿目录 '" + str(articles_folder_path) + "' 失败: " + str(e) + "。跳过。") 
            continue
        if not article_files: 
            print("  目录 '" + str(articles_folder_path) + "' 无 .html/.txt 文件。") 
            continue
        if num_to_publish == 0: 
            print("  存稿数量为0，跳过文件处理。") 
            continue
            
        articles_processed_count = 0
        for i, file_name in enumerate(article_files):
            if articles_processed_count >= num_to_publish: 
                print("  已达存稿上限 (" + str(num_to_publish) + ")。") 
                break
            full_file_path = os.path.join(articles_folder_path, file_name)
            print("\n  [" + str(i+1) + "/" + str(len(article_files)) + "] 处理文件: " + str(file_name))

            article_config_data = {
                'appid': appid, 
                'author': author_name, 
                'is_original': is_original_bool,
                'is_comment_enabled': is_comment_bool,
                'comment_permission': comment_permission,
                'content_source_url': "", 
                'html_file_full_path': full_file_path
            }
            if process_single_article(article_config_data, access_token, proxies=current_proxies):
                print("    文章 '" + str(file_name) + "' API调用成功。")
                archived_dir = os.path.join(articles_folder_path, ARCHIVED_FOLDER_NAME)
                if not os.path.exists(archived_dir): 
                    try: 
                        os.makedirs(archived_dir)
                        print("    创建文件夹: " + str(archived_dir))
                    except OSError as e: 
                        error_message = "    错误：创建" + "“已发内容”" + "文件夹 '" + str(archived_dir) + "' 失败: " + str(e) + "。文件不移动。"
                        print(error_message)
                        continue
                destination_path = os.path.join(archived_dir, file_name)
                print("    准备移动: '" + str(full_file_path) + "' -> '" + str(destination_path) + "'")
                try: 
                    shutil.move(full_file_path, destination_path)
                    print("    文件已移动.")
                except Exception as e: 
                    print("    移动文件失败: " + str(e))
                articles_processed_count += 1
            else: 
                print("    处理或创建草稿失败: " + str(file_name))
        
        if articles_processed_count == 0 and num_to_publish > 0: 
            print("  警告: " + str(account_name) + " 成功处理0篇 (目标: " + str(num_to_publish) + ").")
        elif articles_processed_count < num_to_publish: 
            print("  注意: " + str(account_name) + " 成功处理 " + str(articles_processed_count) + "/" + str(num_to_publish) + " 篇.")

    print("\n所有账号处理完毕.")

if __name__ == "__main__":
    main() 