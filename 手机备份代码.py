import os
import re
from datetime import datetime
import win32com.client
import time
import logging
import sys

# 配置日志记录
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("photo_backup.log", encoding='utf-8'),
        logging.StreamHandler()
    ]
)

# 配置路径
backup_dir = r"G:\intern pan\vivo sync\iQOO 12+b6568e\图片\新备份"
videos_dir = os.path.join(backup_dir, "视频")

# 创建必要的目录
os.makedirs(backup_dir, exist_ok=True)
os.makedirs(videos_dir, exist_ok=True)

# 设置起始日期 (2025-08-30)
start_date = datetime(2025, 8, 25)

def get_file_date(filename):
    """从文件名中提取日期信息"""
    date_match = re.search(r'(\d{4})(\d{2})(\d{2})', filename)
    if date_match:
        year, month, day = map(int, date_match.groups())
        try:
            return datetime(year, month, day)
        except ValueError:
            return None
    return None

def should_backup(filename):
    """检查文件是否符合备份条件"""
    file_date = get_file_date(filename)
    if not file_date or file_date < start_date:
        return False
    return filename.startswith(("IMG_", "video_"))

def safe_get_mod_time(item):
    """安全获取文件修改时间"""
    try:
        return time.mktime(item.ModifyDate.timetuple())
    except (OverflowError, ValueError, AttributeError):
        return time.time()

def get_item_size(item):
    """安全获取文件大小"""
    try:
        return int(item.ExtendedProperty("Size"))
    except:
        try:
            return item.Size
        except:
            return 0

def backup_file_shell(item, dest_path, source_size, context=""):
    """只保留方案2 (Shell CopyHere) 的备份逻辑"""
    filename = os.path.basename(dest_path)
    
    # 校验是否已存在完整文件
    if os.path.exists(dest_path):
        if os.path.getsize(dest_path) == source_size:
            logging.debug(f"跳过已存在的文件: {filename}")
            return True
        else:
            # 文件存在但大小不符，清理掉准备重新复制
            try:
                os.remove(dest_path)
            except Exception:
                pass

    start_time = time.time()
    success = False
    
    try:
        shell = win32com.client.Dispatch("Shell.Application")
        dest_folder = shell.NameSpace(os.path.dirname(dest_path))
        
        # 16 = FOF_NOCONFIRMATION (不提示确认), 1024 = FOF_NOERRORUI (不显示错误UI)
        dest_folder.CopyHere(item, 16 | 1024)
        
        # 轮询等待文件传输完成
        timeout = 120
        while time.time() - start_time < timeout:
            if os.path.exists(dest_path):
                try:
                    # Windows MTP 复制时，文件锁定可能导致 getsize 报错，用 try 保护
                    current_size = os.path.getsize(dest_path)
                    if current_size == source_size:
                        success = True
                        break
                except PermissionError:
                    pass # 文件正被锁定写入中，继续等待
            time.sleep(0.5)
            
        if success:
            elapsed = time.time() - start_time
            speed = (source_size / 1024 / 1024) / elapsed if elapsed > 0 else 0
            logging.info(f"成功: {filename} ({context}) - 速度: {speed:.1f} MB/s")
        else:
            logging.warning(f"超时或失败: {filename} 未能完整复制")
            
    except Exception as e:
        logging.error(f"复制失败 [{filename}]: {str(e)}")
        
    return success

def print_progress(current, total, start_time):
    """打印进度条"""
    if total == 0:
        return
    progress = current / total
    percent = progress * 100
    elapsed = time.time() - start_time
    
    if current > 0:
        remaining = (total - current) * (elapsed / current)
        time_str = f"剩余: {int(remaining // 60)}分{int(remaining % 60)}秒"
    else:
        time_str = "剩余: 计算中..."
        
    bar_length = 30
    filled = int(bar_length * progress)
    bar = '█' * filled + '-' * (bar_length - filled)
    
    sys.stdout.write(f"\r进度: |{bar}| {percent:.1f}% ({current}/{total}) {time_str}")
    sys.stdout.flush()

def main():
    # 初始化 COM
    shell = win32com.client.Dispatch("Shell.Application")
    namespace = shell.NameSpace(17) # ssfDRIVES

    # 定位手机
    phone_device = next((item for item in namespace.Items() if item.Name == "iQOO 13"), None)
    if not phone_device:
        logging.error("未找到手机设备 'iQOO 12'")
        return

    # 定位相机目录
    try:
        internal_storage = phone_device.GetFolder.Items().Item(0)
        camera_folder = internal_storage.GetFolder.Items().Item("DCIM").GetFolder.Items().Item("Camera")
    except Exception as e:
        logging.error(f"无法定位到 Camera 目录: {e}")
        return

    # ---------------- 核心优化点：单次 O(N) 遍历建立内存索引 ----------------
    logging.info("开始扫描手机文件，建立索引 (仅遍历一次)...")
    start_collect = time.time()
    
    tasks = []       # 存储所有待备份文件的元数据
    group_dict = {}  # 专门用于存储照片组的字典
    
    items = camera_folder.GetFolder.Items()
    total_items = items.Count
    
    for idx in range(total_items):
        try:
            item = items.Item(idx)
            filename = item.Name
            
            if should_backup(filename):
                file_size = get_item_size(item)
                mod_time = safe_get_mod_time(item)
                
                # 视频独立作为一个任务
                if filename.startswith("video_"):
                    tasks.append({
                        'type': 'video', 'time': mod_time, 'item': item, 
                        'name': filename, 'size': file_size, 'dest': videos_dir, 'context': '视频'
                    })
                
                # 照片根据前缀进行分组归类
                elif filename.startswith("IMG_"):
                    prefix = os.path.splitext(filename)[0]
                    if prefix not in group_dict:
                        group_dict[prefix] = []
                        
                    group_dict[prefix].append({
                        'type': 'image', 'time': mod_time, 'item': item, 
                        'name': filename, 'size': file_size, 'context': f'分组[{prefix}]'
                    })
                    
            if idx % 100 == 0:
                sys.stdout.write(f"\r扫描进度: {idx}/{total_items} ({(idx/total_items)*100:.1f}%)")
                sys.stdout.flush()
                
        except Exception as e:
            pass # 忽略部分无法读取的文件

    # 将组信息展开追加到 tasks 列表中，并创建对应的目录路径
    for prefix, files in group_dict.items():
        group_dest = os.path.join(backup_dir, prefix)
        for f in files:
            f['dest'] = group_dest
            tasks.append(f)

    # 按文件修改时间从新到旧排序
    tasks.sort(key=lambda x: x['time'], reverse=True)
    
    total_files = len(tasks)
    elapsed_collect = time.time() - start_collect
    logging.info(f"\n扫描完毕！找到 {total_files} 个需备份文件，耗时: {elapsed_collect:.1f}秒")

    if total_files == 0:
        return

    # ---------------- 执行备份 ----------------
    processed_count = 0
    start_time = time.time()
    print_progress(0, total_files, start_time)

    for task in tasks:
        os.makedirs(task['dest'], exist_ok=True)
        dest_file_path = os.path.join(task['dest'], task['name'])
        
        success = backup_file_shell(task['item'], dest_file_path, task['size'], task['context'])
        
        processed_count += 1
        print_progress(processed_count, total_files, start_time)

    print("\n")
    elapsed_total = time.time() - start_time
    logging.info(f"所有备份完成！总耗时: {int(elapsed_total // 60)}分{int(elapsed_total % 60)}秒")

if __name__ == "__main__":
    main()