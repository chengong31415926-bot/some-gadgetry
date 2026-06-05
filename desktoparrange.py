import os
import shutil
import logging
from datetime import datetime

# ========== 配置 ==========
# 使用原始字符串避免转义问题
TARGET_DIR = r"G:\Desktop"
# 不参与整理的扩展名（系统文件、临时文件等）
EXCLUDE_EXT = {"url", "lnk", "exe", "ini", "tmp", "reg", "log"}
# 是否递归处理子文件夹（True 则处理所有子目录下的文件）
RECURSIVE = False
# 同名文件冲突时的处理方式: 'rename' 或 'skip' 或 'overwrite'
CONFLICT_ACTION = "rename"
# 是否启用日志
LOG_FILE = "file_organizer.log"
# =========================

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler()
    ]
)

def get_unique_filename(dest_folder, filename):
    """如果目标文件夹已存在同名文件，生成一个不重复的新文件名"""
    name, ext = os.path.splitext(filename)
    counter = 1
    new_name = filename
    while os.path.exists(os.path.join(dest_folder, new_name)):
        new_name = f"{name} ({counter}){ext}"
        counter += 1
    return new_name

def organize_file(file_path, dest_base_dir):
    """
    将单个文件移动到以扩展名命名的子文件夹中
    返回: (success, message)
    """
    if not os.path.isfile(file_path):
        return False, f"跳过非文件: {file_path}"
    
    filename = os.path.basename(file_path)
    ext = os.path.splitext(filename)[1][1:]  # 去掉点号，获取扩展名
    ext_lower = ext.lower()
    
    # 跳过排除的扩展名
    if ext_lower in EXCLUDE_EXT:
        logging.debug(f"跳过排除类型: {filename}")
        return False, "跳过排除扩展名"
    
    # 目标文件夹路径
    dest_folder = os.path.join(dest_base_dir, ext_lower)
    dest_path = os.path.join(dest_folder, filename)
    
    # 创建目标文件夹
    try:
        os.makedirs(dest_folder, exist_ok=True)
    except Exception as e:
        return False, f"无法创建文件夹 {dest_folder}: {e}"
    
    # 处理目标文件已存在的情况
    if os.path.exists(dest_path):
        if CONFLICT_ACTION == "skip":
            return False, f"目标已存在，跳过: {filename}"
        elif CONFLICT_ACTION == "rename":
            new_filename = get_unique_filename(dest_folder, filename)
            dest_path = os.path.join(dest_folder, new_filename)
            logging.info(f"文件重命名: {filename} -> {new_filename}")
        # 如果为 "overwrite"，则直接覆盖（不处理）
    
    # 移动文件
    try:
        shutil.move(file_path, dest_path)
        return True, f"已移动: {filename} -> {os.path.relpath(dest_path, TARGET_DIR)}"
    except Exception as e:
        return False, f"移动失败 {filename}: {e}"

def organize_directory(root_dir, recursive=False):
    """整理指定目录下的所有文件"""
    success_count = 0
    skip_count = 0
    fail_count = 0
    
    if recursive:
        # 递归遍历所有子文件夹
        for dirpath, dirnames, filenames in os.walk(root_dir):
            # 跳过已经整理过的目标文件夹（避免无限循环）
            # 注意：这里简单处理，不移动已经在“扩展名”文件夹中的文件
            # 判断当前目录本身是否已经是类似“jpg”的扩展名文件夹
            if os.path.basename(dirpath) in EXCLUDE_EXT:
                continue
            # 实际上，更安全的做法是：不处理以扩展名命名的文件夹下的文件
            # 但为了简化，这里不做深度限制，用户应确保不重复整理同一目录
            for filename in filenames:
                file_path = os.path.join(dirpath, filename)
                success, msg = organize_file(file_path, TARGET_DIR)
                if success:
                    success_count += 1
                    logging.info(msg)
                elif "跳过" in msg:
                    skip_count += 1
                    logging.debug(msg)
                else:
                    fail_count += 1
                    logging.error(msg)
    else:
        # 仅处理顶层目录下的文件
        for item in os.listdir(root_dir):
            file_path = os.path.join(root_dir, item)
            if not os.path.isfile(file_path):
                continue
            success, msg = organize_file(file_path, TARGET_DIR)
            if success:
                success_count += 1
                logging.info(msg)
            elif "跳过" in msg:
                skip_count += 1
                logging.debug(msg)
            else:
                fail_count += 1
                logging.error(msg)
    
    return success_count, skip_count, fail_count

def main():
    # 检查目标目录是否存在
    if not os.path.exists(TARGET_DIR):
        logging.error(f"目标目录不存在: {TARGET_DIR}")
        return
    
    logging.info(f"开始整理目录: {TARGET_DIR}")
    logging.info(f"递归模式: {RECURSIVE}")
    logging.info(f"冲突处理策略: {CONFLICT_ACTION}")
    
    start_time = datetime.now()
    success, skip, fail = organize_directory(TARGET_DIR, RECURSIVE)
    elapsed = (datetime.now() - start_time).total_seconds()
    
    logging.info("=" * 50)
    logging.info(f"整理完成！耗时 {elapsed:.2f} 秒")
    logging.info(f"成功移动: {success} 个文件")
    logging.info(f"跳过: {skip} 个文件")
    logging.info(f"失败: {fail} 个文件")
    logging.info("=" * 50)

if __name__ == "__main__":
    main()