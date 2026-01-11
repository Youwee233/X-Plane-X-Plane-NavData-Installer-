import os
import sys
import shutil
import subprocess
import tempfile
import configparser
import time

# ================= 1. 配置文件管理 =================
CONFIG_FILE = "rules_config.ini"


def load_config():
    config = configparser.ConfigParser()
    if not os.path.exists(CONFIG_FILE):
        config['SETTINGS'] = {
            'AutoDeleteArchive': 'True',
            'BandizipPath': r'C:\Program Files\Bandizip\Bandizip.exe'
        }
        # 预设一些常见的 X-Plane 路径规则示例
        config['RULES'] = {
            'X-Plane 12': r'D:\X-Plane 12\Custom Data',
            'FENIX A320': r'D:\Games\MSFS\Community\fnx-aircraft-320\NavData'
        }
        save_config(config)
    else:
        config.read(CONFIG_FILE, encoding='utf-8')
        if 'SETTINGS' not in config:
            old_path = config.get('PATHS', 'BandizipPath', fallback=r'C:\Program Files\Bandizip\Bandizip.exe')
            config['SETTINGS'] = {'AutoDeleteArchive': 'True', 'BandizipPath': old_path}
            save_config(config)
    return config


def save_config(config):
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        config.write(f)


# ================= 2. 快捷方式自动化 =================
def ensure_local_shortcut():
    """在当前目录下创建一个方便拖拽的快捷方式"""
    try:
        import win32com.client
        current_script = os.path.abspath(__file__)
        current_dir = os.path.dirname(current_script)
        shortcut_path = os.path.join(current_dir, "X-Plane 导航数据安装器.lnk")

        if not os.path.exists(shortcut_path):
            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut(shortcut_path)
            shortcut.TargetPath = sys.executable
            shortcut.Arguments = f'"{current_script}"'
            shortcut.WorkingDirectory = current_dir
            shortcut.IconLocation = sys.executable
            shortcut.Description = "X-Plane 导航数据安装器 - 拖拽压缩包至此"
            shortcut.save()
            print("✨ 已生成快捷方式：X-Plane 导航数据安装器")
    except Exception:
        pass


# ================= 3. 管理菜单 =================
def interactive_menu(config):
    ensure_local_shortcut()
    while True:
        os.system('cls' if os.name == 'nt' else 'clear')
        print("========================================")
        print("       X-Plane 导航数据安装器 - 管理菜单")
        print("========================================")
        print(f" 1. Bandizip 路径: {config['SETTINGS'].get('BandizipPath')}")
        print(f" 2. 自动删除原包: {'[开启 ✅]' if config['SETTINGS'].getboolean('AutoDeleteArchive') else '[关闭 ❌]'}")
        print("-" * 40)
        print(" 3. 查看 / 添加 / 修改 安装规则")
        print(" 4. 删除现有规则")
        print("-" * 40)
        print(" 0. 保存并退出")
        print("========================================")

        choice = input("请选择操作序号: ").strip()

        if choice == '1':
            path = input("请粘贴 Bandizip.exe 的完整路径: ").strip().strip('"')
            if os.path.exists(path):
                config['SETTINGS']['BandizipPath'] = path
                save_config(config)
            else:
                print("❌ 路径无效，请检查路径是否正确！");
                time.sleep(1.5)
        elif choice == '2':
            current = config['SETTINGS'].getboolean('AutoDeleteArchive')
            config['SETTINGS']['AutoDeleteArchive'] = str(not current)
            save_config(config)
        elif choice == '3':
            print("\n现有规则:")
            for k, v in config['RULES'].items(): print(f"  {k} -> {v}")
            name = input("\n请输入压缩包内子包名称 (不带.zip): ").strip()
            path = input("请输入对应的目标安装目录: ").strip().strip('"')
            if name and path:
                config['RULES'][name] = path
                save_config(config)
        elif choice == '4':
            keys = list(config['RULES'].keys())
            for i, k in enumerate(keys): print(f" [{i + 1}] {k}")
            idx = input("请输入要删除的规则序号: ").strip()
            if idx.isdigit() and 0 < int(idx) <= len(keys):
                del config['RULES'][keys[int(idx) - 1]]
                save_config(config)
        elif choice == '0':
            break


# ================= 4. 核心处理逻辑 =================
def merge_copy(src, dst):
    if not os.path.exists(dst):
        os.makedirs(dst, exist_ok=True)
    for item in os.listdir(src):
        s = os.path.join(src, item)
        d = os.path.join(dst, item)
        if os.path.isdir(s):
            merge_copy(s, d)
        else:
            shutil.copy2(s, d)


def process_archive(archive_path, config):
    bz_path = config['SETTINGS'].get('BandizipPath')
    auto_delete = config['SETTINGS'].getboolean('AutoDeleteArchive')
    rules = config['RULES']

    if not os.path.exists(bz_path):
        print(f"❌ 错误：未找到 Bandizip。请先双击运行脚本设置路径。")
        input("按回车退出...");
        return

    with tempfile.TemporaryDirectory() as stage1_dir:
        print(f"🚀 X-Plane 导航数据安装器正在处理: {os.path.basename(archive_path)}")
        subprocess.run([bz_path, "x", f"-o:{stage1_dir}", "-y", archive_path], capture_output=True)

        found_zips = []
        match_count = 0

        for root, _, files in os.walk(stage1_dir):
            for file in files:
                if file.lower().endswith('.zip'):
                    name = os.path.splitext(file)[0]
                    found_zips.append(name)
                    if name in rules:
                        match_count += 1
                        print(f"📦 匹配子包: {file}")
                        with tempfile.TemporaryDirectory() as stage2_dir:
                            subprocess.run([bz_path, "x", f"-o:{stage2_dir}", "-y", os.path.join(root, file)],
                                           capture_output=True)
                            actual_src = stage2_dir
                            content = os.listdir(stage2_dir)
                            if len(content) == 1 and os.path.isdir(os.path.join(stage2_dir, content[0])):
                                actual_src = os.path.join(stage2_dir, content[0])
                            merge_copy(actual_src, rules[name])
                            print(f"✅ 数据已分发至: {rules[name]}")

        if match_count == 0:
            print("\n❌ 匹配失败！压缩包内含有的子包名为：")
            for n in sorted(list(set(found_zips))): print(f" - {n}")
            input("\n按回车退出并检查管理菜单中的规则设置...")
        else:
            print(f"\n✨ 安装任务全部完成！")
            if auto_delete:
                try:
                    os.remove(archive_path)
                except:
                    pass
            time.sleep(2)


if __name__ == "__main__":
    conf = load_config()
    if len(sys.argv) > 1:
        # 拖拽模式
        for arg in sys.argv[1:]:
            if arg.lower().endswith(('.rar', '.zip', '.7z')):
                process_archive(arg, conf)
    else:
        # 管理菜单模式
        interactive_menu(conf)