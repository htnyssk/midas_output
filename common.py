import os
import sys
import yaml


def get_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def load_config():
    """config.yaml を読み込み (設定dict, 基準フォルダ) を返す。"""
    base_dir = get_base_dir()
    config_path = os.path.join(base_dir, "config.yaml")
    if not os.path.exists(config_path):
        print("Error: Configuration file 'config.yaml' not found.")
        sys.exit(1)
    with open(config_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f), base_dir
