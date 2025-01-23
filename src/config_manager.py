import json
import os

class ConfigManager:
    def __init__(self, config_file='config.json'):
        self.config_file = config_file
        self.config = self.load_config()

    def load_config(self):
        default_config = {
            'recent_sheets': [],
            'output_dir': os.path.join(os.path.expanduser("~"), "Downloads")  # 设置默认下载目录
        }
        
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    # 使用默认值更新配置
                    default_config.update(config)
                    return default_config
            except:
                return default_config
        return default_config

    def save_config(self):
        # 确保目录存在
        config_dir = os.path.dirname(os.path.abspath(self.config_file))
        os.makedirs(config_dir, exist_ok=True)
        
        # 确保配置完整
        if 'recent_sheets' not in self.config:
            self.config['recent_sheets'] = []
        if 'output_dir' not in self.config:
            self.config['output_dir'] = ''
        
        with open(self.config_file, 'w', encoding='utf-8') as f:
            json.dump(self.config, f, ensure_ascii=False, indent=2)

    def add_sheet(self, sheet_url, sheet_name=''):
        print(f"[ConfigManager] 添加sheet, URL: {sheet_url}")
        sheet_id = self.extract_sheet_id(sheet_url)
        print(f"[ConfigManager] 提取的sheet_id: {sheet_id}")
        
        if sheet_id:
            recent_sheets = self.config['recent_sheets']
            # 检查是否已存在
            for sheet in recent_sheets:
                if sheet['id'] == sheet_id:
                    print(f"[ConfigManager] Sheet已存在: {sheet_id}")
                    return
            # 添加新的sheet
            new_sheet = {
                'id': sheet_id,
                'url': sheet_url,
                'name': sheet_name
            }
            print(f"[ConfigManager] 添加新sheet: {new_sheet}")
            recent_sheets.append(new_sheet)
            self.save_config()

    @staticmethod
    def extract_sheet_id(input_str: str) -> str | None:
        match input_str:
            case str() if '/' not in input_str:
                return input_str
            case str() if '/spreadsheets/d/' in input_str:
                parts = input_str.split('/spreadsheets/d/')
                return parts[1].split('/')[0].split('?')[0]
            case str() if '/spreadsheets/' in input_str:
                parts = input_str.split('/spreadsheets/')
                return parts[1].split('/')[0].split('?')[0]
            case _:
                return None 