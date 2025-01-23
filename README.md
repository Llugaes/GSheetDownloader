# GSheet下载器

一个简单的GUI工具,用于将Google Sheets文档下载为Excel格式。

## 功能特点

- 支持批量下载多个Google Sheets文档
- 保留原始工作表格式和数据
- 自动跳过隐藏的工作表
- 支持自定义输出目录
- 保存最近使用过的Sheet列表

## 安装要求

- Python 3.10+
- 依赖包:
  - google-api-python-client==2.108.0
  - google-auth-oauthlib==1.1.0
  - openpyxl==3.1.2

## 开发环境设置

1. 克隆仓库
```bash
git clone https://github.com/Llugaes/GSheetDownloader.git
cd GSheetDownloader
```

2. 创建并激活虚拟环境
```bash
python -m venv venv
source venv/bin/activate  # Linux/Mac
venv\Scripts\activate     # Windows
```

3. 安装依赖
```bash
pip install -r requirements.txt
```

## 使用说明

### 认证设置

1. 打开下载器后,点击右上角的"认证设置"按钮
2. 在弹出的认证设置窗口中,点击"选择认证文件"按钮
3. 在文件选择对话框中,选择从GCP官网中导出的 `credentials.json` 文件
4. 选择完成后会自动打开浏览器进行Google账号授权
5. 授权完成后即可使用下载功能

### 下载Google表格

1. 在输入框中粘贴Google表格的URL或ID
2. 点击"添加"按钮将表格添加到下载列表
3. 选择保存Excel文件的输出目录
4. 点击"下载选中"或"下载全部"按钮开始下载
5. 下载完成后会在指定目录生成对应的Excel文件

### 注意事项

- 首次使用需要完成认证设置才能使用下载功能
- 认证信息会保存在用户目录下的 `.gsheet_downloader` 文件夹中
- 如需重新认证,可在认证设置中删除现有认证信息

## 打包发布

使用PyInstaller打包为可执行文件:

```bash
pyinstaller gsheet_downloader.spec
```

打包后的文件将生成在 `dist` 目录下。

## 配置文件

程序会在当前目录下创建 `config.json` 文件保存配置信息:
- recent_sheets: 最近使用的Sheet列表
- output_dir: 默认输出目录

## 开发相关文件说明

- `src/gui_main.py`: 主程序入口和GUI实现
- `src/gsheet_to_excel_async.py`: Google Sheets下载核心逻辑
- `src/config_manager.py`: 配置管理
- `runtime_hook.py`: PyInstaller运行时钩子
- `gsheet_downloader.spec`: PyInstaller打包配置

## 许可证

MIT License

## 贡献指南

1. Fork 项目
2. 创建特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 开启 Pull Request