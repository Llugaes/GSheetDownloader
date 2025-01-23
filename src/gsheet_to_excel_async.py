### ============================================
# 使用指南：
# 1. 请确保您的认证JSON文件路径已设置到环境变量`GOOGLE_APPLICATION_CREDENTIALS`中
#    这样GCP的API会自动读取该路径下的JSON文件以获取认证信息。
# 2. 以命令行方式运行此脚本，并提供多个Google Spreadsheet的ID：
#    python gsheet_to_excel.py <spreadsheet_id1> <spreadsheet_id2> ... --output-dir <path/to/save/>
### ============================================

import asyncio
import os
import argparse
from openpyxl import Workbook
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

# 环境变量名
CREDENTIALS_ENV_VAR = "GCP_CREDENTIALS_JSON"
TOKEN_ENV_VAR = "GCP_TOKEN_JSON"
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']


def get_sheets_service_v4():
    creds = None
    credentials_json_path = os.getenv(CREDENTIALS_ENV_VAR)
    token_json_path = os.getenv(TOKEN_ENV_VAR)
    
    print(f"[Auth] 凭证路径: {credentials_json_path}")
    print(f"[Auth] Token路径: {token_json_path}")

    # 首先检查环境变量是否设置
    if not credentials_json_path or not token_json_path:
        print("[Auth] 环境变量未设置，尝试使用默认路径")
        app_data_dir = os.path.join(os.path.expanduser("~"), ".gsheet_downloader")
        credentials_json_path = os.path.join(app_data_dir, "credentials.json")
        token_json_path = os.path.join(app_data_dir, "token.json")
        
        # 设置环境变量
        os.environ[CREDENTIALS_ENV_VAR] = credentials_json_path
        os.environ[TOKEN_ENV_VAR] = token_json_path
        
        print(f"[Auth] 使用默认路径 - 凭证: {credentials_json_path}")
        print(f"[Auth] 使用默认路径 - Token: {token_json_path}")

    # 检查凭证文件是否存在
    if not os.path.exists(credentials_json_path):
        raise ValueError(f"找不到凭证文件: {credentials_json_path}")

    # 尝试加载现有token
    if os.path.exists(token_json_path):
        try:
            creds = Credentials.from_authorized_user_file(token_json_path, SCOPES)
            print("[Auth] 成功加载现有token")
        except Exception as e:
            print(f"[Auth] 加载token失败: {str(e)}")
            creds = None
    
    # 处理凭证
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                print("[Auth] 刷新过期token")
                creds.refresh(Request())
            except Exception as e:
                print(f"[Auth] 刷新token失败: {str(e)}")
                creds = None
        
        if not creds:
            print("[Auth] 开始新的认证流程")
            try:
                flow = InstalledAppFlow.from_client_secrets_file(
                    credentials_json_path, SCOPES
                )
                auth_url, _ = flow.authorization_url()
                setattr(get_sheets_service_v4, 'auth_url', auth_url)
                
                creds = flow.run_local_server(port=0, open_browser=True)
                
                # 保存新token
                os.makedirs(os.path.dirname(token_json_path), exist_ok=True)
                with open(token_json_path, "w") as token:
                    token.write(creds.to_json())
                print("[Auth] 新token已保存")
            except Exception as e:
                raise ValueError(f"认证过程失败: {str(e)}")

    return build('sheets', 'v4', credentials=creds)

def _download_gsheet(service, spreadsheet_id, output_dir):
    # 获取电子表格元数据以获得文件名
    spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = spreadsheet.get('sheets', [])
    sheet_title = spreadsheet.get('properties', {}).get('title', 'google_sheet_data')

    # 处理输出路径逻辑
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    output_path = os.path.join(output_dir, f"{sheet_title}.xlsx")

    # 初始化一个Excel writer对象
    excel_writer = pd.ExcelWriter(output_path, engine='openpyxl')

    # 遍历所有工作表并获取数据
    for sheet in sheets:
        sheet_name = sheet['properties']['title']
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=sheet_name
        ).execute()
        
        values = result.get('values', [])
        
        if values:
            # 过滤掉完全为空的行
            values = [row for row in values if any(row)]
            
            # 确保所有行的列数与表头一致
            num_columns = len(values[0])
            for i in range(1, len(values)):
                if len(values[i]) < num_columns:
                    values[i].extend([None] * (num_columns - len(values[i])))
                elif len(values[i]) > num_columns:
                    values[i] = values[i][:num_columns]
            
            # 将数据转换为DataFrame
            df = pd.DataFrame(values[1:], columns=values[0])
            df.to_excel(excel_writer, sheet_name=sheet_name, index=False)
        else:
            print(f"No data found in sheet: {sheet_name}")

    # 保存并关闭Excel文件
    excel_writer.close()

    print(f"数据已保存到 {output_path}")

async def _download_gsheet_async(service, spreadsheet_id, output_dir):
    try:
        print(f"[Async] 开始下载单个文件: {spreadsheet_id}")
        
        # 验证输出目录
        if not output_dir or not isinstance(output_dir, str):
            output_dir = os.path.join(os.path.expanduser("~"), "Downloads")
            print(f"[Async] 使用默认输出目录: {output_dir}")
        
        # 确保输出目录是绝对路径
        output_dir = os.path.abspath(output_dir)
        print(f"[Async] 使用绝对路径: {output_dir}")
        
        # 确保输出目录存在
        try:
            os.makedirs(output_dir, exist_ok=True)
            print(f"[Async] 输出目录已确认: {output_dir}")
        except Exception as e:
            raise ValueError(f"创建输出目录失败: {str(e)}")
        
        if not os.path.exists(output_dir):
            raise ValueError(f"输出目录不存在且无法创建: {output_dir}")
            
        print(f"[Async] 获取spreadsheet信息")
        spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = spreadsheet.get('sheets', [])
        file_name = spreadsheet.get('properties', {}).get('title', 'untitled')
        
        excel_path = os.path.join(output_dir, f"{file_name}.xlsx")
        print(f"[Async] 输出文件路径: {excel_path}")
        
        wb = Workbook()
        wb.remove(wb.active)
        
        for sheet in sheets:
            sheet_name = sheet['properties']['title']
            print(f"[Async] 处理工作表: {sheet_name}")
            
            if sheet['properties'].get('hidden', False):
                print(f"[Async] 跳过隐藏工作表: {sheet_name}")
                continue
                
            result = service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=sheet_name
            ).execute()
            
            values = result.get('values', [])
            if not values:
                print(f"[Async] 工作表为空: {sheet_name}")
                continue
            
            print(f"[Async] 写入工作表: {sheet_name}")
            ws = wb.create_sheet(sheet_name)
            for row in values:
                ws.append(row)
        
        print(f"[Async] 保存Excel文件: {excel_path}")
        wb.save(excel_path)
        print(f"[Async] 文件保存成功: {excel_path}")
        return excel_path
        
    except Exception as e:
        error_msg = f"下载 {spreadsheet_id} 时出错: {str(e)}"
        print(f"[Async] {error_msg}")
        raise Exception(error_msg)

# 添加兼容性函数
async def _async_timeout(seconds):
    try:
        await asyncio.sleep(seconds)
        raise TimeoutError()
    except asyncio.CancelledError:
        pass
    finally:
        yield

def download_google_sheet(spreadsheet_id, output_dir):
    try:
        service = get_sheets_service_v4()

        _download_gsheet(service, spreadsheet_id, output_dir)
    except HttpError as err:
        print(f"An HTTP error occurred: {err}")
    except ValueError as ve:
        print(f"ValueError: {ve}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

async def download_google_sheet_async(spreadsheet_id, output_dir):
    try:
        service = get_sheets_service_v4()

        await asyncio.to_thread(_download_gsheet, service, spreadsheet_id, output_dir)
    except HttpError as err:
        print(f"An HTTP error occurred: {err}")
    except ValueError as ve:
        print(f"ValueError: {ve}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        
async def download_multi_google_sheet_async(spreadsheet_id_list, output_dir):
    try:
        print(f"[Async] 开始多文件下载，sheet_ids={spreadsheet_id_list}, output_dir={output_dir}")
        service = get_sheets_service_v4()
        print("[Async] 获取service成功")
        
        # 使用 gather 替代 TaskGroup
        tasks = []
        for spreadsheet_id in spreadsheet_id_list:
            print(f"[Async] 创建下载任务: {spreadsheet_id}")
            task = _download_gsheet_async(service, spreadsheet_id, output_dir)
            tasks.append(task)
            
        print(f"[Async] 创建任务列表: {len(tasks)}个任务")
        
        if not tasks:
            raise ValueError("没有创建任何下载任务")
            
        results = await asyncio.gather(*tasks, return_exceptions=True)
        print("[Async] 所有任务完成")
        
        # 检查每个任务的结果
        for i, result in enumerate(results):
            if isinstance(result, Exception):
                print(f"[Async] 任务 {i} 失败: {str(result)}")
                raise result
            else:
                print(f"[Async] 任务 {i} 成功: {result}")
        
        setattr(download_multi_google_sheet_async, 'success', True)
        
    except Exception as e:
        error_msg = str(e)
        print(f"[Async] 下载出错: {error_msg}")
        setattr(download_multi_google_sheet_async, 'last_error', error_msg)
        setattr(download_multi_google_sheet_async, 'success', False)
        raise Exception(error_msg)

def main():
    parser = argparse.ArgumentParser(description="从Google Sheets下载数据并保存为Excel文件")
    parser.add_argument("spreadsheet_ids", nargs='+', help="需要下载的Google Spreadsheet的ID列表, 可以在URL中找到")
    parser.add_argument("--output-dir", help="Excel文件的保存目录（不包含文件名）", required=True)
    
    args = parser.parse_args()
    
    # for spreadsheet_id in args.spreadsheet_ids:
    #     download_google_sheet(spreadsheet_id, args.output_dir)
    asyncio.run(download_multi_google_sheet_async(args.spreadsheet_ids, args.output_dir))
if __name__ == "__main__":
    main()
