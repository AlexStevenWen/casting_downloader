import json
import argparse
import sys
# 假設 access_input.py 和 DataAccessManger 類別已存在
from access_input import DataAccessManger

def main():
    """
    主執行函式：解析參數、讀取設定、執行資料處理任務，並回傳執行結果。

    :return: 一個包含狀態和訊息的字典。
             - 成功: {'status': 'success', 'message': '...', 'data': {...}}
             - 失敗: {'status': 'error', 'message': '...'}
    """
    # 1. 設定參數解析器
    parser = argparse.ArgumentParser(description="資料庫匯出與 CAD 檔案複製工具")
    parser.add_argument("--config", required=True, help="指向 JSON 設定檔的路徑")
    args = parser.parse_args()

    # 2. 讀取 JSON 設定檔
    try:
        with open(args.config, "r", encoding="utf-8") as f:
            config = json.load(f)
    except FileNotFoundError:
        error_msg = f"錯誤：設定檔 '{args.config}' 不存在。"
        print(error_msg)
        return {"status": "error", "message": error_msg}
    except json.JSONDecodeError:
        error_msg = f"錯誤：設定檔 '{args.config}' 格式不正確。"
        print(error_msg)
        return {"status": "error", "message": error_msg}
    except Exception as e:
        error_msg = f"讀取設定檔時發生未知錯誤: {e}"
        print(error_msg)
        return {"status": "error", "message": error_msg}

    # 3. 初始化 DataAccessManger
    try:
        kwargs = {}
        # 處理【必填】參數
        kwargs['db_path'] = config['db_path']
        kwargs['table_export_path'] = config['table_export_path']
        kwargs['cad_source_path'] = config['cad_source_path']
        kwargs['cad_target_path'] = config['cad_target_path']

        # 處理【選填】參數
        if 'product_versions' in config:
            kwargs['product_versions'] = config['product_versions']

        data_handler = DataAccessManger(**kwargs)

    except KeyError as e:
        error_msg = f"錯誤：設定檔中缺少必要的鍵值: {e}"
        print(error_msg)
        return {"status": "error", "message": error_msg}
    except Exception as e:
        error_msg = f"初始化 DataAccessManger 時發生錯誤: {e}"
        print(error_msg)
        return {"status": "error", "message": error_msg}

    # 4. 執行主要功能
    try:
        print("\n>> 開始執行資料匯出...")
        data_handler.export_tables()
        
        print("\n>> 開始執行 CAD 檔案複製...")
        data_handler.copy_cad_files()
        
    except Exception as e:
        error_msg = f"執行任務時發生錯誤: {e}"
        print(error_msg)
        return {"status": "error", "message": error_msg}
    
    success_msg = "所有任務執行完畢！"
    print(f"\n>> {success_msg}")
    
    # 回傳結構化的成功資訊
    return {
        "status": "success",
        "message": success_msg,
        "data": {
            "table_export_path": config['table_export_path'],
            "cad_target_path": config['cad_target_path']
        }
    }

if __name__ == "__main__":
    result = main()

    # 美化輸出，讓結果一目了然
    print("\n--- 執行結果 ---")
    print(json.dumps(result, indent=4, ensure_ascii=False))
    
    # 根據結果狀態碼來結束程式
    # 成功時 exit code 為 0，失敗時為 1
    if result["status"] == "error":
        sys.exit(1)
    else:
        sys.exit(0)