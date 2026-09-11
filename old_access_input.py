import pyodbc
import pandas as pd
import os
import shutil
import json
from tqdm import tqdm
import argparse
class DataAccessManger:
    
    """
    一個用於處理資料庫和檔案管理的通用類別。
    - 初始化時直接接收所需的操作路徑，而不是根目錄。
    - 從 Access 資料庫匯出資料表。
    - 複製指定的檔案目錄結構。
    """
    def __init__(self, db_path: str, table_export_path: str, cad_source_path: str, cad_target_path: str):
        """
        初始化 DataManager。

        Args:
            db_path (str): Access 資料庫的完整檔案路徑 (例如 'C:/.../CastDataBase.accdb')。
            table_export_path (str): 匯出 CSV/Excel 檔案的目標資料夾路徑。
            cad_source_path (str): CAD 檔案的來源根目錄。
            cad_target_path (str): 要將 CAD 檔案複製到的目標根目錄。
        """
        # --- 1. 直接儲存傳入的具體路徑 ---
        self.db_path = db_path
        self.table_export_path = table_export_path
        self.cad_source_path = cad_source_path
        self.cad_target_path = cad_target_path
        
        # --- 2. 建立資料庫連線字串 ---
        if not os.path.exists(self.db_path):
            # 提前檢查，避免連線時才報錯
            raise FileNotFoundError(f"指定的資料庫檔案不存在: {self.db_path}")
            
        self.conn_str = (
            r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
            f'DBQ={self.db_path};'
        )

        # --- 3. 確保輸出目錄存在 ---
        os.makedirs(self.table_export_path, exist_ok=True)
        # 複製 CAD 的目標目錄在執行時再創建，確保結構正確
        
        print("DataManager 初始化成功，路徑設定如下：")
        print(f"  - 資料庫路徑: {self.db_path}")
        print(f"  - 資料表匯出路徑: {self.table_export_path}")
        print(f"  - CAD 來源路徑: {self.cad_source_path}")
        print(f"  - CAD 目標路徑: {self.cad_target_path}")
    def save_list_to_json(self, data_list, filename):
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(data_list, f, indent=2, ensure_ascii=False)
        print(f">> 工作清單已儲存至 {filename}")
    def export_tables_to_files(self):
        """
        連接到 Access 資料庫，並將所有資料表匯出為 CSV 和 Excel 檔案。
        """
        conn = None
        try:
            print("\n>> 開始從 Access 匯出資料表...")
            conn = pyodbc.connect(self.conn_str)
            cursor = conn.cursor()

            all_table_names = [
                row.table_name for row in cursor.tables()
                if row.table_type == 'TABLE' and not row.table_name.startswith('~')
            ]
            # 將列表和要儲存的檔名傳遞過去
            all_table_names_path = os.path.join(self.table_export_path, f'tables_to_export.json')
            self.save_list_to_json(all_table_names,all_table_names_path )
            if not all_table_names:
                print("資料庫中找不到任何資料表。")
                return

            print(f"找到 {len(all_table_names)} 個資料表，準備匯出...")

            for table_name in tqdm(all_table_names, desc="匯出進度"):
                query = f'SELECT * FROM [{table_name}]'
                df = pd.read_sql(query, conn)
                csv_file_path = os.path.join(self.table_export_path, f'{table_name}.csv')
                xlsx_file_path = os.path.join(self.table_export_path, f'{table_name}.xlsx')
                df.to_csv(csv_file_path, index=False)
                df.to_excel(xlsx_file_path, index=False)
            
            print("所有資料表已成功匯出！")

        except pyodbc.Error as ex:
            sqlstate = ex.args[0]
            print(f"資料庫連線或查詢錯誤: {sqlstate}\n{ex}")
        except Exception as e:
            print(f"發生未預期的錯誤: {e}")
        finally:
            if conn:
                conn.close()
                print("資料庫連線已關閉。")
        
    def copy_cad_files(self):
        """
        從來源資料夾遞迴複製 CAD 檔案到目標資料夾，並顯示進度條。
        """
        print("\n>> 開始複製 CAD 檔案結構...")
        source_root = self.cad_source_path
        target_root = self.cad_target_path

        if not os.path.exists(source_root):
            print(f"錯誤：CAD 來源目錄不存在 -> {source_root}")
            return
            
        total_tasks = sum(
            len(os.listdir(os.path.join(source_root, customer)))
            for customer in os.listdir(source_root)
            if os.path.isdir(os.path.join(source_root, customer))
        ) * 2

        with tqdm(total=total_tasks, desc="複製進度", unit="目錄") as pbar:
            for customer in os.listdir(source_root):
                customer_path = os.path.join(source_root, customer)
                if not os.path.isdir(customer_path): continue
                
                for product_code in os.listdir(customer_path):
                    product_path = os.path.join(customer_path, product_code)
                    if not os.path.isdir(product_path): continue
                    
                    for subfolder in ['毛胚', '組樹']:
                        subfolder_path = os.path.join(product_path, subfolder)
                        if os.path.exists(subfolder_path):
                            target_path = os.path.join(target_root, customer, product_code, subfolder)
                            os.makedirs(target_path, exist_ok=True)
                            
                            for item in os.listdir(subfolder_path):
                                src_item = os.path.join(subfolder_path, item)
                                dst_item = os.path.join(target_path, item)
                                if os.path.isdir(src_item):
                                    shutil.copytree(src_item, dst_item, dirs_exist_ok=True)
                                else:
                                    shutil.copy2(src_item, dst_item)
                        pbar.update(1)

        print("CAD 檔案複製完成！")

# --- 如何使用這個新設計的 Class ---
if __name__ == '__main__':
    # --- 步驟 1: 在 Class 外部準備所有需要的路徑 ---
    # 這裡依然可以使用您自訂義的模組來產生根目錄
    from SySPath import ChynWangDatapy, ChynWangPojectpy
    
    CW_root = ChynWangDatapy()
    CP_root = ChynWangPojectpy()

    # 根據根目錄組合出最終需要的【具體路徑】
    db_full_path = os.path.join(CP_root, "access_vba_scripts", "CastDataBase.accdb")
    export_folder_path = os.path.join(CP_root, 'data', 'access_inputs_table')
    cad_source_folder_path = os.path.join(CW_root, "3D_CAD")
    cad_target_folder_path = os.path.join(CP_root, "data", "access_raw_CAD")

    # --- 步驟 2: 將這些具體路徑傳入 Class 來建立物件 ---
    # 這種設計讓 DataManager 類別更加獨立和靈活
    # 它不再關心路徑是如何來的，只關心拿到路徑後要做什麼
    try:
        data_handler = DataAccessManger(
            db_path=db_full_path,
            table_export_path=export_folder_path,
            cad_source_path=cad_source_folder_path,
            cad_target_path=cad_target_folder_path
        )

        # --- 步驟 3: 呼叫需要的功能 ---
        data_handler.export_tables_to_files()
        data_handler.copy_cad_files()        
        # # 取消註解以執行 CAD 檔案複製
        # data_handler.copy_cad_files()

    except FileNotFoundError as e:
        print(f"初始化失敗：{e}")
    except Exception as e:
        print(f"執行過程中發生錯誤：{e}")
