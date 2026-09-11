# -*- coding: utf-8 -*-
"""
一個用於處理 Access 資料庫和 CAD 檔案結構的自動化工具。

功能:
1. 完整備份模式：完整匯出所有資料庫資料表，並複製所有相關的 CAD 檔案。
   同時會在輸出目錄生成一個 'tables_to_export.json' 檔案，列出所有資料表。
2. 篩選模式：僅針對一個或多個指定的「品名版本」，匯出相關資料表數據，
   並複製對應的 CAD 檔案（組樹下的特定版本資料夾 + 對應的毛胚資料夾）。
   同樣會在主輸出目錄生成 'tables_to_export.json'。

設計哲學:
採用物件導向設計，在 DataAccessManger 物件初始化時，透過是否傳入
'product_versions' 列表來決定其工作模式。物件一旦建立，其工作模式
即被鎖定。之後只需呼叫其公開方法 export_tables() 和 copy_cad_files() 
即可執行對應操作，無需再傳遞參數。
"""
import pyodbc
import pandas as pd
import os
import shutil
import json
from tqdm import tqdm
from typing import List, Optional

class DataAccessManger:
    """
    一個用於處理資料庫和檔案管理的通用類別。
    在初始化時設定好所有參數 (包括可選的篩選列表)，
    物件將記住其工作模式 (篩選或完整)。
    之後可分別呼叫 export_tables() 和 copy_cad_files()。
    """
    def __init__(self, db_path: str, table_export_path: str, cad_source_path: str, cad_target_path: str, product_versions: Optional[List[str]] = None):
        """
        初始化 DataManager。

        Args:
            db_path (str): Access 資料庫的路徑。
            table_export_path (str): 匯出資料表的目標路徑。
            cad_source_path (str): CAD 檔案的來源路徑。
            cad_target_path (str): CAD 檔案的目標路徑。
            product_versions (Optional[List[str ]], optional): 
                需要篩選的 '品名版本' 列表。
                如果提供，物件將鎖定在【篩選模式】。
                如果為 None，物件將鎖定在【完整備份模式】。
        """
        self.db_path = db_path
        self.table_export_path = table_export_path
        self.cad_source_path = cad_source_path
        self.cad_target_path = cad_target_path
        self.product_versions = product_versions
        
        if not os.path.exists(self.db_path):
            raise FileNotFoundError(f"指定的資料庫檔案不存在: {self.db_path}")
            
        self.conn_str = (
            r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
            f'DBQ={self.db_path};'
        )
        os.makedirs(self.table_export_path, exist_ok=True)
        print("DataManager 初始化成功。")
        if self.product_versions:
            print(f"模式：篩選模式，目標版本共 {len(self.product_versions)} 個。")
        else:
            print("模式：完整備份模式。")

    def export_tables(self):
        """
        執行資料表匯出。
        會根據物件初始化時的狀態，自動選擇完整或篩選模式。
        """
        if self.product_versions:
            self._export_filtered_mode()
        else:
            self._export_full_mode()

    def copy_cad_files(self):
        """
        執行 CAD 檔案複製。
        會根據物件初始化時的狀態，自動選擇完整或篩選模式。
        """
        if self.product_versions:
            self._copy_specific_cad_files()
        else:
            self._copy_full_cad_files()

    # --- 以下為內部輔助方法 (Helper Methods) ---

    def _save_list_to_json(self, data_list, filename):
        """將列表儲存為 JSON 檔案。"""
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(data_list, f, indent=2, ensure_ascii=False)
        print(f">> 資料庫中的資料表清單已儲存至: {filename}")

    def _export_full_mode(self):
        print("\n>> 【完整模式】開始從 Access 匯出所有資料表...")
        conn = None
        try:
            conn = pyodbc.connect(self.conn_str)
            cursor = conn.cursor()
            all_table_names = [row.table_name for row in cursor.tables() if row.table_type == 'TABLE' and not row.table_name.startswith('~')]
            
            if not all_table_names:
                print("資料庫中找不到任何資料表。")
                return

            # --- 重新加入儲存 JSON 的功能 ---
            json_path = os.path.join(self.table_export_path, 'tables_to_export.json')
            self._save_list_to_json(all_table_names, json_path)
            
            print(f"找到 {len(all_table_names)} 個資料表，準備匯出...")
            for table_name in tqdm(all_table_names, desc="完整匯出進度"):
                df = pd.read_sql(f'SELECT * FROM [{table_name}]', conn)
                csv_path = os.path.join(self.table_export_path, f'{table_name}.csv')
                xlsx_path = os.path.join(self.table_export_path, f'{table_name}.xlsx')
                df.to_csv(csv_path, index=False, encoding='utf-8-sig')
                df.to_excel(xlsx_path, index=False)
            print("所有資料表已成功匯出！")
        except pyodbc.Error as ex:
            print(f"資料庫錯誤: {ex}")
        finally:
            if conn: conn.close()

    def _export_filtered_mode(self):
        """
        【篩選模式】已更新：
        - 對於含 '品名版本' 的表，將所有篩選結果合併儲存到一個以「純粹表名」命名的檔案。
        - 對於不含 '品名版本' 的表，以原名完整匯出。
        """
        print(f"\n>> 【篩選模式】啟動，處理 {len(self.product_versions)} 個目標版本...")
        
        conn = None
        try:
            conn = pyodbc.connect(self.conn_str)
            cursor = conn.cursor()
            all_table_names = [row.table_name for row in cursor.tables() if row.table_type == 'TABLE' and not row.table_name.startswith('~')]

            if not all_table_names:
                print("資料庫中找不到任何資料表。")
                return

            json_path = os.path.join(self.table_export_path, 'tables_to_export.json')
            self._save_list_to_json(all_table_names, json_path)
            
            print(f"找到 {len(all_table_names)} 個資料表，準備進行篩選與匯出...")
            for table_name in tqdm(all_table_names, desc=f"批次處理資料表"):
                columns = [row.column_name for row in cursor.columns(table=table_name)]
                
                df_to_save = None
                
                if '品名版本' in columns:
                    # 1. 處理需要篩選的資料表
                    placeholders = ', '.join(['?'] * len(self.product_versions))
                    query = f"SELECT * FROM [{table_name}] WHERE [品名版本] IN ({placeholders})"
                    # 將所有篩選結果一次性讀取到一個 DataFrame
                    df_to_save = pd.read_sql(query, conn, params=self.product_versions)
                else:
                    # 2. 處理不需篩選的資料表
                    df_to_save = pd.read_sql(f'SELECT * FROM [{table_name}]', conn)

                # 如果查詢結果不為空，則儲存檔案
                if df_to_save is not None and not df_to_save.empty:
                    csv_path = os.path.join(self.table_export_path, f'{table_name}.csv')
                    xlsx_path = os.path.join(self.table_export_path, f'{table_name}.xlsx')
                    df_to_save.to_csv(csv_path, index=False, encoding='utf-8-sig')
                    df_to_save.to_excel(xlsx_path, index=False)

            print(f"所有指定版本的篩選匯出任務完成！")
        except pyodbc.Error as ex:
            print(f"資料庫錯誤: {ex}")
        finally:
            if conn: conn.close()
    
    def _copy_full_cad_files(self):
        print("\n>> 【完整模式】正在複製所有CAD檔案...")
        source_root = self.cad_source_path
        target_root = self.cad_target_path

        if not os.path.isdir(source_root):
            print(f"錯誤：CAD 來源目錄不存在 -> {source_root}")
            return
            
        all_dirs_to_copy = []
        for customer in os.listdir(source_root):
            customer_path = os.path.join(source_root, customer)
            if not os.path.isdir(customer_path): continue
            
            for product_code in os.listdir(customer_path):
                product_path = os.path.join(customer_path, product_code)
                if not os.path.isdir(product_path): continue
                all_dirs_to_copy.append(product_path)

        # 定義要忽略的檔案（例如系統檔）
        ignore_func = shutil.ignore_patterns('desktop.ini', 'Thumbs.db', '.DS_Store')

        with tqdm(total=len(all_dirs_to_copy), desc="完整複製進度", unit="產品") as pbar:
            for product_path in all_dirs_to_copy:
                for subfolder in ['毛胚', '組樹']:
                    subfolder_path = os.path.join(product_path, subfolder)
                    if os.path.exists(subfolder_path):
                        relative_path = os.path.relpath(subfolder_path, source_root)
                        target_path = os.path.join(target_root, relative_path)
                        
                        try:
                            # 加入 ignore 參數
                            shutil.copytree(
                                subfolder_path, 
                                target_path, 
                                dirs_exist_ok=True, 
                                ignore=ignore_func
                            )
                        except Exception as e:
                            # 即使部分檔案失敗（如權限問題），也繼續處理下一個資料夾
                            print(f"\n警告：跳過部分檔案或目錄 {subfolder_path}，原因：{e}")
                            
                pbar.update(1)
        print("CAD 檔案完整複製完成！")

    def _copy_specific_cad_files(self):
        print(f"\n>> 【篩選模式】開始尋找並複製 {len(self.product_versions)} 個版本的 CAD 檔案...")
        source_root = self.cad_source_path
        target_root = self.cad_target_path
        
        if not os.path.isdir(source_root):
            print(f"錯誤：CAD 來源目錄不存在 -> {source_root}")
            return
            
        versions_to_find = set(self.product_versions)
        found_versions = set()

        with tqdm(total=len(versions_to_find), desc="尋找CAD檔案") as pbar:
            for dirpath, dirnames, filenames in os.walk(source_root):
                intersection = versions_to_find.intersection(dirnames)
                
                if not intersection:
                    continue

                if os.path.basename(dirpath) == '組樹':
                    for version in intersection:
                        found_zushu_path = os.path.join(dirpath, version)
                        product_root_path = os.path.dirname(dirpath)
                        
                        relative_path_zushu = os.path.relpath(found_zushu_path, source_root)
                        target_zushu_path = os.path.join(target_root, relative_path_zushu)
                        shutil.copytree(found_zushu_path, target_zushu_path, dirs_exist_ok=True)
                        
                        source_maopei_path = os.path.join(product_root_path, '毛胚')
                        if os.path.exists(source_maopei_path):
                            relative_path_maopei = os.path.relpath(source_maopei_path, source_root)
                            target_maopei_path = os.path.join(target_root, relative_path_maopei)
                            shutil.copytree(source_maopei_path, target_maopei_path, dirs_exist_ok=True)
                        
                        if version not in found_versions:
                           pbar.set_description(f"找到: {version}")
                           pbar.update(1)
                           found_versions.add(version)
                
                if found_versions == versions_to_find:
                    pbar.n = pbar.total
                    pbar.refresh()
                    break
        
        not_found = versions_to_find - found_versions
        if not_found:
            print(f"\n注意：在 {source_root} 中未能找到以下版本的組樹子資料夾: {list(not_found)}")
        
        print("特定 CAD 檔案複製完成！")
"""
# --- 如何使用這個新設計的 Class ---
if __name__ == '__main__':
    # --- 步驟 1: 在 Class 外部準備所有需要的路徑 ---
    # 這裡依然可以使用您自訂義的模組來產生根目錄
    from SySPath import ChynWangDatapy, ChynWangPojectpy
    
    CW_root = ChynWangDatapy()
    CP_root = ChynWangPojectpy()

    # 根據根目錄組合出最終需要的【具體路徑】
    db_full_path = os.path.join(CP_root, "access_vba_scripts", "CastDataBase.accdb")
    export_folder_path = os.path.join(CP_root, 'data', 'access_inputs_table_test')
    cad_source_folder_path = os.path.join(CW_root, "3D_CAD")
    cad_target_folder_path = os.path.join(CP_root, "data", "access_raw_CAD_test")

    # --- 步驟 2: 將這些具體路徑傳入 Class 來建立物件 ---
    # 這種設計讓 DataManager 類別更加獨立和靈活
    # 它不再關心路徑是如何來的，只關心拿到路徑後要做什麼
    try:
        data_handler = DataAccessManger(
            db_path=db_full_path,
            table_export_path=export_folder_path,
            cad_source_path=cad_source_folder_path,
            cad_target_path=cad_target_folder_path,product_versions=["103-200159-M210416"]
        )

        # --- 步驟 3: 呼叫需要的功能 ---
        data_handler.export_tables()
        data_handler.copy_cad_files()        
        # # 取消註解以執行 CAD 檔案複製
        # data_handler.copy_cad_files()

    except FileNotFoundError as e:
        print(f"初始化失敗：{e}")
    except Exception as e:
        print(f"執行過程中發生錯誤：{e}")
"""

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
        data_handler.export_tables()
        data_handler.copy_cad_files()        
        # # 取消註解以執行 CAD 檔案複製
        # data_handler.copy_cad_files()

    except FileNotFoundError as e:
        print(f"初始化失敗：{e}")
    except Exception as e:
        print(f"執行過程中發生錯誤：{e}")
