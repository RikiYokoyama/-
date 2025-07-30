import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter import font

class FileListApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ファイル名一覧作成ツール")
        self.root.geometry("600x500")
        
        # 選択されたフォルダパス
        self.folder_path = tk.StringVar()
        self.output_path = tk.StringVar()
        
        # フォントの設定
        default_font = font.nametofont("TkDefaultFont")
        default_font.configure(size=10)
        
        self.create_widgets()
    
    def create_widgets(self):
        # メインフレーム
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # フォルダ選択セクション
        folder_frame = ttk.LabelFrame(main_frame, text="対象フォルダ", padding="10")
        folder_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(folder_frame, text="フォルダパス:").grid(row=0, column=0, sticky=tk.W)
        
        folder_entry = ttk.Entry(folder_frame, textvariable=self.folder_path, width=50)
        folder_entry.grid(row=1, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        
        ttk.Button(folder_frame, text="参照", command=self.select_folder).grid(row=1, column=1)
        
        # 出力設定セクション
        output_frame = ttk.LabelFrame(main_frame, text="出力設定", padding="10")
        output_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(output_frame, text="出力ファイル名:").grid(row=0, column=0, sticky=tk.W)
        
        output_entry = ttk.Entry(output_frame, textvariable=self.output_path, width=50)
        output_entry.grid(row=1, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        self.output_path.set("output.xlsx")  # デフォルト値
        
        ttk.Button(output_frame, text="保存先選択", command=self.select_output).grid(row=1, column=1)
        
        # プレビューセクション
        preview_frame = ttk.LabelFrame(main_frame, text="ファイル一覧プレビュー", padding="10")
        preview_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        # Treeviewでファイル一覧を表示
        self.tree = ttk.Treeview(preview_frame, columns=("filename",), show="headings", height=10)
        self.tree.heading("filename", text="ファイル名")
        self.tree.column("filename", width=500)
        
        # スクロールバー
        scrollbar = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # ボタンセクション
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=10)
        
        ttk.Button(button_frame, text="プレビュー更新", command=self.update_preview).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="Excelファイル作成", command=self.create_excel).pack(side=tk.LEFT)
        
        # ステータスバー
        self.status_var = tk.StringVar()
        self.status_var.set("フォルダを選択してください")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # グリッドの重みを設定
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)
        folder_frame.columnconfigure(0, weight=1)
        output_frame.columnconfigure(0, weight=1)
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
    
    def select_folder(self):
        """フォルダ選択ダイアログを開く"""
        folder = filedialog.askdirectory(title="対象フォルダを選択")
        if folder:
            self.folder_path.set(folder)
            self.status_var.set(f"フォルダが選択されました: {folder}")
            self.update_preview()
    
    def select_output(self):
        """出力ファイルの保存先選択ダイアログを開く"""
        file_path = filedialog.asksaveasfilename(
            title="保存先を選択",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            self.output_path.set(file_path)
            self.status_var.set(f"保存先が設定されました: {file_path}")
    
    def update_preview(self):
        """ファイル一覧のプレビューを更新"""
        folder = self.folder_path.get()
        if not folder:
            messagebox.showwarning("警告", "フォルダが選択されていません。")
            return
        
        if not os.path.exists(folder):
            messagebox.showerror("エラー", "指定されたフォルダが存在しません。")
            return
        
        try:
            # 既存のアイテムをクリア
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            # ファイル一覧を取得して表示
            file_names = os.listdir(folder)
            file_names.sort()  # アルファベット順にソート
            
            for file_name in file_names:
                self.tree.insert("", tk.END, values=(file_name,))
            
            self.status_var.set(f"ファイル数: {len(file_names)}個")
            
        except PermissionError:
            messagebox.showerror("エラー", "フォルダにアクセスできません。権限を確認してください。")
        except Exception as e:
            messagebox.showerror("エラー", f"ファイル一覧の取得中にエラーが発生しました: {str(e)}")
    
    def create_excel(self):
        """Excelファイルを作成"""
        folder = self.folder_path.get()
        output = self.output_path.get()
        
        if not folder:
            messagebox.showwarning("警告", "フォルダが選択されていません。")
            return
        
        if not output:
            messagebox.showwarning("警告", "出力ファイル名が指定されていません。")
            return
        
        if not os.path.exists(folder):
            messagebox.showerror("エラー", "指定されたフォルダが存在しません。")
            return
        
        try:
            # ファイル名を取得
            file_names = os.listdir(folder)
            file_names.sort()  # アルファベット順にソート
            
            # データフレームを作成
            df = pd.DataFrame(file_names, columns=["File Name"])
            
            # Excelファイルに書き込む
            df.to_excel(output, index=False)
            
            self.status_var.set(f"完了: {len(file_names)}個のファイル名を {output} に保存しました")
            messagebox.showinfo("完了", f"ファイル名を {output} に保存しました。\nファイル数: {len(file_names)}個")
            
        except PermissionError:
            messagebox.showerror("エラー", "ファイルの保存に失敗しました。ファイルが他のアプリケーションで開かれている可能性があります。")
        except Exception as e:
            messagebox.showerror("エラー", f"Excelファイル作成中にエラーが発生しました: {str(e)}")

def main():
    root = tk.Tk()
    app = FileListApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()