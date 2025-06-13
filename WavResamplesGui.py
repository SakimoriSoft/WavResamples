import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import librosa
import soundfile as sf # pysomefile は soundfile のことと解釈します
import os
import threading
import queue

# tkinterdnd2 が利用可能か最初に確認します
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:
    # GUI表示前にエラーを出すため、Tkinterのmessageboxは使わずprintとexitで対応
    print("エラー: 必須ライブラリ tkinterdnd2 が見つかりません。")
    print("ターミナルで次のようにインストールしてください: pip install tkinterdnd2")
    import sys
    sys.exit(1)


class AudioResamplerApp(TkinterDnD.Tk): # DND機能のためにTkinterDnD.Tkを継承
    def __init__(self):
        super().__init__()
        self.title("WAVサンプリング周波数一括変更ツール")
        self.geometry("850x600")

        self.resample_task_queue = queue.Queue()
        self.resample_results_queue = queue.Queue()
        self.worker_thread = None
        self.auto_output_dir = None
        self.is_shutting_down = False

        self._setup_ui()

    def _setup_ui(self):
        # --- ファイルリストフレーム ---
        list_frame = ttk.LabelFrame(self, text="ファイルリスト (WAVファイルをここにドラッグ＆ドロップ)")
        list_frame.pack(padx=10, pady=10, fill="both", expand=True)

        # Treeview (多列リストボックスとして使用)
        columns = ("filename", "filepath", "samplerate", "status")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings")
        self.tree.heading("filename", text="ファイル名")
        self.tree.heading("filepath", text="ファイルパス")
        self.tree.heading("samplerate", text="サンプリング周波数 (Hz)")
        self.tree.heading("status", text="状態")

        self.tree.column("filename", width=180, anchor=tk.W, stretch=tk.NO)
        self.tree.column("filepath", width=370, anchor=tk.W, stretch=tk.NO)
        self.tree.column("samplerate", width=150, anchor=tk.CENTER, stretch=tk.NO)
        self.tree.column("status", width=100, anchor=tk.CENTER, stretch=tk.NO)

        # スクロールバー
        scrollbar_y = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(list_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        
        scrollbar_y.pack(side=tk.RIGHT, fill="y")
        scrollbar_x.pack(side=tk.BOTTOM, fill="x")
        self.tree.pack(side=tk.LEFT, fill="both", expand=True)


        # ドラッグ＆ドロップ設定
        self.tree.drop_target_register(DND_FILES)
        self.tree.dnd_bind('<<Drop>>', self.handle_drop)
        self.tree.bind('<<TreeviewSelect>>', self.on_tree_select) # 選択変更イベント

        # --- コントロールフレーム ---
        control_frame = ttk.Frame(self)
        control_frame.pack(padx=10, pady=(0, 5), fill="x")

        ttk.Label(control_frame, text="目標サンプリング周波数:").pack(side=tk.LEFT, padx=(0,5))
        self.target_sr_var = tk.StringVar(value="44100")
        self.target_sr_entry = ttk.Entry(control_frame, textvariable=self.target_sr_var, width=10)
        self.target_sr_entry.pack(side=tk.LEFT, padx=5)

        self.unit_var = tk.StringVar(value="Hz")
        self.unit_combobox = ttk.Combobox(control_frame, textvariable=self.unit_var, values=["Hz", "kHz"], width=5, state="readonly")
        self.unit_combobox.pack(side=tk.LEFT, padx=(0,10))
        self.unit_combobox.current(0) # "Hz" をデフォルト選択

        self.auto_resample_var = tk.BooleanVar(value=False)
        self.auto_resample_check = ttk.Checkbutton(control_frame, text="自動で変更する", variable=self.auto_resample_var, command=self.toggle_auto_resample_mode)
        self.auto_resample_check.pack(side=tk.LEFT, padx=10)

        self.resample_button = ttk.Button(control_frame, text="一括変換実行", command=self.start_resampling_process)
        self.resample_button.pack(side=tk.LEFT, padx=10)
        
        self.clear_button = ttk.Button(control_frame, text="リストクリア", command=self.clear_list)
        self.clear_button.pack(side=tk.LEFT, padx=5)

        self.delete_button = ttk.Button(control_frame, text="選択消去", command=self.delete_selected_items, state=tk.DISABLED)
        self.delete_button.pack(side=tk.LEFT, padx=5)

        # --- ステータスバー ---
        self.status_var = tk.StringVar()
        self.status_label = ttk.Label(self, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        self.status_label.pack(side=tk.BOTTOM, fill="x", pady=(5,0))
        self.status_var.set("準備完了。WAVファイルをドラッグ＆ドロップしてください。")

        self.toggle_auto_resample_mode() # Initialize button states based on checkbox
        self.after(100, self.process_resample_results) # Start polling for results
        self.protocol("WM_DELETE_WINDOW", self.on_closing) # Handle window close

    def handle_drop(self, event):
        try:
            # event.data はTclリスト形式のファイルパス文字列
            raw_paths = self.tk.splitlist(event.data)
            
            files_to_add = []
            for path_str in raw_paths:
                # TkinterDnDからのパスは通常既に正規化されている
                if os.path.isfile(path_str): # 実際にファイルか確認
                    files_to_add.append(path_str)
            
            if not files_to_add:
                self.status_var.set("有効なファイルパスがドロップされませんでした。")
                return

            added_count = 0
            skipped_non_wav = 0
            skipped_duplicate = 0

            for file_path in files_to_add:
                if not file_path.lower().endswith((".wav", ".wave")):
                    skipped_non_wav += 1
                    continue

                filename = os.path.basename(file_path)
                # 重複チェックのために絶対パスを使用
                filepath_abs = os.path.abspath(file_path) 

                is_duplicate = False
                for item_id_check in self.tree.get_children(): # Renamed item_id to avoid conflict
                    # valuesのインデックス1がファイルパス
                    if self.tree.item(item_id_check, "values")[1] == filepath_abs:
                        is_duplicate = True
                        skipped_duplicate += 1
                        break
                
                if is_duplicate:
                    continue

                try:
                    # soundfile.infoでサンプリング周波数のみ取得 (librosa.loadより軽量)
                    info = sf.info(filepath_abs)
                    original_sr = info.samplerate
                    # Treeviewにアイテムを追加し、そのIDを取得
                    item_id = self.tree.insert("", tk.END, values=(filename, filepath_abs, original_sr, "")) # 初期状態は空
                    added_count += 1

                    if self.auto_resample_var.get(): # 自動変換モードがONの場合
                        if not self.auto_output_dir:
                            messagebox.showinfo("出力先指定", "自動変換用の出力先フォルダを最初に指定してください。")
                            self.auto_output_dir = filedialog.askdirectory(title="自動変換ファイルの保存先フォルダを選択")
                            if not self.auto_output_dir:
                                self.status_var.set("自動変換の出力先が未指定のため、キューに追加できませんでした。")
                                self.tree.set(item_id, column="status", value="出力先未指定")
                                continue # このファイルのキューイングをスキップ
                            else:
                                self.status_var.set(f"自動変換の出力先: {self.auto_output_dir}")
                        
                        try:
                            target_sr_hz, _ = self._get_target_sr_from_gui() # 現在の目標SRを取得
                            self.tree.set(item_id, column="status", value="キュー済")
                            # タスクキューには item_id, filepath_abs, target_sr_hz, output_dir, filename, original_sr を渡す
                            self.resample_task_queue.put((item_id, filepath_abs, target_sr_hz, self.auto_output_dir, filename, original_sr))
                            self._ensure_worker_thread_running()
                        except ValueError as ve: # 目標SR値が無効な場合
                             self.tree.set(item_id, column="status", value="SR値エラー")
                             self.status_var.set(f"目標SR値エラーのためキュー追加失敗: {ve}")
                except Exception as e:
                    self.status_var.set(f"エラー: {filename} の情報取得失敗 - {e}")
                    print(f"Error getting info for {filepath_abs}: {e}")
            
            messages = []
            if added_count > 0:
                messages.append(f"{added_count} 個のWAVファイルを追加しました。")
            if skipped_non_wav > 0:
                messages.append(f"{skipped_non_wav} 個の非WAVファイルを無視しました。")
            if skipped_duplicate > 0:
                messages.append(f"{skipped_duplicate} 個の重複ファイルを無視しました。")
            
            if not messages and files_to_add:
                 self.status_var.set("追加可能なWAVファイルが見つかりませんでした。")
            elif messages:
                self.status_var.set(" ".join(messages))
            # files_to_addが空の場合は最初のifで捕捉される

        except Exception as e:
            self.status_var.set(f"ドロップ処理エラー: {e}")
            messagebox.showerror("ドロップエラー", f"ファイルの処理中に予期せぬエラーが発生しました: {e}")

        self.on_tree_select() # ドロップ後、選択状態が変わる可能性があるのでボタン状態更新

    def _get_target_sr_from_gui(self):
        """GUIから目標サンプリング周波数を読み取り、Hz単位の整数と元の入力文字列を返す。"""
        target_sr_input_str = self.target_sr_var.get()
        try:
            target_sr_input = float(target_sr_input_str)
        except ValueError:
            raise ValueError("目標サンプリング周波数には有効な数値を入力してください。")
            
        selected_unit = self.unit_var.get()
        target_sr_hz = 0

        if selected_unit == "kHz":
            target_sr_hz = int(target_sr_input * 1000)
        elif selected_unit == "Hz":
            target_sr_hz = int(target_sr_input)
        else: # 通常は発生しない
            raise ValueError("無効な単位が選択されています。")

        if target_sr_hz <= 0:
            raise ValueError("目標サンプリング周波数は正の整数である必要があります。")
        return target_sr_hz, target_sr_input_str

    def clear_list(self):
        """リストビューの内容をすべてクリアします。"""
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.status_var.set("ファイルリストがクリアされました。")
        self.on_tree_select() # クリア後は何も選択されていないのでボタン状態更新

    def start_resampling_process(self):
        """リスト内のファイルのサンプリング周波数を一括変換します。"""
        items = self.tree.get_children()
        if not items:
            messagebox.showwarning("情報なし", "変換対象のファイルがリストにありません。")
            self.status_var.set("リストにファイルがありません。")
            return

        try:
            target_sr, _ = self._get_target_sr_from_gui()
        except ValueError as e:
            messagebox.showerror("入力エラー", str(e))
            self.status_var.set(str(e))
            return

        output_dir = filedialog.askdirectory(title="変換後のファイルの保存先フォルダを選択してください")
        if not output_dir:
            self.status_var.set("保存先フォルダが選択されませんでした。処理を中止します。")
            return

        self.resample_button.config(state=tk.DISABLED)
        self.clear_button.config(state=tk.DISABLED)
        self.delete_button.config(state=tk.DISABLED) # 処理中は消去ボタンも無効化
        self.auto_resample_check.config(state=tk.DISABLED) # 処理中は自動変換チェックも無効化
        self.status_var.set("変換処理中...")
        self.update_idletasks() # GUIの更新を強制

        processed_count = 0
        error_count = 0
        skipped_count = 0

        for item_id in items:
            values = self.tree.item(item_id, "values")
            filename, filepath, original_sr_str, _ = values # 既存のステータスは無視
            original_sr = int(original_sr_str)

            self.status_var.set(f"処理中: {filename}...")
            self.update_idletasks()

            # バッチモードでは同期的に処理し、ファイルごとにUIを更新
            result_status, message = self._perform_single_resample_logic(filepath, original_sr, target_sr, output_dir, filename)
            self.tree.set(item_id, column="status", value=result_status)
            if result_status == "処理済":
                processed_count += 1
                if "スキップ" in message: # スキップされた場合
                    skipped_count +=1
            elif result_status == "エラー":
                error_count += 1
                messagebox.showerror("変換エラー", f"{filename} の変換中にエラーが発生しました:\n{message}")

        self.resample_button.config(state=tk.NORMAL)
        self.clear_button.config(state=tk.NORMAL)
        self.auto_resample_check.config(state=tk.NORMAL) # 自動変換チェックを有効化
        self.on_tree_select() # 処理完了後、選択状態に応じて消去ボタンの状態を更新

        final_message_parts = [f"{processed_count}個のファイルを処理しました。"]
        if skipped_count > 0:
             final_message_parts.append(f"({skipped_count}個スキップ)")
        if error_count > 0:
            final_message = f"処理完了。{processed_count-error_count-skipped_count}個成功、{error_count}個エラー、{skipped_count}個スキップ。"
            messagebox.showwarning("処理完了（一部エラーあり）", final_message)
        else:
            final_message = f"処理完了。全てのファイル ({processed_count-skipped_count}個) が正常に変換されました。"
            if skipped_count > 0:
                final_message += f" ({skipped_count}個は既に目標周波数だったためスキップ)"
            messagebox.showinfo("処理完了", final_message)
        
        self.status_var.set(final_message)

    def delete_selected_items(self):
        """Treeviewで選択されているアイテムを削除します。"""
        selected_items = self.tree.selection()
        if not selected_items:
            self.status_var.set("消去するアイテムが選択されていません。")
            return

        for item_id in selected_items:
            self.tree.delete(item_id)
        
        self.status_var.set(f"{len(selected_items)} 個のアイテムをリストから消去しました。")
        self.on_tree_select() # 削除後、選択状態が変わるのでボタン状態更新

    def on_tree_select(self, event=None):
        """Treeviewのアイテム選択状態に応じて「選択消去」ボタンの有効/無効を切り替えます。"""
        if self.tree.selection():
            self.delete_button.config(state=tk.NORMAL)
        else:
            self.delete_button.config(state=tk.DISABLED)

    def toggle_auto_resample_mode(self):
        """自動変換モードのON/OFFを切り替え、関連するUIの状態を更新する。"""
        if self.auto_resample_var.get():
            self.resample_button.config(state=tk.DISABLED)
            self.status_var.set("自動変換モード ON。ファイルドロップで自動処理します。")
            if not self.auto_output_dir:
                messagebox.showinfo("出力先指定", "自動変換用の出力先フォルダを最初に指定してください。")
                self.auto_output_dir = filedialog.askdirectory(title="自動変換ファイルの保存先フォルダを選択")
                if not self.auto_output_dir:
                    self.status_var.set("自動変換の出力先が未指定です。ファイル追加時に再度確認します。")
                    # 自動変換をOFFに戻すか、ユーザーに再度促すかなどの対応も検討可能
                    # self.auto_resample_var.set(False) # 例: OFFに戻す
                    # self.toggle_auto_resample_mode() # UI状態を再更新
                else:
                    self.status_var.set(f"自動変換の出力先: {self.auto_output_dir}")
            self._ensure_worker_thread_running()
        else:
            self.resample_button.config(state=tk.NORMAL)
            self.status_var.set("自動変換モード OFF。「一括変換実行」ボタンで処理します。")
            # ここでキューをクリアしたり、ワーカースレッドに停止を指示するロジックも検討可能

    def _ensure_worker_thread_running(self):
        """ワーカースレッドが実行中でなければ起動する。"""
        if self.worker_thread is None or not self.worker_thread.is_alive():
            self.worker_thread = threading.Thread(target=self._worker_resample_files, daemon=True)
            self.worker_thread.start()
            print("ワーカースレッドを開始しました。")

    def _worker_resample_files(self):
        """ワーカースレッドのメインループ。タスクキューからファイル処理タスクを取得して実行する。"""
        print("ワーカースレッド実行中...")
        while not self.is_shutting_down:
            try:
                # タスクキューからアイテムを取得 (item_id, filepath, target_sr, output_dir, filename, original_sr)
                item_id, filepath, target_sr, output_dir, filename, original_sr = self.resample_task_queue.get(timeout=1)
                
                # GUIに「処理中」を通知 (結果キュー経由が望ましいが、簡略化のため直接更新も検討)
                # self.resample_results_queue.put((item_id, "処理中...", "")) # より良い方法

                result_status, message = self._perform_single_resample_logic(filepath, original_sr, target_sr, output_dir, filename)
                self.resample_results_queue.put((item_id, result_status, message))
                self.resample_task_queue.task_done()
            except queue.Empty:
                continue # タイムアウト、キューが空ならループを継続
            except Exception as e:
                print(f"ワーカースレッドで予期せぬエラー: {e}")
                # item_idが取得できていれば、そのアイテムのエラーとして結果キューに通知することも可能
                # if 'item_id' in locals(): # Check if item_id was assigned
                #    self.resample_results_queue.put((item_id, "エラー", str(e)))
        print("ワーカースレッドを終了します。")

    def _perform_single_resample_logic(self, filepath, original_sr, target_sr, output_dir, filename):
        """単一ファイルのサンプリング周波数変換処理を実行し、結果ステータスとメッセージを返す。"""
        try:
            # original_sr は引数で渡されるようになったので、sf.infoの再呼び出しは不要
            if original_sr == target_sr:
                msg = f"スキップ: {filename} (既に目標サンプリング周波数です)"
                print(msg)
                return "処理済", msg

            y, sr_librosa_original = librosa.load(filepath, sr=None) # 元のSRでロード
            y_resampled = librosa.resample(y=y, orig_sr=sr_librosa_original, target_sr=target_sr)
            base, ext = os.path.splitext(filename)
            output_filename = f"{base}_resampled_{target_sr}Hz{ext}"
            output_path = os.path.join(output_dir, output_filename)
            sf.write(output_path, y_resampled, target_sr, subtype='PCM_16')
            success_msg = f"変換成功: {output_filename}"
            print(success_msg)
            return "処理済", success_msg
        except Exception as e:
            error_msg = f"エラー: {filename} の変換に失敗 - {e}"
            print(error_msg)
            return "エラー", str(e) # エラーメッセージ全体を返す

    def process_resample_results(self):
        """結果キューをポーリングし、GUIを更新する。"""
        try:
            while not self.resample_results_queue.empty():
                item_id, status, message = self.resample_results_queue.get_nowait()
                if self.tree.exists(item_id): # アイテムがまだ存在するか確認
                    self.tree.set(item_id, column="status", value=status)
                    # ステータスバーにはファイル名と結果を表示
                    filename_in_tree = self.tree.item(item_id, "values")[0]
                    self.status_var.set(f"{filename_in_tree}: {status} {(' - ' + message) if message else ''}")
                self.resample_results_queue.task_done()
        finally:
            if not self.is_shutting_down:
                self.after(100, self.process_resample_results) # 100ms後に再度実行

    def on_closing(self):
        """ウィンドウクローズ時の処理。"""
        if messagebox.askokcancel("終了確認", "アプリケーションを終了しますか？"):
            self.is_shutting_down = True
            print("シャットダウン処理を開始します...")
            # ワーカースレッドに終了を待つ (より堅牢な停止メカニズムも検討可能)
            if self.worker_thread and self.worker_thread.is_alive():
                print("ワーカースレッドの終了を待機中...")
                # タスクキューが空になるまで待つか、タイムアウトを設定
                # self.resample_task_queue.join() # 全タスク完了を待つ場合
                self.worker_thread.join(timeout=2.0) # 最大2秒待つ
                if self.worker_thread.is_alive():
                    print("ワーカースレッドがタイムアウト後も実行中です。")
                else:
                    print("ワーカースレッドは正常に終了しました。")
            self.destroy()

if __name__ == "__main__":
    app = AudioResamplerApp()
    app.mainloop()
