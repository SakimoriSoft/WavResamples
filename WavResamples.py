import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import librosa
import soundfile as sf
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
        self.geometry("900x600")

        self.resample_task_queue = queue.Queue()
        self.resample_results_queue = queue.Queue()
        self.worker_thread = None
        self.auto_output_dir = None # 自動変換モード時の出力先
        self.last_individual_output_dir = None # 個別変換モード時の最後の出力先
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
        self.tree.column("filepath", width=433, anchor=tk.W, stretch=tk.NO) # 初期幅を調整
        self.tree.column("samplerate", width=150, minwidth=150, anchor=tk.CENTER, stretch=tk.NO)
        self.tree.column("status", width=100, minwidth=100, anchor=tk.CENTER, stretch=tk.NO)

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
        self.auto_resample_check = ttk.Checkbutton(control_frame, text="自動で変更する", variable=self.auto_resample_var, command=self.on_auto_resample_toggle)
        self.auto_resample_check.pack(side=tk.LEFT, padx=10)

        self.save_to_source_var = tk.BooleanVar(value=False)
        self.save_to_source_check = ttk.Checkbutton(control_frame, text="ソース元に保存", variable=self.save_to_source_var, command=self.on_save_to_source_toggle)
        self.save_to_source_check.pack(side=tk.LEFT, padx=5)

        self.resample_button = ttk.Button(control_frame, text="一括変換実行", command=self.start_resampling_process)
        self.resample_button.pack(side=tk.LEFT, padx=10)

        self.individual_resample_button = ttk.Button(control_frame, text="選択ファイル変換", command=self.start_selected_resampling_process, state=tk.DISABLED)
        self.individual_resample_button.pack(side=tk.LEFT, padx=5)
        
        self.clear_button = ttk.Button(control_frame, text="リストクリア", command=self.clear_list)
        self.clear_button.pack(side=tk.LEFT, padx=5)

        self.delete_button = ttk.Button(control_frame, text="選択消去", command=self.delete_selected_items, state=tk.DISABLED)
        self.delete_button.pack(side=tk.LEFT, padx=5)

        # --- ステータスバー ---
        self.status_var = tk.StringVar()
        self.status_label = ttk.Label(self, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        self.status_label.pack(side=tk.BOTTOM, fill="x", pady=(5,0))
        self.status_var.set("準備完了。WAVファイルをドラッグ＆ドロップしてください。")

        self.update_status_and_button_states() # Initialize button states and status message
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

                    if self.auto_resample_var.get(): # 自動変換モードがONの場合のみキューイング
                        output_dir_for_task = None
                        if self.save_to_source_var.get():
                            output_dir_for_task = os.path.dirname(filepath_abs)
                        else: # ソース元に保存しない場合 -> auto_output_dir を使う
                            if not self.auto_output_dir: # auto_output_dir が必須なのに未設定
                                self.status_var.set("自動変換エラー: 出力先フォルダが未指定です。")
                                self.tree.set(item_id, column="status", value="出力先未指定")
                                if self.auto_resample_var.get(): # まだONなら警告しOFFにする
                                    messagebox.showerror("自動変換エラー", "自動変換用の出力先フォルダが設定されていません。\n「自動で変更する」をOFFにするか、設定を見直してください。")
                                    self.auto_resample_var.set(False)
                                    self.update_status_and_button_states()
                                continue # このファイルのキューイングをスキップ
                            output_dir_for_task = self.auto_output_dir

                        if output_dir_for_task: # 出力先が確定した場合のみキューへ
                            try:
                                target_sr_hz, _ = self._get_target_sr_from_gui() # 現在の目標SRを取得
                                self.tree.set(item_id, column="status", value="キュー済")
                                # タスクキューには item_id, filepath_abs, target_sr_hz, output_dir_for_task, filename, original_sr を渡す
                                self.resample_task_queue.put((item_id, filepath_abs, target_sr_hz, output_dir_for_task, filename, original_sr))
                                self.status_var.set(f"キュー追加: {filename}")
                                self._ensure_worker_thread_running()
                            except ValueError as ve: # 目標SR値が無効な場合
                                 self.tree.set(item_id, column="status", value="SR値エラー")
                                 self.status_var.set(f"目標SR値エラーのためキュー追加失敗: {ve}")
                        # else の場合、output_dir_for_task が None で、既にエラー処理されているはず

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

        output_dir_for_batch = None
        if not self.save_to_source_var.get():
            output_dir_for_batch = filedialog.askdirectory(title="変換後のファイルの保存先フォルダを選択してください")
            if not output_dir_for_batch:
                self.status_var.set("保存先フォルダが選択されませんでした。処理を中止します。")
                # ボタン状態を元に戻す必要はない（この関数内で有効化されるため）
                return

        self.resample_button.config(state=tk.DISABLED)
        self.clear_button.config(state=tk.DISABLED)
        self.delete_button.config(state=tk.DISABLED) # 処理中は消去ボタンも無効化
        self.individual_resample_button.config(state=tk.DISABLED) # 処理中は個別変換ボタンも無効化
        self.auto_resample_check.config(state=tk.DISABLED) # 処理中は自動変換チェックも無効化
        self.save_to_source_check.config(state=tk.DISABLED) # ソース元保存チェックも無効化
        # self.status_var.set("変換処理中...") # 個別ファイル処理前に設定するため、ここでは不要
        # self.update_idletasks() # GUIの更新を強制

        processed_count = 0
        error_count = 0
        skipped_count = 0
        actually_converted_count = 0 # スキップを除いた実際に変換されたファイル数

        for item_id in items:
            values = self.tree.item(item_id, "values")
            filename, filepath, original_sr_str, _ = values # 既存のステータスは無視
            original_sr = int(original_sr_str)

            # 処理開始前に「処理中」に更新
            self.tree.set(item_id, column="status", value="処理中...")
            self.status_var.set(f"処理中: {filename}...")
            self.update_idletasks() # GUIを即時更新

            current_output_dir = ""
            if self.save_to_source_var.get():
                current_output_dir = os.path.dirname(filepath)
            else:
                current_output_dir = output_dir_for_batch # この時点でNoneではないはず

            result_status, message = self._perform_single_resample_logic(filepath, original_sr, target_sr, current_output_dir, filename)
            
            # 処理完了後に状態を更新
            self.tree.set(item_id, column="status", value=result_status)
            self.status_var.set(f"{filename}: {result_status} {(' - ' + message) if message and result_status != '処理中...' else ''}") # メッセージがある場合のみ表示
            self.update_idletasks() # GUIを即時更新
            
            if result_status == "処理済":
                processed_count += 1
                if "スキップ" in message: # スキップされた場合
                    skipped_count +=1
                else: # スキップされなかった場合（実際に変換された）
                    actually_converted_count +=1
            elif result_status == "エラー":
                error_count += 1
                # messagebox.showerror は最後にまとめて表示するため、ここでは表示しない方針も検討可能

        self.resample_button.config(state=tk.NORMAL)
        self.clear_button.config(state=tk.NORMAL)
        self.auto_resample_check.config(state=tk.NORMAL) # 自動変換チェックを有効化
        # self.individual_resample_button の状態は on_tree_select で更新される
        self.save_to_source_check.config(state=tk.NORMAL) # ソース元保存チェックも有効化
        self.on_tree_select() # 処理完了後、選択状態に応じて消去ボタンの状態を更新

        final_message_parts = [f"{actually_converted_count}個のファイルを変換しました。"]
        if skipped_count > 0:
             final_message_parts.append(f"{skipped_count}個スキップ。")
        if error_count > 0:
            final_message = f"処理完了。{actually_converted_count}個成功、{error_count}個エラー、{skipped_count}個スキップ。"
            messagebox.showwarning("処理完了（一部エラーあり）", final_message)
        else:
            final_message = f"処理完了。{actually_converted_count}個のファイルが正常に変換されました。"
            if skipped_count > 0:
                final_message += f" ({skipped_count}個は目標周波数と同一のためスキップ)"
            messagebox.showinfo("処理完了", final_message)
        
        self.status_var.set(final_message)

    def start_selected_resampling_process(self):
        """リストで選択されたファイルのサンプリング周波数を変換します。"""
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("情報なし", "変換対象のファイルが選択されていません。")
            self.status_var.set("選択されたファイルがありません。")
            return

        try:
            target_sr, _ = self._get_target_sr_from_gui()
        except ValueError as e:
            messagebox.showerror("入力エラー", str(e))
            self.status_var.set(str(e))
            return

        output_dir_for_selected = None
        if not self.save_to_source_var.get():
            # 前回の個別変換時の出力先を初期ディレクトリとして提案
            initial_dir = self.last_individual_output_dir if self.last_individual_output_dir else None
            output_dir_for_selected = filedialog.askdirectory(
                title="選択ファイルの保存先フォルダを選択してください",
                initialdir=initial_dir
            )
            if not output_dir_for_selected:
                self.status_var.set("保存先フォルダが選択されませんでした。処理を中止します。")
                return
            self.last_individual_output_dir = output_dir_for_selected # 記憶

        # 処理中は関連ボタンを無効化
        self.resample_button.config(state=tk.DISABLED)
        self.individual_resample_button.config(state=tk.DISABLED)
        self.clear_button.config(state=tk.DISABLED)
        self.delete_button.config(state=tk.DISABLED)
        self.auto_resample_check.config(state=tk.DISABLED)
        self.save_to_source_check.config(state=tk.DISABLED)

        processed_count = 0
        error_count = 0
        skipped_count = 0
        actually_converted_count = 0

        for item_id in selected_items:
            values = self.tree.item(item_id, "values")
            filename, filepath, original_sr_str, _ = values
            original_sr = int(original_sr_str)

            self.tree.set(item_id, column="status", value="処理中...")
            self.status_var.set(f"処理中: {filename}...")
            self.update_idletasks()

            current_output_dir = os.path.dirname(filepath) if self.save_to_source_var.get() else output_dir_for_selected

            result_status, message = self._perform_single_resample_logic(filepath, original_sr, target_sr, current_output_dir, filename)
            
            self.tree.set(item_id, column="status", value=result_status)
            self.status_var.set(f"{filename}: {result_status} {(' - ' + message) if message and result_status != '処理中...' else ''}")
            self.update_idletasks()

            if result_status == "処理済":
                processed_count += 1
                if "スキップ" in message:
                    skipped_count +=1
                else:
                    actually_converted_count +=1
            elif result_status == "エラー":
                error_count += 1

        # 処理完了後、ボタン状態を元に戻す
        self.auto_resample_check.config(state=tk.NORMAL)
        self.save_to_source_check.config(state=tk.NORMAL)
        self.clear_button.config(state=tk.NORMAL) # リストクリアボタンを有効に戻す
        self.update_status_and_button_states() # これで一括変換ボタンなども適切に更新される
        self.on_tree_select() # 選択消去ボタンと個別変換ボタンの状態を更新

        final_message = f"選択ファイル処理完了。{actually_converted_count}個成功、{error_count}個エラー、{skipped_count}個スキップ。"
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
        """Treeviewのアイテム選択状態に応じて「選択消去」「選択ファイル変換」ボタンの有効/無効を切り替えます。"""
        has_selection = bool(self.tree.selection())
        auto_mode = self.auto_resample_var.get()

        if has_selection:
            self.delete_button.config(state=tk.NORMAL)
            self.individual_resample_button.config(state=tk.NORMAL if not auto_mode else tk.DISABLED)
        else:
            self.delete_button.config(state=tk.DISABLED)
            self.individual_resample_button.config(state=tk.DISABLED)

    def on_auto_resample_toggle(self):
        """「自動で変更する」チェックボックスの状態変更時の処理。"""
        auto_mode_now = self.auto_resample_var.get()
        save_to_source = self.save_to_source_var.get()

        if auto_mode_now and not save_to_source: # 自動ON かつ ソース保存OFF の場合
            if not self.auto_output_dir:
                messagebox.showinfo("出力先指定", "自動変換用の出力先フォルダを指定してください。\n「ソース元に保存」がOFFのため、出力先が必要です。")
                new_dir = filedialog.askdirectory(title="自動変換ファイルの保存先フォルダを選択")
                if new_dir:
                    self.auto_output_dir = new_dir
                else:
                    self.auto_resample_var.set(False) # 指定がなければ自動変換をOFFに戻す
                    messagebox.showwarning("出力先未指定", "出力先が指定されなかったため、自動変換をOFFにしました。")
        self.update_status_and_button_states()

    def on_save_to_source_toggle(self):
        """「ソース元に保存」チェックボックスの状態変更時の処理。"""
        auto_mode = self.auto_resample_var.get()
        save_to_source_now = self.save_to_source_var.get()

        if auto_mode and not save_to_source_now: # 自動ON かつ ソース保存OFF になった/なっている場合
             if not self.auto_output_dir:
                messagebox.showinfo("出力先指定", "「ソース元に保存」がOFFのため、自動変換用の出力先フォルダを指定してください。")
                new_dir = filedialog.askdirectory(title="自動変換ファイルの保存先フォルダを選択")
                if new_dir:
                    self.auto_output_dir = new_dir
                else:
                    self.auto_resample_var.set(False) # ソース保存OFFで出力先も未指定なら自動変換もOFF
                    messagebox.showwarning("出力先未指定", "出力先が指定されなかったため、自動変換をOFFにしました。")
        self.update_status_and_button_states()

    def update_status_and_button_states(self):
        """現在のモード設定に基づいてUI（ボタン状態、ステータスメッセージ）を更新する。"""
        auto_mode = self.auto_resample_var.get()
        save_to_source = self.save_to_source_var.get()

        # 処理中でなければボタンの状態を更新
        # (処理中は start_resampling_process や _worker_resample_files で直接制御)
        is_processing_manually = self.resample_button['state'] == tk.DISABLED and not auto_mode
        
        if not is_processing_manually: # 手動処理中でない場合のみボタン状態を更新
            if auto_mode:
                self.resample_button.config(state=tk.DISABLED)
                self.individual_resample_button.config(state=tk.DISABLED) # 自動モード中は個別変換も不可
            else:
                self.resample_button.config(state=tk.NORMAL)
                # 個別変換ボタンの状態は on_tree_select で選択状態に応じて更新される
                self.on_tree_select() # 自動モードOFFになったら選択状態を再評価

        # ステータスメッセージの更新ロジック
        if auto_mode:
            self.resample_button.config(state=tk.DISABLED)
            if save_to_source:
                self.status_var.set("自動変換 ON (ソース元へ保存)。ファイルドロップで自動処理。")
            else: # 自動ON、ソース保存OFF
                if not self.auto_output_dir:
                    self.status_var.set("自動変換 ON (出力先未指定)。設定を確認してください。")
                else:
                     self.status_var.set(f"自動変換 ON (出力先: {os.path.basename(self.auto_output_dir) if self.auto_output_dir else '未指定'})。ファイルドロップで自動処理。")
            self._ensure_worker_thread_running()
        else: # 自動変換OFF
            # 手動処理中でなければステータスを更新
            if not is_processing_manually:
                self.resample_button.config(state=tk.NORMAL) # is_processing_manually でなければ NORMAL に戻す
                # self.individual_resample_button の状態は on_tree_select で制御
                if save_to_source:
                    self.status_var.set("手動変換 (ソース元へ保存)。「一括変換実行」ボタンで処理。")
                else: # 自動OFF、ソース保存OFF
                    self.status_var.set("手動変換 (指定フォルダへ保存)。「一括変換実行」ボタンで処理。")

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
            item_id = None # エラーハンドリングのために初期化
            try:
                # タスクキューからアイテムを取得 (item_id, filepath, target_sr, output_dir, filename, original_sr)
                item_id, filepath, target_sr, output_dir, filename, original_sr = self.resample_task_queue.get(timeout=1)
                
                # GUIに「処理中」を通知
                self.resample_results_queue.put((item_id, "処理中...", None)) # メッセージはNone

                result_status, message = self._perform_single_resample_logic(filepath, original_sr, target_sr, output_dir, filename)
                self.resample_results_queue.put((item_id, result_status, message))
                self.resample_task_queue.task_done()
            except queue.Empty:
                continue # タイムアウト、キューが空ならループを継続
            except Exception as e:
                print(f"ワーカースレッドで予期せぬエラー: {e}")
                if item_id: # item_idが取得できていればエラーを通知
                   self.resample_results_queue.put((item_id, "エラー", str(e)))
                # item_idがNoneの場合（キュー取得前など）は、特定アイテムのエラーとして通知できない
        print("ワーカースレッドを終了します。")

    def _perform_single_resample_logic(self, filepath, original_sr, target_sr, output_dir, filename):
        """単一ファイルのサンプリング周波数変換処理を実行し、結果ステータスとメッセージを返す。"""
        try:
            # original_sr は引数で渡されるようになったので、sf.infoの再呼び出しは不要
            if original_sr == target_sr:
                msg = f"スキップ: {filename} (既に目標サンプリング周波数です)"
                # print(msg) # ログ出力は呼び出し元や専用ロガーで行う方が良い場合もある
                return "処理済", msg

            y, sr_librosa_original = librosa.load(filepath, sr=None) # 元のSRでロード
            y_resampled = librosa.resample(y=y, orig_sr=sr_librosa_original, target_sr=target_sr)
            base, ext = os.path.splitext(filename)
            output_filename = f"{base}_resampled_{target_sr}Hz{ext}"
            output_path = os.path.join(output_dir, output_filename)
            
            # 出力先ディレクトリが存在しない場合は作成
            if not os.path.exists(output_dir):
                try:
                    os.makedirs(output_dir)
                    print(f"作成された出力ディレクトリ: {output_dir}")
                except OSError as ose:
                    # ディレクトリ作成失敗時のエラーハンドリング
                    error_msg = f"エラー: 出力ディレクトリの作成に失敗しました ({output_dir}) - {ose}"
                    print(error_msg)
                    return "エラー", error_msg # ディレクトリ作成失敗もエラーとして返す

            sf.write(output_path, y_resampled, target_sr, subtype='PCM_16')
            success_msg = f"変換成功: {output_filename}"
            # print(success_msg)
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
                    if status == "処理中...":
                        self.status_var.set(f"{filename_in_tree}: 処理中...")
                    else:
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
