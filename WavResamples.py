import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import librosa # オーディオ処理ライブラリ
import soundfile as sf
import os
import threading
import queue
import numpy as np

# tkinterdnd2 が利用可能か最初に確認します
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:
    # GUI表示前にエラーを出すため、Tkinterのmessageboxは使わずprintとexitで対応
    print("エラー: 必須ライブラリ tkinterdnd2 が見つかりません。")
    print("ターミナルで次のようにインストールしてください: pip install tkinterdnd2")
    import sys
    sys.exit(1)


class AudioResamplerApp(TkinterDnD.Tk): # ドラッグ＆ドロップ機能のためにTkinterDnD.Tkを継承
    def __init__(self):
        """アプリケーションのメインクラスを初期化します。

        ウィンドウのタイトル、サイズ、および変換タスクを管理するための
        キューやスレッドなどの内部変数をセットアップします。
        """
        super().__init__()
        self.title("WAVサンプリング周波数・ステレオ・ビット深度変換ツール")

        # --- テーマ設定 ---
        self.style = ttk.Style(self)
        self.style.theme_use('clam') # カスタマイズしやすいテーマを選択
        self.dark_mode_var = tk.BooleanVar(value=True) # デフォルトはダークモード

        # ウィンドウサイズ
        window_width = 1280
        window_height = 700

        # 画面のサイズを取得
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        # ウィンドウの位置を画面の中央に設定
        center_x = (screen_width - window_width) // 2
        center_y = (screen_height - window_height) // 2

        # ウィンドウを一度仮の位置(0,0)に配置して、タイトルバーの高さを取得
        self.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
        self.update_idletasks()
        # y=0に配置した際のクライアント領域のy座標が、タイトルバーの高さに相当
        title_bar_height = self.winfo_rooty() - self.winfo_y()

        # タイトルバーを含むウィンドウ全体の高さを考慮して中央座標を計算
        window_total_height = window_height + title_bar_height

        # 計算された中央座標にウィンドウを配置
        self.geometry(f'{window_width}x{window_height}+{center_x}+{center_y - title_bar_height}')

        self.resample_task_queue = queue.Queue()
        self.resample_results_queue = queue.Queue()
        self.worker_thread = None
        self.auto_output_dir = None # 自動変換モード時の出力先
        self.last_individual_output_dir = None # 個別変換モード時の最後の出力先
        self.is_shutting_down = False # アプリケーション終了処理中フラグ
        self._is_resizing_column = False # カラムリサイズ中フラグ
        self._process_timer_id = None # 結果ポーリング用のタイマーID

        self._setup_ui() # UIのセットアップ
        self._apply_theme() # 初期テーマを適用


    # UI要素のセットアップを行うメソッド
    def _setup_ui(self):
        """ユーザーインターフェース（UI）のすべてのウィジェットをセットアップします。

        ファイルリスト表示用のTreeview、各種設定用のコンボボックスやチェックボックス、
        操作ボタン、ステータスバーなどをウィンドウ上に配置し、
        イベントハンドラをバインドします。
        """
        # --- ファイルリストフレーム ---
        list_frame = ttk.LabelFrame(self, text="ファイルリスト (WAVファイルをここにドラッグ＆ドロップ)")
        list_frame.pack(padx=10, pady=10, fill="both", expand=True)

        # list_frame 内で grid を使用してウィジェットを配置
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)

        # Treeview (多列リストボックスとして使用)
        columns = ("filename", "filepath", "samplerate", "channels", "bitdepth", "status") # bitdepth列を追加
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings")
        self.tree.heading("filename", text="ファイル名")
        self.tree.heading("filepath", text="ファイルパス")
        self.tree.heading("samplerate", text="サンプリング周波数 (Hz)")
        self.tree.heading("channels", text="チャンネル数") # 新しい列のヘッダー
        self.tree.heading("bitdepth", text="ビット深度") # 新しい列のヘッダー
        self.tree.heading("status", text="状態")

        self.tree.column("filename", width=250, anchor=tk.W, stretch=tk.NO) # 幅調整
        self.tree.column("filepath", width=540, minwidth=200, anchor=tk.W, stretch=tk.NO) # 自動伸縮を無効化
        self.tree.column("samplerate", width=140, minwidth=140, anchor=tk.CENTER, stretch=tk.NO) # 幅調整
        self.tree.column("channels", width=90, minwidth=90, anchor=tk.CENTER, stretch=tk.NO) # 新しい列の幅
        self.tree.column("bitdepth", width=100, minwidth=100, anchor=tk.CENTER, stretch=tk.NO) # 新しい列の幅
        self.tree.column("status", width=100, minwidth=100, anchor=tk.CENTER, stretch=tk.NO)

        # スクロールバー
        scrollbar_y = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        self.scrollbar_y = scrollbar_y # インスタンス変数として保持
        self.scrollbar_x = ttk.Scrollbar(list_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=self.scrollbar_x.set)

        # grid を使用して配置
        self.tree.grid(row=0, column=0, sticky="nsew")
        self.scrollbar_y.grid(row=0, column=1, sticky="ns")
        # 水平スクロールバーは _update_horizontal_scrollbar で動的に配置する

        # ドラッグ＆ドロップ設定
        self.tree.drop_target_register(DND_FILES)
        self.tree.dnd_bind('<<Drop>>', self.handle_drop)
        self.tree.bind('<<TreeviewSelect>>', self.on_tree_select) # 選択変更イベント
        self.tree.bind('<Map>', self._on_map) # 初回表示時に一度だけ実行
        self.tree.bind('<Configure>', self._update_horizontal_scrollbar) # ウィジェットサイズ変更時にスクロールバーを更新

        # --- コントロールフレーム ---
        control_frame = ttk.Frame(self)
        control_frame.pack(padx=10, pady=(0, 5), fill="x")

        ttk.Label(control_frame, text="目標サンプリング周波数:").pack(side=tk.LEFT, padx=(0,5))
        self.target_sr_var = tk.StringVar(value="44.1 kHz") # デフォルトを設定
        sr_values = ["22.05 kHz", "24 kHz", "32 kHz", "44.1 kHz", "48 kHz"]
        self.target_sr_combobox = ttk.Combobox(control_frame, textvariable=self.target_sr_var, values=sr_values, width=10, state="readonly")
        self.target_sr_combobox.pack(side=tk.LEFT, padx=(0, 10))

        # ターゲットビット深度の選択肢を追加
        ttk.Label(control_frame, text="目標ビット深度:").pack(side=tk.LEFT, padx=(10,5))
        self.target_bit_depth_var = tk.StringVar(value="16bit (PCM_16)")
        self.target_bit_depth_combobox = ttk.Combobox(control_frame, textvariable=self.target_bit_depth_var,
                                                      values=["16bit (PCM_16)", "8bit (PCM_S8)"], width=15, state="readonly")
        self.target_bit_depth_combobox.pack(side=tk.LEFT, padx=(0,10))
        self.target_bit_depth_combobox.current(0)

        # --- 各種操作ボタン ---
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

        # --- テーマ切り替え ---
        # 右端に配置
        self.theme_toggle_check = ttk.Checkbutton(
            control_frame,
            text="ダークモード",
            variable=self.dark_mode_var,
            command=self._toggle_theme
        )
        self.theme_toggle_check.pack(side=tk.RIGHT, padx=10)

        # --- ステータスバー ---
        self.status_var = tk.StringVar()
        # reliefをFLATに変更し、背景色を少し変える
        self.status_label = ttk.Label(self, textvariable=self.status_var, relief=tk.FLAT, anchor=tk.W)
        self.status_label.pack(side=tk.BOTTOM, fill="x", pady=(5,0), ipady=2) # ipadyで少し高さを出す
        self.status_var.set("準備完了。WAVファイルをドラッグ＆ドロップしてください。")

        # self.update_status_and_button_states() # 初期状態は「準備完了」メッセージのままにするため、ここでは呼ばない
        self._process_timer_id = self.after(100, self.process_resample_results) # 変換結果キューのポーリングを開始
        self.protocol("WM_DELETE_WINDOW", self.on_closing) # ウィンドウを閉じる際の処理を登録

    def _toggle_theme(self):
        """テーマを切り替えてUIに適用します。"""
        self._apply_theme()

    def _apply_theme(self):
        """現在のテーマ設定をUI全体に適用します。"""
        is_dark = self.dark_mode_var.get()

        if is_dark:
            # ダークテーマのカラーパレット
            bg_color = '#2E2E2E'
            fg_color = '#EAEAEA'
            select_bg = '#4A4A4A'
            entry_bg = '#3C3C3C'
            button_bg = '#5A5A5A'
            button_active_bg = '#6A6A6A'
            disabled_fg = '#888888'
            status_bar_bg = '#3C3C3C'
        else:
            # ライトテーマのカラーパレット
            bg_color = '#F0F0F0'
            fg_color = '#000000'
            select_bg = '#B4D5FF'
            entry_bg = '#FFFFFF'
            button_bg = '#E1E1E1'
            button_active_bg = '#C0C0C0'
            disabled_fg = '#A0A0A0'
            status_bar_bg = '#E5E5E5'

        # ウィンドウ自体の背景色
        self.configure(bg=bg_color)

        # グローバルなウィジェットスタイルの設定
        self.style.configure('.', background=bg_color, foreground=fg_color, borderwidth=1)
        self.style.configure('TFrame', background=bg_color)
        self.style.configure('TLabel', background=bg_color, foreground=fg_color)
        self.style.configure('TLabelFrame', background=bg_color, bordercolor=fg_color)
        self.style.configure('TLabelFrame.Label', background=bg_color, foreground=fg_color)

        # ボタンのスタイル
        self.style.configure('TButton', background=button_bg, foreground=fg_color, borderwidth=1, focusthickness=3, focuscolor='none')
        self.style.map('TButton',
            background=[('active', button_active_bg), ('disabled', '#4A4A4A' if is_dark else '#D0D0D0')],
            foreground=[('disabled', disabled_fg)])

        # チェックボタンのスタイル
        self.style.configure('TCheckbutton', background=bg_color, foreground=fg_color)
        self.style.map('TCheckbutton',
            background=[('active', bg_color)],
            indicatorcolor=[('selected', fg_color), ('!selected', disabled_fg)],
            foreground=[('disabled', disabled_fg)])

        # Treeviewのスタイル
        self.style.configure('Treeview', background=entry_bg, fieldbackground=entry_bg, foreground=fg_color, rowheight=25)
        self.style.map('Treeview', background=[('selected', select_bg)], foreground=[('selected', fg_color)])
        self.style.configure("Treeview.Heading", background=button_bg, foreground=fg_color, relief="flat")
        self.style.map("Treeview.Heading", background=[('active', button_active_bg)])

        # スクロールバーのスタイル
        self.style.configure('Vertical.TScrollbar', background=button_bg, troughcolor=bg_color, arrowcolor=fg_color)
        self.style.configure('Horizontal.TScrollbar', background=button_bg, troughcolor=bg_color, arrowcolor=fg_color)
        self.style.map('TScrollbar', background=[('active', button_active_bg)])

        # Comboboxのスタイル
        self.style.map('TCombobox',
                       fieldbackground=[('readonly', entry_bg)],
                       selectbackground=[('readonly', select_bg)],
                       selectforeground=[('readonly', fg_color)],
                       foreground=[('readonly', fg_color)])

        # ステータスバーのスタイル
        self.status_label.configure(background=status_bar_bg, foreground=fg_color)

    def _on_map(self, event=None):
        """初回表示時に一度だけファイルパス列の幅を調整します。"""
        self.tree.unbind('<Map>') # 一度実行したら解除
        self._adjust_filepath_column()
        # カラムリサイズを検知するために、ヘッダーのドラッグ・リリースイベントにバインド
        self.tree.bind("<ButtonPress-1>", self._on_column_press, "+")
        self.tree.bind("<B1-Motion>", self._on_column_motion, "+")
        self.tree.bind("<ButtonRelease-1>", self._on_column_release, "+")

    def _adjust_filepath_column(self):
        """ファイルパス列の幅を、他の列の幅を引いた残りのスペースに合わせます。"""
        self.update_idletasks()
        tree_width = self.tree.winfo_width()
        other_columns_width = 0
        columns = list(self.tree['columns'])
        columns.remove('filepath') # filepath列を除外

        for col_id in columns:
            other_columns_width += self.tree.column(col_id, 'width')

        # # 垂直スクロールバーが表示されている場合、その幅を考慮
        # scrollbar_width = self.scrollbar_y.winfo_width() if self.scrollbar_y.winfo_ismapped() else 0
        # new_filepath_width = tree_width - other_columns_width - (scrollbar_width // 2) 
        new_filepath_width = tree_width - other_columns_width

        min_width = self.tree.column('filepath', 'minwidth')
        new_filepath_width = max(new_filepath_width, min_width)

        self.tree.column('filepath', width=new_filepath_width)
        self._update_horizontal_scrollbar() # 幅調整後にスクロールバーの状態を更新

    def _on_column_press(self, event):
        """カラムヘッダーがクリックされたときの処理。"""
        # ユーザーがカラムの境界線（separator）をドラッグしてリサイズを開始したかチェック
        if self.tree.identify_region(event.x, event.y) == "separator":
            self._is_resizing_column = True

    def _on_column_motion(self, event):
        """カラムリサイズ中にマウスが動いたときの処理。"""
        if self._is_resizing_column:
            self._update_horizontal_scrollbar()

    def _on_column_release(self, event):
        """マウスボタンが離されたときの処理。"""
        if self._is_resizing_column:
            self._is_resizing_column = False
            self._update_horizontal_scrollbar() # 最後に状態を更新

    def _update_horizontal_scrollbar(self, event=None):
        """Treeviewの幅と全カラムの合計幅を比較し、水平スクロールバーの表示/非表示を切り替えます。"""
        # update_idletasks() を呼び出して、ウィジェットのサイズが最新であることを保証
        self.update_idletasks()

        tree_width = self.tree.winfo_width()
        
        total_columns_width = sum(self.tree.column(c, 'width') for c in self.tree['columns'])

        # Treeviewの表示幅より列の合計幅が大きい場合のみスクロールバーを表示
        if total_columns_width > tree_width:
            self.scrollbar_x.grid(row=1, column=0, sticky="ew")
        else:
            self.scrollbar_x.grid_remove()

    # ファイルがドロップされた際のイベントハンドラ
    def handle_drop(self, event):
        """Treeviewへのファイルドラッグ＆ドロップを処理します。

        ドロップされたファイルパスを取得し、WAVファイルのみをリストに追加します。
        ファイルのサンプリング周波数とチャンネル数を取得して表示し、
        重複ファイルは無視します。
        自動変換モードが有効な場合は、変換タスクをキューに追加します。

        Args:
            event: TkinterDnDから渡されるドロップイベントオブジェクト。
                   event.dataにファイルパスのリストが含まれます。
        """
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
                # 拡張子が .wav または .wave のファイルのみを対象とする
                if not file_path.lower().endswith((".wav", ".wave")):
                    skipped_non_wav += 1
                    continue

                filename = os.path.basename(file_path)
                # 重複チェックのために絶対パスを使用
                filepath_abs = os.path.abspath(file_path) 

                is_duplicate = False # 重複フラグ
                for item_id_check in self.tree.get_children(): # Renamed item_id to avoid conflict
                    # valuesのインデックス1がファイルパス
                    if self.tree.item(item_id_check, "values")[1] == filepath_abs:
                        is_duplicate = True
                        skipped_duplicate += 1
                        break
                
                if is_duplicate:
                    continue

                try:
                    # soundfile.infoでサンプリング周波数とチャンネル数を取得
                    info = sf.info(filepath_abs)
                    original_sr = info.samplerate
                    original_channels = info.channels # チャンネル数を取得
                    original_subtype = info.subtype # ビット深度（サブタイプ）を取得

                    # Treeviewにアイテムを追加し、そのIDを取得
                    item_id = self.tree.insert("", tk.END, values=(filename, filepath_abs, original_sr, original_channels, original_subtype, ""))
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
                                target_channels = self._get_target_channels_from_gui() # 現在の目標チャンネル数を取得
                                target_subtype = self._get_target_subtype_from_gui() # 現在の目標ビット深度を取得
                                self.tree.set(item_id, column="status", value="キュー済")
                                # タスクキューに渡す情報にビット深度も追加
                                self.resample_task_queue.put((item_id, filepath_abs, target_sr_hz, target_channels, target_subtype, output_dir_for_task, filename, original_sr, original_channels, original_subtype))
                                self.status_var.set(f"キュー追加: {filename}")
                                self._ensure_worker_thread_running()
                            except ValueError as ve: # 目標SR値やチャンネル値が無効な場合
                                 self.tree.set(item_id, column="status", value="設定値エラー")
                                 self.status_var.set(f"変換設定値エラーのためキュー追加失敗: {ve}")
                        # else の場合、output_dir_for_task が None で、既にエラー処理されているはず

                except Exception as e:
                    self.status_var.set(f"エラー: {filename} の情報取得失敗 - {e}")
                    print(f"Error getting info for {filepath_abs}: {e}")

            # 処理結果をステータスバーに表示
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

        # update_status_and_button_states() を呼ぶと、ドロップ結果のメッセージが上書きされてしまうため、呼ばない。
        # ボタンの状態は、ユーザーがアイテムを選択した際に on_tree_select() によって更新されるため、ここでは不要。

    def _get_target_sr_from_gui(self):
        """GUIから目標サンプリング周波数を読み取ります。

        コンボボックスの選択値 (例: "44.1 kHz") をパースし、Hz単位の整数に変換して返します。
        無効な入力値の場合はValueErrorを発生させます。

        Returns:
            tuple[int, str]: (Hz単位の目標サンプリング周波数, GUIでの入力文字列)

        Raises:
            ValueError: パースに失敗した場合や、値が0以下の場合。
        """
        target_sr_input_str = self.target_sr_var.get()
        try:
            parts = target_sr_input_str.split()
            value = float(parts[0])
            unit = parts[1].lower()

            if unit == "khz":
                target_sr_hz = int(value * 1000)
            elif unit == "hz":
                target_sr_hz = int(value)
            else:
                raise ValueError(f"無効な単位です: {parts[1]}")

        except (ValueError, IndexError) as e:
            raise ValueError(f"目標サンプリング周波数の値 '{target_sr_input_str}' をパースできませんでした。") from e

        if target_sr_hz <= 0:
            raise ValueError("目標サンプリング周波数は正の整数である必要があります。")
        return target_sr_hz, target_sr_input_str

    def _get_target_channels_from_gui(self):
        """目標チャンネル数を取得します。現在はステレオ(2)に固定されています。

        Returns:
            int: 目標チャンネル数 (2: ステレオ)。
        """
        return 2

    def _get_target_subtype_from_gui(self):
        """GUIから目標ビット深度に対応するsubtype文字列を取得します。

        Returns:
            str: soundfileで利用可能なサブタイプ文字列 (例: "PCM_16")

        Raises:
            ValueError: 無効な選択肢の場合。
        """
        selected_str = self.target_bit_depth_var.get()
        if "16bit" in selected_str:
            return "PCM_16"
        elif "8bit" in selected_str:
            return "PCM_S8"
        else:
            raise ValueError("無効なビット深度が選択されています。")

    def clear_list(self):
        """ファイルリスト（Treeview）の内容をすべてクリアします。"""
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.status_var.set("ファイルリストがクリアされました。")
        self.on_tree_select() # クリア後は何も選択されていないのでボタン状態更新

    # 「一括変換実行」ボタンが押されたときの処理
    def start_resampling_process(self):
        """リスト内のすべてのファイルの変換処理を開始します（一括変換）。

        GUIから設定値を取得し、リスト内の各ファイルに対して変換処理を実行します。
        処理中はUIを無効化し、完了後に結果をメッセージボックスで表示します。
        """
        items = self.tree.get_children()
        if not items:
            # リストが空の場合は警告を表示
            messagebox.showwarning("情報なし", "変換対象のファイルがリストにありません。")
            self.status_var.set("リストにファイルがありません。")
            return

        try:
            target_sr, _ = self._get_target_sr_from_gui()
            target_channels = self._get_target_channels_from_gui()
            target_subtype = self._get_target_subtype_from_gui()
        except ValueError as e:
            messagebox.showerror("入力エラー", str(e))
            self.status_var.set(str(e))
            return

        # 「ソース元に保存」がOFFの場合、保存先フォルダを選択させる
        output_dir_for_batch = None
        if not self.save_to_source_var.get():
            output_dir_for_batch = filedialog.askdirectory(title="変換後のファイルの保存先フォルダを選択してください")
            if not output_dir_for_batch:
                self.status_var.set("保存先フォルダが選択されませんでした。処理を中止します。")
                return

        # 処理中はUIを無効化
        self.resample_button.config(state=tk.DISABLED)
        self.clear_button.config(state=tk.DISABLED)
        self.delete_button.config(state=tk.DISABLED) 
        self.individual_resample_button.config(state=tk.DISABLED) 
        self.auto_resample_check.config(state=tk.DISABLED) 
        self.save_to_source_check.config(state=tk.DISABLED) 

        # カウンタの初期化
        processed_count = 0
        error_count = 0
        skipped_count = 0
        actually_converted_count = 0 

        for item_id in items:
            values = self.tree.item(item_id, "values")
            # GUIからはファイル名とパスのみ取得
            filename, filepath, _, _, _, _ = values

            # GUIの値を信頼せず、処理直前にファイルから直接メタデータを再取得
            try:
                info = sf.info(filepath)
                original_sr = info.samplerate
                original_channels = info.channels
                original_subtype = info.subtype
            except Exception as e:
                self.tree.set(item_id, column="status", value="エラー")
                self.status_var.set(f"{filename}: メタデータ読込エラー - {e}")
                error_count += 1
                continue # 次のファイルへ
            # GUIのステータスを更新
            self.tree.set(item_id, column="status", value="処理中...")
            self.status_var.set(f"処理中: {filename}...")
            self.update_idletasks() 

            # 出力先ディレクトリを決定
            current_output_dir = ""
            if self.save_to_source_var.get():
                current_output_dir = os.path.dirname(filepath)
            else:
                current_output_dir = output_dir_for_batch 

            # 変換ロジックにビット深度の情報も渡す
            result_status, message = self._perform_single_resample_logic(filepath, original_sr, original_channels, original_subtype, target_sr, target_channels, target_subtype, current_output_dir, filename)
            
            # 結果をGUIに反映
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

        # 処理完了後、UIを再度有効化
        self.resample_button.config(state=tk.NORMAL)
        self.clear_button.config(state=tk.NORMAL)
        self.auto_resample_check.config(state=tk.NORMAL) 
        self.save_to_source_check.config(state=tk.NORMAL) 
        self.on_tree_select() 

        # 最終結果をメッセージボックスとステータスバーで表示
        final_message_parts = [f"{actually_converted_count}個のファイルを変換しました。"]
        if skipped_count > 0:
             final_message_parts.append(f"{skipped_count}個スキップ。")
        if error_count > 0:
            final_message = f"処理完了。{actually_converted_count}個成功、{error_count}個エラー、{skipped_count}個スキップ。"
            messagebox.showwarning("処理完了（一部エラーあり）", final_message)
        else:
            final_message = f"処理完了。{actually_converted_count}個のファイルが正常に変換されました。"
            if skipped_count > 0:
                final_message += f" ({skipped_count}個は目標周波数とチャンネル数と同一のためスキップ)"
            messagebox.showinfo("処理完了", final_message)
        
        self.status_var.set(final_message)

    # 「選択ファイル変換」ボタンが押されたときの処理
    def start_selected_resampling_process(self):
        """リストで選択されているファイルの変換処理を開始します。

        GUIから設定値を取得し、選択された各ファイルに対して変換処理を実行します。
        処理中はUIを無効化し、完了後に結果をメッセージボックスで表示します。
        """
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("情報なし", "変換対象のファイルが選択されていません。")
            self.status_var.set("選択されたファイルがありません。")
            return

        try:
            target_sr, _ = self._get_target_sr_from_gui()
            target_channels = self._get_target_channels_from_gui()
            target_subtype = self._get_target_subtype_from_gui()
        except ValueError as e:
            messagebox.showerror("入力エラー", str(e))
            self.status_var.set(str(e))
            return

        # 「ソース元に保存」がOFFの場合、保存先フォルダを選択させる（前回選択したフォルダを記憶）
        output_dir_for_selected = None
        if not self.save_to_source_var.get():
            initial_dir = self.last_individual_output_dir if self.last_individual_output_dir else None
            output_dir_for_selected = filedialog.askdirectory(
                title="選択ファイルの保存先フォルダを選択してください",
                initialdir=initial_dir
            )
            if not output_dir_for_selected:
                self.status_var.set("保存先フォルダが選択されませんでした。処理を中止します。")
                return
            self.last_individual_output_dir = output_dir_for_selected 

        # 処理中はUIを無効化
        self.resample_button.config(state=tk.DISABLED)
        self.individual_resample_button.config(state=tk.DISABLED)
        self.clear_button.config(state=tk.DISABLED)
        self.delete_button.config(state=tk.DISABLED)
        self.auto_resample_check.config(state=tk.DISABLED)
        self.save_to_source_check.config(state=tk.DISABLED)

        # カウンタの初期化
        processed_count = 0
        error_count = 0
        skipped_count = 0
        actually_converted_count = 0

        for item_id in selected_items:
            values = self.tree.item(item_id, "values")
            # GUIからはファイル名とパスのみ取得
            filename, filepath, _, _, _, _ = values

            # GUIの値を信頼せず、処理直前にファイルから直接メタデータを再取得
            try:
                info = sf.info(filepath)
                original_sr = info.samplerate
                original_channels = info.channels
                original_subtype = info.subtype
            except Exception as e:
                self.tree.set(item_id, column="status", value="エラー")
                self.status_var.set(f"{filename}: メタデータ読込エラー - {e}")
                error_count += 1
                continue # 次のファイルへ
            # GUIのステータスを更新
            self.tree.set(item_id, column="status", value="処理中...")
            self.status_var.set(f"処理中: {filename}...")
            self.update_idletasks()

            # 出力先ディレクトリを決定
            current_output_dir = os.path.dirname(filepath) if self.save_to_source_var.get() else output_dir_for_selected

            # 変換処理を実行
            result_status, message = self._perform_single_resample_logic(filepath, original_sr, original_channels, original_subtype, target_sr, target_channels, target_subtype, current_output_dir, filename)
            # 結果をGUIに反映
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

        # 処理完了後、UIを再度有効化
        self.auto_resample_check.config(state=tk.NORMAL)
        self.save_to_source_check.config(state=tk.NORMAL)
        self.clear_button.config(state=tk.NORMAL) 
        self.update_status_and_button_states() 
        self.on_tree_select() 

        # 最終結果をメッセージボックスとステータスバーで表示
        final_message = f"選択ファイル処理完了。{actually_converted_count}個成功、{error_count}個エラー、{skipped_count}個スキップ。"
        messagebox.showinfo("処理完了", final_message)
        self.status_var.set(final_message)

    # 「選択消去」ボタンが押されたときの処理
    def delete_selected_items(self):
        """ファイルリスト（Treeview）で選択されているアイテムを削除します。"""
        selected_items = self.tree.selection()
        if not selected_items:
            self.status_var.set("消去するアイテムが選択されていません。")
            return

        for item_id in selected_items:
            self.tree.delete(item_id)
        
        self.status_var.set(f"{len(selected_items)} 個のアイテムをリストから消去しました。")
        self.on_tree_select() # 削除後、選択状態が変わるのでボタン状態更新

    # Treeviewのアイテム選択が変更されたときのイベントハンドラ
    def on_tree_select(self, event=None):
        """ファイルリストのアイテム選択状態の変更をハンドルします。

        アイテムが選択されているかどうかに応じて、「選択消去」ボタンと
        「選択ファイル変換」ボタンの有効/無効状態を切り替えます。

        Args:
            event: Tkinterから渡されるイベントオブジェクト（通常は使用しない）。
        """
        has_selection = bool(self.tree.selection())
        auto_mode = self.auto_resample_var.get()

        if has_selection:
            self.delete_button.config(state=tk.NORMAL)
            self.individual_resample_button.config(state=tk.NORMAL if not auto_mode else tk.DISABLED)
        else:
            self.delete_button.config(state=tk.DISABLED)
            self.individual_resample_button.config(state=tk.DISABLED)

    # 「自動で変更する」チェックボックスの状態が変更されたときの処理
    def on_auto_resample_toggle(self):
        """「自動で変更する」チェックボックスの状態変更を処理します。

        自動変換モードがONになり、かつ出力先が指定されていない場合、
        ユーザーにフォルダ選択ダイアログを表示します。
        """
        auto_mode_now = self.auto_resample_var.get()
        save_to_source = self.save_to_source_var.get()

        # 自動モードONかつソース元保存OFFの場合、出力先が設定されていなければ尋ねる
        if auto_mode_now and not save_to_source:
            if not self.auto_output_dir:
                messagebox.showinfo("出力先指定", "自動変換用の出力先フォルダを指定してください。\n「ソース元に保存」がOFFのため、出力先が必要です。")
                new_dir = filedialog.askdirectory(title="自動変換ファイルの保存先フォルダを選択")
                if new_dir:
                    self.auto_output_dir = new_dir
                else:
                    self.auto_resample_var.set(False) 
                    messagebox.showwarning("出力先未指定", "出力先が指定されなかったため、自動変換をOFFにしました。")
        self.update_status_and_button_states()

    # 「ソース元に保存」チェックボックスの状態が変更されたときの処理
    def on_save_to_source_toggle(self):
        """「ソース元に保存」チェックボックスの状態変更を処理します。

        自動変換モードがONの状態でこのチェックボックスがOFFにされた場合、
        出力先が指定されていなければユーザーにフォルダ選択ダイアログを表示します。
        """
        auto_mode = self.auto_resample_var.get()
        save_to_source_now = self.save_to_source_var.get()

        # 自動モードONかつソース元保存OFFになった場合、出力先が設定されていなければ尋ねる
        if auto_mode and not save_to_source_now:
             if not self.auto_output_dir:
                messagebox.showinfo("出力先指定", "「ソース元に保存」がOFFのため、自動変換用の出力先フォルダを指定してください。")
                new_dir = filedialog.askdirectory(title="自動変換ファイルの保存先フォルダを選択")
                if new_dir:
                    self.auto_output_dir = new_dir
                else:
                    self.auto_resample_var.set(False) 
                    messagebox.showwarning("出力先未指定", "出力先が指定されなかったため、自動変換をOFFにしました。")
        self.update_status_and_button_states()

    # 現在のモード設定に基づいてUI（ボタン状態、ステータスメッセージ）を更新する
    def update_status_and_button_states(self):
        """現在の設定モードに応じて、UIのボタン状態とステータスメッセージを更新します。

        自動変換モードか手動変換モードか、また保存先の指定方法に応じて、
        各ボタンの有効/無効を切り替え、ステータスバーに現在の状態を
        分かりやすく表示します。
        """
        auto_mode = self.auto_resample_var.get()
        save_to_source = self.save_to_source_var.get()

        # 手動での一括変換処理中かどうかを判定
        is_processing_manually = self.resample_button['state'] == tk.DISABLED and not auto_mode
        
        # 手動変換中でなければ、モードに応じてボタンの状態を更新
        if not is_processing_manually: 
            if auto_mode:
                self.resample_button.config(state=tk.DISABLED)
                self.individual_resample_button.config(state=tk.DISABLED) 
            else:
                self.resample_button.config(state=tk.NORMAL)
                self.on_tree_select() 

        # モードに応じてステータスバーのメッセージを更新
        if auto_mode:
            self.resample_button.config(state=tk.DISABLED)
            if save_to_source:
                self.status_var.set("自動変換 ON (ソース元へ保存)。ファイルドロップで自動処理。")
            else: 
                if not self.auto_output_dir:
                    self.status_var.set("自動変換 ON (出力先未指定)。設定を確認してください。")
                else:
                     self.status_var.set(f"自動変換 ON (出力先: {os.path.basename(self.auto_output_dir) if self.auto_output_dir else '未指定'})。ファイルドロップで自動処理。")
            self._ensure_worker_thread_running()
        else: # 手動モードの場合
            if not is_processing_manually:
                self.resample_button.config(state=tk.NORMAL) 
                if save_to_source:
                    self.status_var.set("手動変換 (ソース元へ保存)。「一括変換実行」ボタンで処理。")
                else: 
                    self.status_var.set("手動変換 (指定フォルダへ保存)。「一括変換実行」ボタンで処理。")

    # 自動変換用のワーカースレッドを起動・確認する
    def _ensure_worker_thread_running(self):
        """自動変換用のワーカースレッドが実行中でなければ起動します。

        スレッドが未作成か、すでに終了している場合に新しいスレッドを
        作成して開始します。これにより、自動変換タスクを非同期で
        処理できるようになります。
        """
        if self.worker_thread is None or not self.worker_thread.is_alive():
            self.worker_thread = threading.Thread(target=self._worker_resample_files, daemon=True)
            self.worker_thread.start()
            print("ワーカースレッドを開始しました。")

    # ワーカースレッドで実行されるファイル変換処理のメインループ
    def _worker_resample_files(self):
        """ワーカースレッドのメインループです。

        タスクキューを監視し、追加された変換タスクを一つずつ取り出して
        `_perform_single_resample_logic` メソッドで処理します。
        処理結果は結果キューに格納され、メインスレッド（GUI）に通知されます。
        このループはアプリケーション終了フラグが立つまで継続します。
        """
        print("ワーカースレッド実行中...")
        while not self.is_shutting_down:
            item_id = None 
            try:
                # タスクキューからアイテムを取得 (item_id, filepath, target_sr, target_channels, output_dir, filename, original_sr, original_channels)
                item_id, filepath, target_sr, target_channels, target_subtype, output_dir, filename, original_sr, original_channels, original_subtype = self.resample_task_queue.get(timeout=1)
                
                # GUIに「処理中」であることを通知
                self.resample_results_queue.put((item_id, "処理中...", None)) 

                result_status, message = self._perform_single_resample_logic(filepath, original_sr, original_channels, original_subtype, target_sr, target_channels, target_subtype, output_dir, filename)
                # 処理結果を結果キューに入れる
                self.resample_results_queue.put((item_id, result_status, message))
                self.resample_task_queue.task_done()
            except queue.Empty:
                continue 
            except Exception as e:
                print(f"ワーカースレッドで予期せぬエラー: {e}")
                if item_id: 
                   self.resample_results_queue.put((item_id, "エラー", str(e)))
        print("ワーカースレッドを終了します。")

    # 実際のファイル変換ロジック
    def _perform_single_resample_logic(self, filepath, original_sr, original_channels, original_subtype, target_sr, target_channels, target_subtype, output_dir, filename):
        """単一ファイルのサンプリング周波数・チャンネル変換・ビット深度固定のロジックを実行します。

        librosaを使用してオーディオファイルを読み込み、リサンプリングと
        チャンネル変換を行い、soundfileを使用して指定されたビット深度（16bit）で
        新しいファイルとして書き出します。
        変換が不要な場合はスキップします。

        Args:
            filepath (str): 処理対象のファイルパス。
            original_sr (int): 元のサンプリング周波数。
            original_channels (int): 元のチャンネル数。
            original_subtype (str): 元のビット深度(サブタイプ)。
            target_sr (int): 目標のサンプリング周波数。
            target_channels (int): 目標のチャンネル数。
            target_subtype (str): 目標のビット深度(サブタイプ)。
            output_dir (str): 出力先ディレクトリ。
            filename (str): 元のファイル名。

        Returns:
            tuple[str, str]: (処理結果のステータス文字列, 詳細メッセージ)
        """
        try:
            # 1. スキップ判定: 全てのパラメータが目標と一致する場合、ファイル操作を行わずに処理を終了
            if original_sr == target_sr and original_channels == target_channels and original_subtype == target_subtype:
                msg = f"スキップ: {filename} (既に目標設定と同一です)"
                return "処理済", msg

            # 2. 変換処理: スキップされなかった場合は、何らかの変換が必要
            # librosa.loadでステレオを保持するためにはmono=Falseを明示的に指定
            # yは(channels, samples)または(samples,)のndarrayになる
            y, sr_librosa_original = librosa.load(filepath, sr=None, mono=False)

            y_processed = y
            # サンプリング周波数変換
            if sr_librosa_original != target_sr:
                y_processed = librosa.resample(y=y_processed, orig_sr=sr_librosa_original, target_sr=target_sr)

            # チャンネル数変換
            # y_processedの次元数をチェック (モノラルの場合は1次元、ステレオの場合は2次元)
            # librosa.load(mono=False) はステレオの場合 (2, samples) の形状になる
            # soundfile.write は (samples, channels) の形状を期待するため、転置が必要
            if y_processed.ndim == 1 and target_channels == 2: # モノラルからステレオへ（複製）
                y_processed = np.vstack([y_processed, y_processed]) # モノラルを複製してステレオにする
            elif y_processed.ndim == 2 and target_channels == 1: # ステレオからモノラルへ（今回は発生しないはずだが念のため）
                # ステレオからモノラルへのダウンミックスは librosa に任せるか、平均を取るなど
                # 今回はターゲットがステレオ固定なので、このパスは基本的には通らない
                # もし将来的にモノラル変換が必要になった場合のプレースホルダー
                pass
            # チャンネル数が既にtarget_channelsと一致している場合は何もしない

            # soundfile.writeは(frames, channels)形式を期待するため、librosaが返す(channels, frames)を転置
            if y_processed.ndim == 2: # ステレオの場合
                y_processed = y_processed.T # 転置して(samples, channels)にする

            # 3. ファイル書き出し
            base, ext = os.path.splitext(filename)
            bit_depth_str = "16bit" if target_subtype == "PCM_16" else "8bit"
            output_filename = f"{base}_resampled_{target_sr}Hz_{target_channels}ch_{bit_depth_str}{ext}" # ファイル名にビット深度も追加
            output_path = os.path.join(output_dir, output_filename)

            # 出力先ディレクトリが存在しない場合は作成
            if not os.path.exists(output_dir):
                try:
                    os.makedirs(output_dir)
                    print(f"作成された出力ディレクトリ: {output_dir}")
                except OSError as ose:
                    error_msg = f"エラー: 出力ディレクトリの作成に失敗しました ({output_dir}) - {ose}"
                    print(error_msg)
                    return "エラー", error_msg

            sf.write(output_path, y_processed, target_sr, subtype=target_subtype)
            success_msg = f"変換成功: {output_filename}"
            return "処理済", success_msg
        except Exception as e:
            error_msg = f"エラー: {filename} の変換に失敗 - {e}"
            print(error_msg)
            return "エラー", str(e)

    # ワーカースレッドからの結果をGUIに反映させるためのポーリング処理
    def process_resample_results(self):
        """ワーカースレッドからの処理結果をGUIに反映させます。

        結果キューを定期的にポーリングし、キューに結果があれば取り出して
        ファイルリストの「状態」列とステータスバーを更新します。
        このメソッドは `after` メソッドによって定期的に呼び出されます。
        """
        try:
            while not self.resample_results_queue.empty():
                item_id, status, message = self.resample_results_queue.get_nowait()
                # Treeviewからアイテムが削除されている可能性を考慮
                if self.tree.exists(item_id):
                    self.tree.set(item_id, column="status", value=status)
                    filename_in_tree = self.tree.item(item_id, "values")[0]
                    if status == "処理中...":
                        self.status_var.set(f"{filename_in_tree}: 処理中...")
                    else:
                        self.status_var.set(f"{filename_in_tree}: {status} {(' - ' + message) if message else ''}")
                self.resample_results_queue.task_done()
        finally:
            if not self.is_shutting_down:
                # 100ms後に再度このメソッドを呼び出す
                self._process_timer_id = self.after(100, self.process_resample_results)

    # ウィンドウが閉じられるときの処理
    def on_closing(self):
        """ウィンドウが閉じられる際のクリーンアップ処理を実行します。

        ユーザーに終了確認のダイアログを表示し、OKが押されたら
        ワーカースレッドの終了を待ってからアプリケーションを安全に終了します。
        """
        if messagebox.askokcancel("終了確認", "アプリケーションを終了しますか？"):
            self.is_shutting_down = True

            # afterループを止める
            if self._process_timer_id:
                self.after_cancel(self._process_timer_id)

            print("シャットダウン処理を開始します...")
            # ワーカースレッドが動いていれば、終了を待つ
            if self.worker_thread and self.worker_thread.is_alive():
                print("ワーカースレッドの終了を待機中...")
                self.worker_thread.join(timeout=2.0) # 最大2秒待つ
                if self.worker_thread.is_alive():
                    print("ワーカースレッドがタイムアウト後も実行中です。")
                else:
                    print("ワーカースレッドは正常に終了しました。")
            self.destroy()

if __name__ == "__main__":
    # アプリケーションのエントリーポイント
    app = AudioResamplerApp()
    app.mainloop()