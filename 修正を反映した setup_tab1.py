    def setup_tab1(self, bg_color, label_style):
        parent = self.tab1
        parent.configure(bg=bg_color)
        pad = 10

        # ...（①-1 の部分は変更なし）...

        frame2 = ttk.LabelFrame(parent, text="①-2 入力規則を自動セット", style=label_style, padding=pad)
        frame2.pack(fill="x", padx=pad, pady=pad)

        row2 = tk.Frame(frame2, bg=bg_color)
        row2.pack(fill="x", pady=2)
        ttk.Label(row2, text="マスターExcel：").pack(side="left")
        self.val_master = tk.StringVar()
        ttk.Entry(row2, textvariable=self.val_master, width=30, font=self.base_font).pack(side="left", padx=5)
        ttk.Button(row2, text="参照", command=lambda: self.browse_file(self.val_master, "Excel files", ".xlsx")).pack(side="left")

        row3 = tk.Frame(frame2, bg=bg_color)
        row3.pack(fill="x", pady=2)
        ttk.Label(row3, text="マスターシート名：").pack(side="left")
        self.val_master_sheet = tk.StringVar(value="マスタリスト")
        ttk.Entry(row3, textvariable=self.val_master_sheet, width=15, font=self.base_font).pack(side="left", padx=5)
        ttk.Label(row3, text="マスター列：").pack(side="left", padx=10)
        self.val_master_col = tk.StringVar(value="A")
        ttk.Entry(row3, textvariable=self.val_master_col, width=5, font=self.base_font).pack(side="left")

        row4 = tk.Frame(frame2, bg=bg_color)
        row4.pack(fill="x", pady=2)
        ttk.Label(row4, text="対象Excel：").pack(side="left")
        self.val_target = tk.StringVar()
        ttk.Entry(row4, textvariable=self.val_target, width=30, font=self.base_font).pack(side="left", padx=5)
        ttk.Button(row4, text="参照", command=lambda: self.browse_file(self.val_target, "Excel files", ".xlsx")).pack(side="left")

        row5 = tk.Frame(frame2, bg=bg_color)
        row5.pack(fill="x", pady=2)
        ttk.Label(row5, text="対象シート名：").pack(side="left")
        self.val_target_sheet = tk.StringVar(value="校印使用台帳")
        ttk.Entry(row5, textvariable=self.val_target_sheet, width=15, font=self.base_font).pack(side="left", padx=5)
        ttk.Label(row5, text="設定列：").pack(side="left", padx=10)
        self.val_target_col = tk.StringVar(value="E")
        ttk.Entry(row5, textvariable=self.val_target_col, width=5, font=self.base_font).pack(side="left")
        ttk.Label(row5, text="開始行：").pack(side="left", padx=10)
        self.val_start_row = tk.IntVar(value=3)
        ttk.Entry(row5, textvariable=self.val_start_row, width=5, font=self.base_font).pack(side="left")
        ttk.Label(row5, text="終了行（0=最終行まで）：").pack(side="left", padx=10)
        self.val_end_row = tk.IntVar(value=0)
        ttk.Entry(row5, textvariable=self.val_end_row, width=5, font=self.base_font).pack(side="left")
        ttk.Button(row5, text="▶ 入力規則セット", command=self.run_validation_set, style="Green.TButton").pack(side="left", padx=10)