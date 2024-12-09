import tkinter as tk
from tkinter import ttk
from .logic import add_file, clear_list, save_to_excel, delete_file


def create_gui():
    root = tk.Tk()
    root.geometry("700x550")
    root.title("Y-maze processing")

    style = ttk.Style()
    style.theme_use("winnative")

    button_frame = tk.Frame(root)
    button_frame.pack(pady=(7, 0))

    adding_button = tk.Button(button_frame, text="Добавить файлы", width=20)
    adding_button.grid(row=0, column=0, padx=(0, 4))

    clear_button = tk.Button(
        button_frame, text="Очистить список", width=20, state="disabled"
    )
    clear_button.grid(row=0, column=1, padx=(4, 0))

    tooltip_text = "'del' - удалить выбранный файл"

    tooltip = tk.Label(root, text=tooltip_text, font=("Arial", 8), fg="gray")

    tooltip.place(in_=button_frame, relx=1.91, rely=0.1, x=-104, y=0, anchor="ne")

    listbox = tk.Listbox(root, selectmode="browse", width=51, height=10)
    listbox.pack(pady=(5, 0), padx=(195, 195), fill="both", expand=True)

    result_text = ttk.Treeview(
        root, columns=("column1", "column2", "column3", "column4", "column5")
    )

    result_text.configure(height=12)
    result_text.heading("#0", text="\nГруппа, №\n")
    result_text.column("#0", width=143)
    result_text.heading("column1", text="Триплеты")
    result_text.column("column1", width=85, anchor="e")
    result_text.heading("column2", text="     Возвраты в\nпервоначальное")
    result_text.column("column2", width=115, anchor="e")
    result_text.heading("column3", text="Обратные\n возвраты")
    result_text.column("column3", width=95, anchor="e")
    result_text.heading("column4", text="Двигательная\n  активность")
    result_text.column("column4", width=105, anchor="e")
    result_text.heading("column5", text="Эффективность\nпатрулирования")
    result_text.column("column5", width=135, anchor="e")
    result_text.pack(pady=(6, 0), padx=10, fill="both", expand=True)

    save_button = tk.Button(root, text="Сохранить в Excel", width=15, state="disabled")
    save_button.pack(anchor="se", padx=10, pady=10)

    root.grid_columnconfigure(0, weight=1)
    root.grid_columnconfigure(1, weight=1)
    root.grid_rowconfigure(2, weight=1)

    adding_button.config(
        command=lambda: add_file(listbox, result_text, save_button, clear_button)
    )
    clear_button.config(
        command=lambda: clear_list(listbox, result_text, save_button, clear_button)
    )
    save_button.config(command=lambda: save_to_excel())
    listbox.bind(
        "<Delete>",
        lambda event: delete_file(listbox, result_text, save_button, clear_button),
    )

    root.mainloop()
