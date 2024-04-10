import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import os
import subprocess
import sys

files, file_names, result_df = [], [], pd.DataFrame()


def add_file():
    global result_df
    files_list = filedialog.askopenfilenames(
        filetypes=(("Text files", "*.txt"), ("All files", "*.*")),
        title="Выберите файл(ы)")

    for file in files_list:
        data_string = read_data_string_from_file(file)
        if data_string is not None:
            files.append(file)
            file_names.append(os.path.basename(file))
            save_button.config(state='normal')
            clear_button.config(state='normal')
        else:
            messagebox.showerror(
                "Ошибка при чтении файла", f"Убедитесь, что файл {os.path.basename(file)} принадлежит Real Timer.")

    listbox.delete(0, 'end')
    for name in file_names:
        listbox.insert('end', name)
    process_file()


def process_file():
    global result_df
    result_df = pd.DataFrame()
    for file in files:
        data_string = read_data_string_from_file(file)
        write_file(data_string)
        df = read_dataframe()

        temp_result = process_dataframe(df)

        dic = calculate_metrics(temp_result)

        temp_df = create_temp_dataframe(dic)

        result_df = pd.concat([result_df, temp_df], ignore_index=False)
        update_result_text()


def read_data_string_from_file(file):
    with open(file, "r", encoding="cp866") as f:
        lines = f.readlines()

    result = ""
    header_found = False

    for line in lines:
        if line.strip() == "key\tevent\ttime\tdur\ttmofday":
            header_found = True
        if header_found:
            result += line

    if not header_found:
        return None

    return result


def write_file(data_string):
    with open(os.path.join(os.getcwd(), 'data.csv'), 'w', encoding="cp866") as f:
        f.write(data_string)


def read_dataframe():
    df = pd.read_csv(
        os.path.join(
            os.getcwd(),
            'data.csv'),
        delimiter='\t',
        encoding="cp866")
    os.remove(os.path.join(os.getcwd(), 'data.csv'))
    return df


def process_dataframe(df):
    processed_data = []
    sequence = ""
    for i, row in df.iterrows():
        if row['key'] == 'Reset':
            sequence = str(row['tmofday'] + ': ')
        elif row['key'] in {'Num1', 'Num2', 'Num3'}:
            sequence += row['key'].replace('Num', '')
        elif row['key'] == 'Exit':
            processed_data.append(sequence)
    return processed_data


def calculate_metrics(temp_result):
    dic = {}
    for el in temp_result:
        t = el.split(': ')[1][::2]
        triplets = sum(t[i:i + 3] in {'123',
                                      '132',
                                      '213',
                                      '231',
                                      '312',
                                      '321'} for i in range(len(t) - 2))
        returns = sum(t[i:i + 2] in {'11', '22', '33'}
                      for i in range(len(t) - 1))
        rollback = sum(t[i:i + 3] in {'121',
                                      '131',
                                      '212',
                                      '232',
                                      '313',
                                      '323'} for i in range(len(t) - 2))
        activity = len(t)
        if activity <= 2:
            efficiency = 0.0
        else:
            efficiency = round((triplets / (activity - 2)) * 100, 1)

        key = el.split(': ')[0]
        dic[key] = [triplets, rollback, returns, activity, efficiency]
    return dic


def create_temp_dataframe(dic):
    temp_df = pd.DataFrame.from_dict(dic, orient='index',
                                     columns=["Триплеты", "Возвраты", "Обратные возвраты",
                                              "Двигательная активность", "Эффективность патрулирования"])
    return temp_df


def open_directory(save_path):
    os.path.dirname(save_path)

    if sys.platform.startswith('darwin'):
        subprocess.run(['open', '-R', save_path])
    elif sys.platform.startswith('win32'):
        subprocess.run(['explorer', '/select,', save_path.replace('/', '\\')])
    elif sys.platform.startswith('linux'):
        subprocess.run(['nautilus', '--select', save_path])


def save_to_excel():
    try:
        file_path = filedialog.asksaveasfile(
            mode='wb',
            defaultextension='.xlsx',
            filetypes=[('Excel files', '*.xlsx')],
            title='Выберите путь сохранения',
            initialfile='Лист Microsoft Excel.xlsx'
        )

        if file_path:
            save_path = file_path.name

            result_df.to_excel(save_path, index=True)
            root.update()

            open_folder = messagebox.askyesno(
                "Открыть папку?", "Хотите открыть папку с сохраненным файлом?")

            if open_folder:
                open_directory(save_path)

    except Exception as e:
        error_message = f"Ошибка при записи в файл: {str(e)}"
        messagebox.showerror("Ошибка", error_message)

    if result_df.empty:
        save_button.config(state='disabled')
    else:
        save_button.config(state='normal')


def update_result_text():
    result_text.delete(*result_text.get_children())
    for index, row in result_df.iterrows():
        result_text.insert('', 'end', text=str(row.name), values=(
            int(row['Триплеты']), int(row['Возвраты']), int(
                row['Обратные возвраты']),
            int(row['Двигательная активность']), round(row['Эффективность патрулирования'], 2)))

    if result_df.empty:
        save_button.config(state='disabled')
    else:
        save_button.config(state='normal')


def clear_list():
    global result_df, files, file_names
    result_df = pd.DataFrame()
    files = []
    file_names = []
    listbox.delete(0, 'end')
    update_result_text()
    result_text.delete(*result_text.get_children())
    save_button.config(state='disabled')
    clear_button.config(state='disabled')


def delete_file():
    selected_item = listbox.curselection()
    if selected_item:
        index = selected_item[0]
        del files[index]
        del file_names[index]
        listbox.delete(selected_item)
        update_result_text()
        result_text.delete(*result_text.get_children())
        process_file()
    if listbox.size() == 0:
        save_button.config(state='disabled')
        clear_button.config(state='disabled')


root = tk.Tk()
root.geometry('700x550')
root.title('Y-maze processing')

button_frame = tk.Frame(root)
button_frame.pack(pady=(7, 0))

adding_button = tk.Button(
    button_frame,
    text="Добавить файлы",
    command=add_file,
    width=20)
adding_button.grid(row=0, column=0, padx=(0, 4))

clear_button = tk.Button(
    button_frame,
    text="Очистить список",
    command=clear_list,
    width=20,
    state='disabled')
clear_button.grid(row=0, column=1, padx=(4, 0))

tooltip_text = """
'del' - удалить выбранный файл
"""

tooltip = tk.Label(root, text=tooltip_text, font=("Arial", 8), fg="gray")

tooltip.place(
    in_=button_frame,
    relx=1.91,
    rely=-0.45,
    x=-104,
    y=0,
    anchor="ne")

listbox = tk.Listbox(root, selectmode='browse', width=51, height=10)
listbox.pack(pady=(5, 0), padx=(195, 195), fill='both', expand=True)

style = ttk.Style()
style.theme_use('winnative')

result_text = ttk.Treeview(
    root,
    columns=(
        'column1',
        'column2',
        'column3',
        'column4',
        'column5'))
result_text.configure(height=12)
result_text.heading('#0', text='\nГруппа, №\n')
result_text.column('#0', width=143)
result_text.heading('column1', text='Триплеты')
result_text.column('column1', width=85, anchor='e')
result_text.heading('column2', text='     Возвраты в\nпервоначальное')
result_text.column('column2', width=115, anchor='e')
result_text.heading('column3', text='Обратные\n возвраты')
result_text.column('column3', width=95, anchor='e')
result_text.heading('column4', text='Двигательная\n  активность')
result_text.column('column4', width=105, anchor='e')
result_text.heading('column5', text='Эффективность\nпатрулирования')
result_text.column('column5', width=135, anchor='e')
result_text.pack(pady=(6, 0), padx=10, fill='both', expand=True)


save_button = tk.Button(
    root,
    text="Сохранить в Excel",
    command=save_to_excel,
    width=15,
    state='disabled')
save_button.pack(anchor="se", padx=10, pady=10)

root.grid_columnconfigure(0, weight=1)
root.grid_columnconfigure(1, weight=1)
root.grid_rowconfigure(2, weight=1)

listbox.bind("<Delete>", lambda event: delete_file())

root.mainloop()
