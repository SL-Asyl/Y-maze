from tkinter import filedialog, messagebox
import openpyxl
import pandas as pd
import os
import subprocess
import sys

files, file_names, result_df = [], [], pd.DataFrame()


def add_file(listbox, result_text, save_button, clear_button):
    global result_df
    files_list = filedialog.askopenfilenames(
        filetypes=(("Text files", "*.txt"), ("All files", "*.*")),
        title="Выберите файл(ы)",
    )

    for file in files_list:
        data_string = read_data_string_from_file(file)
        if data_string:
            files.append(file)
            file_names.append(os.path.basename(file))
            save_button.config(state="normal")
            clear_button.config(state="normal")
        else:
            messagebox.showerror(
                "Ошибка при чтении файла",
                f"Убедитесь, что файл {os.path.basename(file)} принадлежит Real Timer.",
            )

    listbox.delete(0, "end")
    for name in file_names:
        listbox.insert("end", name)
    process_file(result_text, save_button)


def process_file(result_text, save_button):
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
        update_result_text(result_text, save_button)


def read_data_string_from_file(file):
    with open(file, "r", encoding="cp866") as f:
        lines = f.readlines()

    result = []
    header_found = False

    for line in lines:
        if line.strip() == "key\tevent\ttime\tdur\ttmofday":
            header_found = True
        if header_found:
            if line.startswith("Reset"):
                parts = line.strip().split("\t")
                if len(parts) < 5 or not parts[4].strip():
                    while len(parts) < 5:
                        parts.append("")
                    parts[4] = "N/A"
                line = "\t".join(parts) + "\n"
            result.append(line)

    return "".join(result) if header_found else None


def write_file(data_string):
    with open(os.path.join(os.getcwd(), "data.csv"), "w", encoding="cp866") as f:
        f.write(data_string)


def read_dataframe():
    df = pd.read_csv(
        os.path.join(os.getcwd(), "data.csv"), delimiter="\t", encoding="cp866"
    )
    os.remove(os.path.join(os.getcwd(), "data.csv"))
    return df


def process_dataframe(df):
    processed_data = []
    sequence = ""

    for _, row in df.iterrows():
        if row["key"] == "Reset":
            tmofday = str(row["tmofday"]) if not pd.isna(row["tmofday"]) else "N/A"
            sequence = tmofday + ": "
        elif row["key"] in {"Num1", "Num2", "Num3"}:
            sequence += row["key"].replace("Num", "")
        elif row["key"] == "Exit":
            processed_data.append(sequence)

    return processed_data


def calculate_metrics(temp_result):
    metrics = {}
    for el in temp_result:
        t = el.split(": ")[1][::2]
        triplets = sum(
            t[i : i + 3] in {"123", "132", "213", "231", "312", "321"}
            for i in range(len(t) - 2)
        )
        returns = sum(t[i : i + 2] in {"11", "22", "33"} for i in range(len(t) - 1))
        rollback = sum(
            t[i : i + 3] in {"121", "131", "212", "232", "313", "323"}
            for i in range(len(t) - 2)
        )
        activity = len(t)
        efficiency = (
            round((triplets / (activity - 2)) * 100, 1) if activity > 2 else 0.0
        )

        key = el.split(": ")[0]
        metrics[key] = [triplets, rollback, returns, activity, efficiency]

    return metrics


def create_temp_dataframe(metrics):
    return pd.DataFrame.from_dict(
        metrics,
        orient="index",
        columns=[
            "Триплеты",
            "Возвраты",
            "Обратные возвраты",
            "Двигательная активность",
            "Эффективность патрулирования",
        ],
    )


def save_to_excel():
    try:
        file_path = filedialog.asksaveasfile(
            mode="wb",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Выберите путь сохранения",
            initialfile="Лист Microsoft Excel.xlsx",
        )

        if file_path:
            save_path = file_path.name

            result_df.to_excel(save_path, index=True)
            adjust_column_width(save_path)

            open_file = messagebox.askyesno(
                "Открыть файл?", "Хотите открыть сохраненный файл?"
            )

            if open_file:
                open_saved_file(save_path)

    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка при записи в файл: {e}")


def adjust_column_width(file_path):
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active

    for column in sheet.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        sheet.column_dimensions[column_letter].width = max_length + 2

    wb.save(file_path)


def open_saved_file(file_path):
    try:
        if sys.platform.startswith("win32"):
            os.startfile(file_path)
        elif sys.platform.startswith("darwin"):
            subprocess.run(["open", file_path])
        elif sys.platform.startswith("linux"):
            subprocess.run(["xdg-open", file_path])
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось открыть файл: {e}")


def clear_list(listbox, result_text, save_button, clear_button):
    global result_df, files, file_names
    result_df = pd.DataFrame()
    files = []
    file_names = []
    listbox.delete(0, "end")
    update_result_text(result_text, save_button)
    result_text.delete(*result_text.get_children())
    save_button.config(state="disabled")
    clear_button.config(state="disabled")


def delete_file(listbox, result_text, save_button, clear_button):
    selected_item = listbox.curselection()
    if selected_item:
        index = selected_item[0]
        del files[index]
        del file_names[index]
        listbox.delete(selected_item)
        update_result_text(result_text, save_button)
        result_text.delete(*result_text.get_children())
        process_file(result_text, save_button)

    if listbox.size() == 0:
        save_button.config(state="disabled")
        clear_button.config(state="disabled")


def update_result_text(result_text, save_button):
    result_text.delete(*result_text.get_children())

    for index, row in result_df.iterrows():
        result_text.insert(
            "",
            "end",
            text=str(row.name),
            values=(
                int(row["Триплеты"]),
                int(row["Возвраты"]),
                int(row["Обратные возвраты"]),
                int(row["Двигательная активность"]),
                round(row["Эффективность патрулирования"], 2),
            ),
        )

    save_button.config(state="normal" if not result_df.empty else "disabled")
