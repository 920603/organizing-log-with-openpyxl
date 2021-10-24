"""
TODO:
    컬럼 (속도, 차로 편측 등) 선택 기능 => 2D listbox?? 
    그래프 출력 => checkbox

CAVEAT:
    시나리오 이름에 _ (언더스코어)를 사용하면 안됨
"""

import os
import tkinter as tk
import tkinter.filedialog as fd
import tkinter.messagebox as msgbox
from util.generator import Generator

root = tk.Tk()
root.title("로그 데이터 정리 툴")
root.geometry("500x600")

frame_1 = tk.Frame(root)
frame_1.pack(fill="x", padx=10, pady=5)


def load_file():
    file_listbox.delete(0, tk.END)

    file_paths = fd.askopenfilenames(filetypes=[("CSV file", "*.csv")])

    for file_path in file_paths:
        file_listbox.insert(tk.END, file_path)


load_file_btn = tk.Button(frame_1, text="로그 파일 열기", command=load_file)
load_file_btn.pack(side="left")


def delete_file():
    for index in reversed(file_listbox.curselection()):
        file_listbox.delete(index)


delete_file_btn = tk.Button(frame_1, text="선택 삭제", command=delete_file)
delete_file_btn.pack(side="right")


frame_2 = tk.Frame(root)
frame_2.pack(fill="x", padx=10, pady=5)

frame_2_top_frame = tk.Frame(frame_2)
frame_2_top_frame.pack(fill="x")

scrollbar_y = tk.Scrollbar(frame_2_top_frame)
scrollbar_x = tk.Scrollbar(frame_2, orient="horizontal")
file_listbox = tk.Listbox(
    frame_2_top_frame,
    selectmode="extended",
    height=10,
    yscrollcommand=scrollbar_y.set,
    xscrollcommand=scrollbar_x.set,
)
scrollbar_y.config(command=file_listbox.yview)
scrollbar_x.config(command=file_listbox.xview)

file_listbox.pack(side="left", expand=True, fill="x")
scrollbar_y.pack(side="right", fill="y")
scrollbar_x.pack(fill="x")

# frame_3 = tk.Frame(root)
# frame_3.pack(fill="x", padx=10, pady=5)

# dest_path_frame = tk.LabelFrame(frame_3, text="저장 경로")
# dest_path_frame.pack(fill="x")

# dest_path_entry = tk.Entry(
#     dest_path_frame,
# )
# dest_path_entry.pack(
#     side="left", fill="x", expand=True, padx=5, pady=5, ipadx=3, ipady=3
# )


# dest_path_btn = tk.Button(dest_path_frame, text="저장 경로 선택", command=select_destination)
# dest_path_btn.pack(side="right", padx=5, pady=5)


frame_4 = tk.Frame(root)
frame_4.pack(fill="x", padx=10, pady=5)

frame_4_top_frame = tk.Frame(frame_4)
frame_4_top_frame.pack(fill="x")

frame_4_bottom_frame = tk.Frame(frame_4)
frame_4_bottom_frame.pack(fill="x")

starting_point_frame = tk.LabelFrame(
    frame_4_top_frame, text="분석 시점 distanceTravelled (m)"
)
starting_point_frame.pack(fill="x", side="left", expand=True)

starting_point_entry = tk.Entry(starting_point_frame)
starting_point_entry.pack(fill="x", expand=True, padx=5, pady=5, ipadx=3, ipady=3)

ending_point_frame = tk.LabelFrame(
    frame_4_top_frame, text="분석 종점 distanceTravelled (m)"
)
ending_point_frame.pack(fill="x", side="right", expand=True)

ending_point_entry = tk.Entry(ending_point_frame)
ending_point_entry.pack(fill="x", expand=True, padx=5, pady=5, ipadx=3, ipady=3)

starting_station_frame = tk.LabelFrame(frame_4_bottom_frame, text="분석 시점 Station (km)")
starting_station_frame.pack(fill="x", side="left", expand=True)

starting_station_entry = tk.Entry(starting_station_frame)
starting_station_entry.pack(fill="x", expand=True, padx=5, pady=5, ipadx=3, ipady=3)

frequency_frame = tk.LabelFrame(frame_4_bottom_frame, text="분석점 빈도 (m)")
frequency_frame.pack(fill="x", side="right", expand=True)

frequency_entry = tk.Entry(frequency_frame)
frequency_entry.pack(fill="x", expand=True, padx=5, pady=5, ipadx=3, ipady=3)


frame_5 = tk.Frame(root)
frame_5.pack()


def select_destination():
    selected_path = fd.asksaveasfilename(
        filetypes=[("XLSX file", "*.xlsx")], defaultextension="xlsx"
    )
    return selected_path


def start():
    if file_listbox.size() == 0:
        msgbox.showwarning(message="로그 파일을 선택해주세요")
        return

    if not starting_point_entry.get().strip().isdecimal():
        msgbox.showwarning(
            message="분석 시작점의 distanceTravelled 값을 미터 단위의 자연수로 입력하세요\n(예. 500)"
        )
        return

    if not ending_point_entry.get().strip().isdecimal():
        msgbox.showwarning(
            message="분석 종료점의 distanceTravelled 값을 미터 단위의 자연수로 입력하세요\n(예. 1400)"
        )
        return

    if not starting_station_entry.get().strip().replace(".", "", 1).isdecimal():
        msgbox.showwarning(message="분석 시작점의 station 값을 킬로미터 단위로 입력하세요\n(예. 14.2)")
        return

    if not frequency_entry.get().strip().isdecimal():
        msgbox.showwarning(message="분석점의 빈도 값을 미터 단위로 입력하세요\n(예. 10)")
        return

    # file manipulation script here
    generator = Generator(
        file_listbox.get(0, tk.END),
        starting_point_entry.get(),
        ending_point_entry.get(),
        starting_station_entry.get(),
        frequency_entry.get(),
    )
    wb = generator.generate_workbook()
    dest = select_destination()
    wb.save(dest)

    # os.startfile(os.path.realpath(dest_path_entry.get().strip()))
    ## change `progress_var` and do `progress_bar.update()`


start_btn = tk.Button(frame_5, text="시작", command=start)
start_btn.pack(side="left", padx=5, pady=5)

quit_btn = tk.Button(frame_5, text="종료", command=root.quit)
quit_btn.pack(side="right", padx=5, pady=5)


root.mainloop()
