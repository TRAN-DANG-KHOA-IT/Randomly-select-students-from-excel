import pandas as pd
import random
import time
from colorama import Fore, Style, Back

# Hàm để đọc dữ liệu từ tệp Excel
def read_excel_data(file_path):
    try:
        data = pd.read_excel(file_path)
        # Chuyển dữ liệu thành danh sách hai chiều
        data_list = data.values.tolist()
        return data_list
    except Exception as e:
        print("Lỗi khi đọc tệp Excel:", e)
        return None

# Hàm rand màu
def rand_color():
    colors = [Fore.RED, Fore.GREEN, Fore.YELLOW, Fore.MAGENTA, Fore.CYAN, Fore.WHITE]
    return random.choice(colors)

# Hàm để quay số ngẫu nhiên từ dữ liệu
def spin_wheel(data):
    if data is None:
        return None
    else:
        # Tạo hiệu ứng chờ đợi
        print(Fore.GREEN+"Đang quay số ", end="")
        for _ in range(10):
            print(rand_color() + ".", end="", flush=True)
            time.sleep(0.2)
        print("", end="\r")
        print("                                                             "+Fore.RESET, end= "\r")
        sl = []
        for i in range(0,len(data),1):
            sl.append(i)
        randnumber = random.choice(sl);
        info = data[randnumber]
        result = f"{Fore.GREEN}Số thứ tự: {rand_color()}{info[0]}{Fore.RESET} - {Fore.GREEN}Tên học sinh: {rand_color()}{info[1]}{Fore.RESET} - {Fore.GREEN}Ngày sinh: {rand_color()}{info[2]}{Fore.RESET} -  {Fore.GREEN}Lớp: {rand_color()}{info[3]}{Fore.RESET}"
        return result

# Hàm main
def main():
    stt = 0
    print(Fore.RESET + Style.RESET_ALL, Back.RESET,end="\r")
    file_path = "12A9.xlsx"
    data = read_excel_data(file_path)
    if data is not None:
        print(Back.GREEN+"Dữ liệu đã được tải thành công từ tệp Excel." + Back.RESET)
        while True:
            nhap = input(Style.NORMAL+Fore.LIGHTCYAN_EX+"Nhấn Enter để quay số...")
            if nhap == "die": break
            result = spin_wheel(data)
            if result is not None:
                stt+=1
                print(f"Kết quả quay lần {stt} là:", end=" ")
                # Sử dụng màu sắc
                print(Fore.GREEN + str(result)+ Style.RESET_ALL + Fore.RESET)
            else:
                print("Không thể quay số. Dữ liệu trống.")
    else: 
        print(Back.RED + "Dữ liệu lấy từ file Excel lỗi hãy báo admin!" +Back.RESET)

if __name__ == "__main__":
    main()
