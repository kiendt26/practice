# coding=utf-8
import os
import pandas as pd

def get_desktop_path():
    home = os.path.expanduser("~")  # Lấy đường dẫn tới thư mục home
    desktop = os.path.join(home, "Desktop")  # Nối thêm "Desktop" vào đường dẫn home
    return desktop

def main():
    print("Init")
    # Lấy tên tệp từ người dùng
    file_name = input("Nhập tên file trên Desktop (Mặc định là auto.xlsx): ")
    if file_name == '':
        file_name = 'auto.xlsx'
    # Lấy đường dẫn đầy đủ tới tệp

    desktop_path = get_desktop_path()
    file_path = os.path.join(desktop_path, file_name)
    print (file_path)
    # Đọc dữ liệu từ file Excel
    sheet1 = pd.read_excel(file_path, sheet_name='Sheet1')
    sheet2 = pd.read_excel(file_path, sheet_name='Sheet2')

    # Kiểm tra và tạo cột "Số lượng" nếu chưa tồn tại
    if 'Số lượng' not in sheet1.columns:
        sheet1['Số lượng'] = pd.Series(dtype=int)

    # Biến để lưu kết quả không tìm thấy
    khong_tim_thay = []

    # So sánh và cập nhật số lượng từ sheet2 vào sheet1
    for idx2, row2 in sheet2.iterrows():
        test = str(row2['Test']).strip()  # Lấy tên sản phẩm từ cột "Test" của sheet2 và loại bỏ khoảng trắng
        sub_test = str(row2.get('Sub-Test', '')).strip()  # Lấy tên sản phẩm từ cột "Sub-Test" và loại bỏ khoảng trắng
        sub_sub_test = str(
            row2.get('Sub-Sub-Test', '')).strip()  # Lấy tên sản phẩm từ cột "Sub-Sub-Test" và loại bỏ khoảng trắng
        so_luong = row2['Retest Count']  # Lấy số lượng từ cột "Retest Count" của sheet2
        # Chỉ cập nhật nếu số lượng lớn hơn 0
        if so_luong > 0:
            # Biến cờ để kiểm tra xem có tìm thấy tên đầy đủ phù hợp hay không
            found = False
            offset=0
            ten_day_du=''
            # Tìm hàng tương ứng trong sheet1 mà chứa các tên rút gọn
            for idx1, row1 in sheet1.iterrows():
                ten_day_du = row1['name']  # Lấy tên đầy đủ từ cột "name" của sheet1

                #Kiểm tra độ dài của tên đầy đủ và tổng độ dài test+sub_test+sub_sub_test không vượt quá 5 ký tự
                offset= len(ten_day_du.strip())-len((test+sub_test+sub_sub_test).strip())
                if not offset <-7 or offset>5:
                    # Kiểm tra nếu tên đầy đủ chứa các tên rút gọn
                    if test in ten_day_du:
                        if not sub_test=='nan'  and sub_test not in ten_day_du:
                            continue
                        if not sub_sub_test=='nan' and sub_sub_test not in ten_day_du:
                            continue
                        sheet1.at[idx1, 'Số lượng'] = so_luong  # Cập nhật số lượng trong sheet1
                        found = True  # Đặt cờ thành True khi tìm thấy kết quả phù hợp
                        break  # Thoát khỏi vòng lặp khi tìm thấy kết quả phù hợp

            # Nếu không tìm thấy tên đầy đủ phù hợp, ghi lại thông tin
            if not found:
                khong_tim_thay.append(test + "^^" + sub_test + "^^" + sub_sub_test + " " + str(so_luong))

    # Ghi dữ liệu cập nhật vào file Excel
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        sheet1.to_excel(writer, sheet_name='Sheet1', index=False)

    # In ra các kết quả không tìm thấy
    for item in khong_tim_thay:
        print(item)

    print("Cập nhật hoàn tất!")

if _name_ == '_main_':
    main()