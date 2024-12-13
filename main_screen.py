# Lê Duy Quân - 3121411176

# Luổng xử lí
# 1. Đăng nhập ( Đăng nhập thành công, Đăng nhập thất bại)
# 2. Main page 
#     - Thêm, Xóa sinh viên
#     - Tìm kiếm
#     - Nhập file Excel ( có thể nhập nhiều file ): Click "Nhập từ Excel" -> Quản lý sinh viên page 
# 3. Quản lí sinh viên page
#     - Table hiển thị các file excel vừa import 
#     - Nut "Sắp xếp theo tổng buổi vắng" -> Phân loại sinh viên vắng theo buổi và các chức năng (1)
#     - Nút "Hiện thị ngày nghỉ" -> Click chọn sinh viên và Click "Hiển thị ngày nghỉ" -> Hiển thị số ngày nghỉ và ngày nghỉ

# 4. (1)
#     - Gửi mail báo cáo: Click chọn sinh viên ở table mới hiện -> Click "Gửi Email Báo Cáo" -> A+( không nghỉ ),A (nghỉ 1 buổi) , B (nghỉ 2 buổi) ,....
#                                                                                             A+,A: Không gửi mail
#                                                                                             B+: Gửi mail sinh viên
#                                                                                             B->D: Gửi mail sinh viên, phụ huynh, gvcn và TBM
      
#     - Tạo File Excel: Lọc thông tin và gửi file excel đến email nhân viên và quản lí ( mỗi tháng gửi 2 lần )
#     - Đọc danh sách mail và chắt lọc thông tin ( Tổng hợp lớp nào gửi và người gửi là ai, Nếu phát hiện deadline đã trễ hạn thì gửi mail cho quanly )
#     - Thêm Q&A/ Chat Box (2)

# 5. (2)
#     - Nhập câu hỏi và câu trả lời -> Nếu mail gửi tới có trong danh sách câu hỏi thì sẻ gửi mail trả lời tự dồng, Nếu không có trong danh sách thì gửi thông tin tới email phụ trách )
#     - Nút "Kiểm tra Email" -> Check mail và trả lời tự động



import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkinter import scrolledtext
from db import Database
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import smtplib
import schedule
import time
import threading
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime, timedelta
import imaplib
import email
from email.header import decode_header
import re

class MainScreen:
    def __init__(self, root):
        self.root = root
        self.db = Database()

        # 
        # 
        self.can_send_email = True  
        self.last_email_sent = datetime.now() - timedelta(minutes=2)  
        # self.start_scheduler()  
        # self.last_email_sent = datetime.now() - timedelta(days=30)  
        # self.start_scheduler()  
        # 
        # 


        self.staff_accounts = {
            "staff1": ("leduyquan2574@gmail.com", "lajewrwnkozpmulx"),
            # Thêm tài khoản nhân viên khác nếu có
        }

        self.setup_ui()
        self.load_data()

    def setup_ui(self):
        self.root.title("Quản lý sinh viên")
        self.root.geometry("1400x600")

        # Center the window on the screen
        window_width = 1400
        window_height = 600

        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)

        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")

        # Frame main page
        self.student_frame = tk.Frame(self.root)
        self.student_frame.pack(fill="both", expand=True)

        # Search frame on the top left
        self.search_frame = tk.Frame(self.student_frame)
        self.search_frame.pack(side="top", pady=10)

        self.search_entry = tk.Entry(self.search_frame,width= 50)
        self.search_entry.pack(side="left", padx=5)

        self.search_button = tk.Button(self.search_frame, text="Tìm kiếm", command=self.search_student, bg="lightcoral", fg="black", width=20)
        self.search_button.pack(side="left", padx=5)

        # Button frame on the right
        self.button_frame = tk.Frame(self.student_frame)
        self.button_frame.pack(side="right", padx=10)

        button_width = 20

        self.refresh_button = tk.Button(self.button_frame, text="Tải lại dữ liệu", command=self.load_data, bg="lightblue", fg="black", width=button_width)
        self.refresh_button.grid(row=0, column=0, pady=5)

        self.add_button = tk.Button(self.button_frame, text="Thêm sinh viên", command=self.add_student, bg="lightgreen", fg="black", width=button_width)
        self.add_button.grid(row=1, column=0, pady=5)

        self.delete_button = tk.Button(self.button_frame, text="Xóa sinh viên", command=self.delete_student, bg="salmon", fg="black", width=button_width)
        self.delete_button.grid(row=2, column=0, pady=5)

        self.import_button = tk.Button(self.button_frame, text="Nhập từ Excel", command=self.import_from_excel, bg="lightyellow", fg="black", width=button_width)
        self.import_button.grid(row=3, column=0, pady=5)

        # Treeview to display the list of students
        self.tree = ttk.Treeview(self.student_frame, columns=("MSSV", "Họ tên", "Lớp", "Môn học", "Số buổi vắng", "Ngày nghỉ"), show="headings")
        for col in self.tree["columns"]:
            self.tree.heading(col, text=col)
        self.tree.pack(fill="both", expand=True)

# 
# 
# Frame để hiển thị dữ liệu từ Excel

        self.imported_frame = tk.Frame(self.root)
        self.imported_frame.pack(fill="both", expand=True)


        self.imported_tree = ttk.Treeview(self.imported_frame, columns=("STT", "MSSV", "Họ đệm", "Tên", "Giới tính", "Ngày sinh", "Lớp", "Môn", "[Thứ ba] - [7->11] - 11/06/2024", "[Thứ ba] - [7->11] - 18/06/2024", "[Thứ ba] - [7->11] - 25/06/2024", "[Thứ ba] - [7->11] - 02/07/2024", "[Thứ ba] - [7->11] - 07/07/2024", "[Thứ ba] - [7->11] - 23/07/2024"), show="headings")
        for col in self.imported_tree["columns"]:
            self.imported_tree.heading(col, text=col)
        self.imported_tree.pack(fill="both", expand=True)

        self.imported_frame.pack_forget()

        button_frame = tk.Frame(self.imported_frame)
        button_frame.pack(pady=10)  

        self.sort_button = tk.Button(button_frame, text="Sắp xếp theo tổng buổi vắng", command=self.sort_imported_data, bg="lightgrey", fg="black", width=20)
        self.sort_button.grid(row=0, column=0, padx=5) 

        back_button = tk.Button(button_frame, text="Quay lại", command=self.back_to_student_view, bg="lightblue", fg="black", width=20)
        back_button.grid(row=0, column=1, padx=5)  

        show_holidays_button = tk.Button(button_frame, text="Hiển thị ngày nghỉ", command=self.show_absence_info, bg="lightgreen", fg="black", width=20)
        show_holidays_button.grid(row=0, column=2, padx=5)  

        show_holidays_button = tk.Button(button_frame, text="Sắp xếp", command=self.sort_classification_data, bg="lightgreen", fg="black", width=20)
        show_holidays_button.grid(row=0, column=3, padx=5)  

        self.search_entry = tk.Entry(button_frame, width=30)
        self.search_entry.grid(row=0, column=4, padx=5)  # Thêm ô nhập liệu vào cột 4

        self.search_button = tk.Button(button_frame, text="Tìm kiếm", command=self.search_students, bg="lightblue", fg="black")
        self.search_button.grid(row=0, column=5, padx=5)  # Thêm nút tìm kiếm vào cột 5

        # Combobox để chọn phân loại
        self.filter_label = tk.Label(button_frame, text="Chọn phân loại:")
        self.filter_label.grid(row=1, column=0, padx=5, pady=5)

        self.filter_combobox = ttk.Combobox(button_frame, values=["Tất cả", "A+", "A", "B+", "B", "C+", "C", "D"])
        self.filter_combobox.current(0)  # Mặc định chọn "Tất cả"
        self.filter_combobox.grid(row=1, column=1, padx=5, pady=5)

        # Nút lọc
        self.filter_button = tk.Button(button_frame, text="Lọc", command=self.filter_classification_data, bg="lightblue", fg="black", width=button_width)
        self.filter_button.grid(row=1, column=2, padx=5, pady=5)



# 
# 
# Funcion xử lí Excel và lưu vào database

    def show_imported_data(self, students_data): #Function show dữ liệu database từ excel
        self.student_frame.pack_forget()
        self.imported_frame.pack(fill="both", expand=True)

        for row in self.imported_tree.get_children():
            self.imported_tree.delete(row)

        for _, row in students_data.iterrows():
            values = [
                row['STT'],                  
                row['Mã sinh viên'],         
                row['Họ đệm'],               
                row['Tên'],                 
                row['Giới tính'],            
                row['Ngày sinh'],            
                row['Lớp'],                 
                row['Môn'],                  
                row['[Thứ ba] - [7->11] - 11/06/2024 (P/K)'],
                row['[Thứ ba] - [7->11] - 18/06/2024 (P/K)'],
                row['[Thứ ba] - [7->11] - 25/06/2024 (P/K)'],
                row['[Thứ ba] - [7->11] - 02/07/2024 (P/K)'],
                row['[Thứ ba] - [7->11] - 09/07/2024 (P/K)'],
                row['[Thứ ba] - [7->11] - 23/07/2024 (P/K)'],
            ]
            
            self.imported_tree.insert("", "end", values=values)





    def import_from_excel(self): #Func import dữ liệu từ database
        column_names = ['STT', 'Mã sinh viên', 'Họ đệm', 'Tên', 'Giới tính', 'Ngày sinh',
                        '[Thứ ba] - [7->11] - 11/06/2024 (P/K)', '[Thứ ba] - [7->11] - 11/06/2024 ST', '[Thứ ba] - [7->11] - 11/06/2024 LD',
                        '[Thứ ba] - [7->11] - 18/06/2024 (P/K)', '[Thứ ba] - [7->11] - 18/06/2024 ST', '[Thứ ba] - [7->11] - 18/06/2024 LD',
                        '[Thứ ba] - [7->11] - 25/06/2024 (P/K)', '[Thứ ba] - [7->11] - 25/06/2024 ST', '[Thứ ba] - [7->11] - 25/06/2024 LD',
                        '[Thứ ba] - [7->11] - 02/07/2024 (P/K)', '[Thứ ba] - [7->11] - 02/07/2024 ST', '[Thứ ba] - [7->11] - 02/07/2024 LD',
                        '[Thứ ba] - [7->11] - 09/07/2024 (P/K)', '[Thứ ba] - [7->11] - 09/07/2024 ST', '[Thứ ba] - [7->11] - 09/07/2024 LD',
                        '[Thứ ba] - [7->11] - 23/07/2024 (P/K)', '[Thứ ba] - [7->11] - 23/07/2024 ST', '[Thứ ba] - [7->11] - 23/07/2024 LD',
                        'Vắng có phép', 'Vắng không phép', 'Tổng số tiết', '(%) vắng']  

        file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xls;*.xlsx")])
        if file_paths:
            all_students_data = pd.DataFrame()  

            for file_path in file_paths:
                try:
                    class_number = pd.read_excel(file_path, header=None, usecols='C', nrows=10).iloc[9, 0]
                    subject = pd.read_excel(file_path, header=None, usecols='C', nrows=9).iloc[8, 0] 

                    students_data = pd.read_excel(file_path, header=None, skiprows=13, index_col=None)
                    students_data.columns = column_names  
                    students_data.reset_index(drop=True, inplace=True)

                    students_data['Lớp'] = class_number
                    students_data['Môn'] = subject  

                    all_students_data = pd.concat([all_students_data, students_data], ignore_index=True)

                except Exception as e:
                    messagebox.showerror("Lỗi", f"Có lỗi khi nhập dữ liệu từ tệp {file_path}: {e}")

            if not all_students_data.empty:
                self.save_to_database(all_students_data)
                self.show_imported_data(all_students_data)



    def save_to_database(self, students_data): #Func save dữ liệu từ excel
        print("Dữ liệu sinh viên:", students_data)
        
        for i, row in students_data.iterrows():
            stt = row['STT']

            ngay_sinh = row['Ngày sinh'].strftime('%Y-%m-%d') if isinstance(row['Ngày sinh'], pd.Timestamp) else row['Ngày sinh']

            lop = row['Lớp'] if not pd.isna(row['Lớp']) else None

            mon = row['Môn'] if not pd.isna(row['Môn']) else None  

            values = [
                stt,
                row['Mã sinh viên'] ,
                row['Họ đệm'] ,
                row['Tên'] ,
                row['Giới tính'] ,
                ngay_sinh ,
                lop ,  
                mon,  
                row['[Thứ ba] - [7->11] - 11/06/2024 (P/K)'],
                row['[Thứ ba] - [7->11] - 18/06/2024 (P/K)'] ,
                row['[Thứ ba] - [7->11] - 25/06/2024 (P/K)'] ,
                row['[Thứ ba] - [7->11] - 02/07/2024 (P/K)'],
                row['[Thứ ba] - [7->11] - 09/07/2024 (P/K)'] ,
                row['[Thứ ba] - [7->11] - 23/07/2024 (P/K)'] ,
            ]


            try:
                self.db.connection.execute(""" 
                    INSERT INTO absences (stt, mssv, ho_dem, ten, gioi_tinh, ngay_sinh, lop, mon,
                    thu_1_status, thu_2_status,
                    thu_3_status,  thu_4_status, 
                    thu_5_status,  thu_6_status)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,?)""",
                    values)
            except Exception as e:
                print(f"Lỗi khi chèn hàng: {e}")

        self.db.connection.commit()

# 
# 
# Send mail

    def send_email(self, to_email, subject, body): #Function send mail
        from_email = "leduyquan2574@gmail.com"  
        from_password = "lajewrwnkozpmulx"        

        to_email = str(to_email)
         
        msg = MIMEMultipart()
        msg['From'] = from_email
        msg['To'] = to_email
        msg['Subject'] = subject

        msg.attach(MIMEText(body, 'plain'))

        try:
            with smtplib.SMTP('smtp.gmail.com', 587) as server:  # SMTP server
                server.starttls()
                server.login(from_email, from_password)
                server.send_message(msg)
            print(f"Email sent to {to_email}")
        except Exception as e:
            print(f"Failed to send email to {to_email}: {e}")



# Show amount of off day' students when click "Hiển thị số ngày nghỉ"
# 
# 
    def show_absence_info(self):
        selected_item = self.imported_tree.selection()
        if selected_item:
            mssv = self.imported_tree.item(selected_item)["values"][1]  # Lấy MSSV từ cột thứ 1
            absence_info = self.db.get_absence_info(mssv)

            if absence_info:
                # Tạo từ điển cho ngày nghỉ với tên ngày cụ thể
                absence_dict = {
                    "[Thứ hai] - [7->11] - 11/06/2024": absence_info.get("thu_1_status", ""),
                    "[Thứ ba] - [7->11] - 18/06/2024": absence_info.get("thu_2_status", ""),
                    "[Thứ tư] - [7->11] - 25/06/2024": absence_info.get("thu_3_status", ""),
                    "[Thứ năm] - [7->11] - 02/07/2024": absence_info.get("thu_4_status", ""),
                    "[Thứ sáu] - [7->11] - 09/07/2024": absence_info.get("thu_5_status", ""),
                    "[Thứ bảy] - [7->11] - 23/07/2024": absence_info.get("thu_6_status", ""),
                }

                # Tính số ngày nghỉ
                absence_count = sum(1 for status in absence_dict.values() if status != '')
                absence_str = "\n".join([f"{day}: {status}" for day, status in absence_dict.items() if status != ''])

                # Hiển thị thông tin
                if absence_str:  # Kiểm tra nếu có trạng thái vắng
                    messagebox.showinfo("Thông tin ngày nghỉ", f"Các trạng thái vắng của sinh viên {mssv}:\n{absence_str}\n\nTổng số ngày nghỉ: {absence_count}")
                else:
                    messagebox.showinfo("Thông tin ngày nghỉ", f"Sinh viên {mssv} không có ngày nghỉ nào.")
            else:
                messagebox.showinfo("Thông tin ngày nghỉ", f"Sinh viên {mssv} không có ngày nghỉ nào.")
        else:
            messagebox.showwarning("Thông báo", "Vui lòng chọn sinh viên để hiển thị thông tin ngày nghỉ.")

# 
# 
# Function hiển thị Phân loại sinh viên

    def sort_imported_data(self):
        students = []
        for row in self.imported_tree.get_children():
            values = self.imported_tree.item(row)["values"]
            students.append(values)

        sorted_students = sorted(students, key=lambda x: (x[6], x[3], x[7], x[8])) 

        # Lưu trữ dữ liệu gốc
        self.original_students_data = sorted_students.copy()  # Lưu dữ liệu đã sắp xếp vào biến gốc

        # Kiểm tra xem classification_frame đã tồn tại chưa
        if hasattr(self, 'classification_frame'):
            self.classification_frame.pack_forget()  # Ẩn frame cũ nếu đã tồn tại
        self.classification_frame = tk.Frame(self.imported_frame)
        self.classification_frame.pack(fill="both", expand=True)

        columns = ("Họ và Tên", "MSSV", "Lớp", "Môn", "Tổng buổi nghỉ", "Phân loại")
        self.classification_tree = ttk.Treeview(self.classification_frame, columns=columns, show="headings")

        self.classification_tree.column("Họ và Tên", anchor="w", width=150)
        self.classification_tree.column("MSSV", anchor="w", width=150)
        self.classification_tree.column("Lớp", anchor="center", width=100)
        self.classification_tree.column("Môn", anchor="center", width=100)
        self.classification_tree.column("Tổng buổi nghỉ", anchor="center", width=100)
        self.classification_tree.column("Phân loại", anchor="center", width=100)

        for col in columns:
            self.classification_tree.heading(col, text=col)

        scrollbar = ttk.Scrollbar(self.classification_frame, orient="vertical", command=self.classification_tree.yview)
        self.classification_tree.configure(yscroll=scrollbar.set)

        self.classification_tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        for student in sorted_students:
            name = f"{student[2]} {student[3]}"  
            class_name = student[6]                
            mssv = student[1]                      
            subject = student[7]                   
            absence_info = self.db.get_absence_info(mssv)  

            absence_count = 0

            if absence_info:
                absence_dict = {
                    "[Thứ hai] - [7->11] - 11/06/2024": absence_info.get("thu_1_status", ""),
                    "[Thứ ba] - [7->11] - 18/06/2024": absence_info.get("thu_2_status", ""),
                    "[Thứ tư] - [7->11] - 25/06/2024": absence_info.get("thu_3_status", ""),
                    "[Thứ năm] - [7->11] - 02/07/2024": absence_info.get("thu_4_status", ""),
                    "[Thứ sáu] - [7->11] - 09/07/2024": absence_info.get("thu_5_status", ""),
                    "[Thứ bảy] - [7->11] - 23/07/2024": absence_info.get("thu_6_status", ""),
                }

                absence_count = sum(1 for status in absence_dict.values() if status != '')

                if absence_count == 0:
                    grade = "A+"
                elif absence_count == 1:
                    grade = "A"
                elif absence_count == 2:
                    grade = "B+"
                elif absence_count == 3:
                    grade = "B"
                elif absence_count == 4:
                    grade = "C+"
                elif absence_count == 5:
                    grade = "C"
                elif absence_count >= 6:
                    grade = "D"
                else:
                    grade = "Không xác định"
            else:
                absence_count = 0
                grade = "A+"

            self.classification_tree.insert("", "end", values=(name, mssv, class_name, subject, absence_count, grade))

        # Thêm các nút chức năng
        send_email_button = tk.Button(self.classification_frame, text="Gửi Email Báo cáo", width=20, command=self.send_email_to_selected_student, bg="lightblue", fg="black")
        send_email_button.pack(pady=10)

        create_excel_button = tk.Button(self.classification_frame, text="Tạo File Excel", width=20, command=self.create_excel_file_ui, bg="lightgreen", fg="black")
        create_excel_button.pack(pady=10)

        recieve_mail = tk.Button(self.classification_frame, text="Check Mail Cá Nhân", width=20, command=self.receive_mail, bg="lightyellow", fg="black")
        recieve_mail.pack(pady=10)

        check_button = tk.Button(self.classification_frame, text="Kiểm tra email của Staff", width=20, command=self.check_staff_emails, bg="lightcoral", fg="black")
        check_button.pack(pady=20)

        self.open_input_button = tk.Button(self.classification_frame, text="Thêm Q&A / Chat Box", width=20, command=self.open_input_window, bg="lightcoral", fg="black")
        self.open_input_button.pack(pady=10)


# Send mail normal (selected by me)
# 
# 
    def send_email_to_selected_student(self):
        selected_item = self.classification_tree.selection()  
        if not selected_item:
            tk.messagebox.showwarning("Cảnh báo", "Vui lòng chọn sinh viên để gửi email.")
            return

        item_values = self.classification_tree.item(selected_item[0])["values"]
        name = item_values[0]  # Tên sinh viên
        class_name = item_values[2]  # Lớp
        absence_count = item_values[4]  # Số buổi vắng
        mssv = item_values[1]  # MSSV

        student_email = f"{mssv}@gmail.com"
        parent_email = f"ph_{mssv}@gmail.com"  
        gvcn_email = f"gvcn_{class_name}@gmail.com"  
        tbm_email = f"tbm_{class_name}@gmail.com"  

        email_subject = "Thông báo kết quả học tập"
        email_content = None

        if absence_count == 2:  # Loại B+
            email_content = f"Kính gửi {name},\n\nBạn đã đạt loại B+. Vui lòng theo dõi kết quả học tập."
            self.send_email(student_email, email_subject, email_content)
            tk.messagebox.showinfo("Thông báo", f"Đã gửi email cho sinh viên {name} (B+).")

        elif absence_count >= 3:  # Loại B, C+, C, D
            email_content = f"Kính gửi {name},\n\nBạn đã đạt loại B/C+/C/D. Vui lòng liên hệ với phụ huynh và giáo viên để cải thiện."
            self.send_email(student_email, email_subject, email_content)  # Gửi cho sinh viên
            self.send_email(parent_email, email_subject, email_content)   # Gửi cho phụ huynh
            self.send_email(gvcn_email, email_subject, email_content)     # Gửi cho giáo viên chủ nhiệm
            self.send_email(tbm_email, email_subject, email_content)      # Gửi cho giáo viên bộ môn
            tk.messagebox.showinfo("Thông báo", f"Đã gửi email cho sinh viên {name}.")

        else:
            tk.messagebox.showinfo("Thông báo", "Sinh viên loại A+ hoặc A, không gửi email.")

# 
# 
# Các Func main page
    def load_data(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        students = self.db.get_students()
        for student in students:
            self.tree.insert("", "end", values=student)

    def add_student(self):
        self.show_student_form("Thêm sinh viên")

    def delete_student(self):
        selected_item = self.tree.selection()
        if selected_item:
            mssv = self.tree.item(selected_item)["values"][0]
            self.db.delete_student(mssv)
            self.load_data()
        else:
            messagebox.showwarning("Xóa sinh viên", "Vui lòng chọn sinh viên để xóa.")

    def search_student(self):
        search_term = self.search_entry.get().strip().lower()
        if search_term:
            found = False
            for row in self.tree.get_children():
                values = self.tree.item(row)["values"]
                mssv, ho_ten = str(values[0]), values[1]  
                if search_term in mssv.lower() or search_term in ho_ten.lower():
                    self.tree.selection_set(row)
                    self.tree.focus(row)
                    found = True
                    break
            if not found:
                messagebox.showinfo("Kết quả tìm kiếm", "Không tìm thấy sinh viên nào.")

# Assignment 1
# 
# 

    def sort_students(self):
            # Lấy dữ liệu từ cây hiển thị
            students = []
            for row in self.tree.get_children():
                values = self.tree.item(row)["values"]
                students.append(values)

            # Sắp xếp theo tổng số buổi vắng, họ tên và lớp
            sorted_students = sorted(students, key=lambda x: (x[4], x[1], x[2]))  # Sắp xếp theo (Số buổi vắng, Họ tên, Lớp)

            # Cập nhật cây hiển thị với dữ liệu đã sắp xếp
            self.tree.delete(*self.tree.get_children())  # Xóa dữ liệu cũ
            for student in sorted_students:
                self.tree.insert("", "end", values=student)

    def back_to_student_view(self):
        self.imported_frame.pack_forget()
        self.student_frame.pack(fill="both", expand=True)
        self.load_data()

    def show_student_form(self, title):
        form = tk.Toplevel(self.root)
        form.title(title)

        form_width, form_height = 400, 350
        screen_width, screen_height = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        x = (screen_width // 2) - (form_width // 2)
        y = (screen_height // 2) - (form_height // 2)
        form.geometry(f"{form_width}x{form_height}+{x}+{y}")

        # Frame for the form
        form_frame = tk.Frame(form, padx=10, pady=10)
        form_frame.pack(fill="both", expand=True)

        # Title label
        title_label = tk.Label(form_frame, text=title, font=("Helvetica", 16, "bold"), pady=10)
        title_label.grid(row=0, column=0, columnspan=2)

        labels = ["MSSV", "Họ tên", "Lớp", "Môn học", "Số buổi vắng", "Ngày nghỉ"]
        entries = {}

        # Create labels and entry fields with increased padding
        for i, label in enumerate(labels):
            tk.Label(form_frame, text=label, font=("Helvetica", 12)).grid(row=i + 1, column=0, pady=(5, 2), sticky="w")
            entry = tk.Entry(form_frame, width=40, font=("Helvetica", 12))
            entry.grid(row=i + 1, column=1, pady=(5, 5))  # Increased bottom padding for entry fields
            entries[label] = entry

        def submit():
            student_data = {
                'mssv': entries["MSSV"].get(),
                'ho_ten': entries["Họ tên"].get(),
                'lop': entries["Lớp"].get(),
                'mon_hoc': entries["Môn học"].get(),
                'so_buoi_vang': int(entries["Số buổi vắng"].get()),  
                'ngay_nghi': entries["Ngày nghỉ"].get()  
            }
            if all(student_data.values()):
                form.destroy()
                self.add_student_to_db(student_data)
            else:
                messagebox.showwarning("Thông báo", "Vui lòng điền đầy đủ thông tin.")

        # Save button with styling
        submit_button = tk.Button(form_frame, text="Lưu", command=submit, bg="lightblue", fg="black", font=("Helvetica", 12), width=20)
        submit_button.grid(row=len(labels) + 1, column=0, columnspan=2, pady=10, sticky="e")

        # Add a padding frame for aesthetics
        form_frame.pack(padx=20, pady=20)

        form.transient(self.root)
        form.grab_set()
        self.root.wait_window(form)



        def submit():
            student_data = {
                'mssv': entries["MSSV"].get(),
                'ho_ten': entries["Họ tên"].get(),
                'lop': entries["Lớp"].get(),
                'mon_hoc': entries["Môn học"].get(),
                'so_buoi_vang': int(entries["Số buổi vắng"].get()),  # Convert to int
                'ngay_nghi': entries["Ngày nghỉ"].get()  # Keep as string
            }
            if all(student_data.values()):
                form.destroy()
                self.add_student_to_db(student_data)
            else:
                messagebox.showwarning("Thông báo", "Vui lòng điền đầy đủ thông tin.")

        tk.Button(form, text="Lưu", command=submit, bg="lightgreen", fg="black", width=20).grid(row=len(labels), column=1, columnspan=4, pady=10, sticky="e")  # Span across two columns

        form.transient(self.root)  
        form.grab_set() 
        self.root.wait_window(form) 

    def add_student_to_db(self, student_data):
        if student_data:
            self.db.add_student(student_data)
            self.load_data()

# Create Excel and Send Mail to emphoyment and manager
# 
# 
# 
# 
    def get_all_student_data(self):
        query = "SELECT mssv, ho_dem, ten, lop, mon FROM absences"
        
        with self.db.connection:
            result = self.db.connection.execute(query).fetchall()
        
        return [{'mssv': row[0], 'ho_dem': row[1], 'ten': row[2], 'lop': row[3], 'mon': row[4]} for row in result]

    def consolidate_student_data(self):
        student_data = self.get_all_student_data()

        consolidated_data = []
        
        for student in student_data:
            mssv = student['mssv']
            absence_info = self.db.get_absence_info(mssv)
            absence_count = sum(1 for key, value in absence_info.items() if 'status' in key and value == 'Absent')
            absence_days = ', '.join([f"{day}: {status}" for day, status in absence_info.items() if 'status' in day])

            consolidated_data.append({
                'MSSV': mssv,
                'Họ tên': f"{student['ho_dem']} {student['ten']}",
                'Lớp': student['lop'],
                'Môn': student['mon'],
                'Số buổi vắng': absence_count,
                'Ngày nghỉ': absence_days,
            })
        
        return consolidated_data

    def create_excel_file_ui(self):
        self.can_send_email = False 
        consolidated_data = self.consolidate_student_data()
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            df = pd.DataFrame(consolidated_data)
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Thông báo", f"File Excel đã được tạo tại: {file_path}")

            subject = "Báo cáo sinh viên vắng nhiều"
            body = "Xin chào, đính kèm báo cáo sinh viên vắng nhiều."
            self.send_email_with_attachment("nhanvien@example.com", subject, body, file_path)
            self.send_email_with_attachment("quanly@example.com", subject, body, file_path)
            messagebox.showinfo("Thông báo", "Email đã được gửi thành công!")

            self.last_email_sent = datetime.now()
            self.can_send_email = True  

    def send_email_with_attachment(self, to_email, subject, body, file_path):
        from_email = "leduyquan2574@gmail.com"  
        from_password = "lajewrwnkozpmulx"        

        msg = MIMEMultipart()
        msg['From'] = from_email
        msg['To'] = to_email
        msg['Subject'] = subject

        msg.attach(MIMEText(body, 'plain'))

        if file_path:
            try:
                with open(file_path, 'rb') as attachment:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename={file_path.split("/")[-1]}')
                msg.attach(part)
            except Exception as e:
                print(f"Không thể đính kèm file: {e}")

        try:
            with smtplib.SMTP('smtp.gmail.com', 587) as server:  # SMTP server 
                server.starttls()
                server.login(from_email, from_password)
                server.send_message(msg)
            print(f"Email đã được gửi tới {to_email}")
        except Exception as e:
            print(f"Không thể gửi email tới {to_email}: {e}")

    def job(self):
        print("Hàm job đã được gọi.")  

        # Xử lý dữ liệu email để lấy các email bị trễ hạn
        missed_deadlines = self.process_email_data(self.fetch_emails())
        
        # Gửi báo cáo nếu có email trễ hạn
        if missed_deadlines:
            self.send_missed_deadline_report(missed_deadlines)

        # Kiểm tra điều kiện gửi email hàng ngày
        if self.can_send_email and datetime.now() - self.last_email_sent >= timedelta(minutes=1):
            consolidated_data = self.consolidate_student_data()
            file_path = "temp_report.xlsx"  

            # Tạo file báo cáo Excel
            df = pd.DataFrame(consolidated_data)
            df.to_excel(file_path, index=False)

            # Gửi email kèm file báo cáo
            subject = "Báo cáo sinh viên vắng nhiều"
            body = "Xin chào, đính kèm báo cáo sinh viên vắng nhiều."
            self.send_email_with_attachment("nhanvien@example.com", subject, body, file_path)
            self.send_email_with_attachment("quanly@example.com", subject, body, file_path)

            # Cập nhật thời gian gửi email cuối cùng
            self.last_email_sent = datetime.now()
            print("Email đã được gửi thành công!")

            # Gửi báo cáo định kỳ (nếu có)
            self.send_scheduled_report()


    # def start_scheduler(self):
    #     schedule.every(5).minutes.do(lambda: self.root.after(0, self.job)) 
    #     # schedule.every(15).days.do(lambda: self.root.after(0, self.job)) 

    #     def run_scheduler():
    #         while True:
    #             schedule.run_pending()
    #             time.sleep(1)  

    #     threading.Thread(target=run_scheduler, daemon=True).start()

# last question of Assignment 2
# 
# 

    def receive_mail(self):
        emails = self.fetch_emails()  # Lấy danh sách 10 email gần nhất

        if emails:
            for email_data in emails:
                # Kiểm tra xem email_data có chứa trường 'date' hay không
                date_info = email_data.get('date', 'Không có ngày')
                print(f"Từ: {email_data['from']}, Tiêu đề: {email_data['subject']}, Ngày: {date_info}")

            messagebox.showinfo("Thông báo", f"Đã nhận {len(emails)} email mới.")

            missed_deadlines = self.process_email_data(emails)

            if missed_deadlines:
                self.send_missed_deadline_report(missed_deadlines)

                self.open_missed_deadline_frame(missed_deadlines)

            self.open_email_info_frame(emails)  
        else:
            messagebox.showinfo("Thông báo", "Không có email mới.")

    def fetch_emails(self):
        imap = imaplib.IMAP4_SSL("imap.gmail.com")
        email_address = "leduyquan2574@gmail.com"  
        password = "lajewrwnkozpmulx"  

        try:
            imap.login(email_address, password)
            imap.select("inbox")

            status, messages = imap.search(None, 'UNSEEN')
            email_ids = messages[0].split()
            email_ids = email_ids[-10:] 

            email_list = []

            for email_id in email_ids:
                res, msg_data = imap.fetch(email_id, "(RFC822)")
                for response_part in msg_data:
                    if isinstance(response_part, tuple):
                        msg = email.message_from_bytes(response_part[1])
                        subject, encoding = decode_header(msg["Subject"])[0]

                        if isinstance(subject, bytes):
                            subject = subject.decode(encoding if encoding else "utf-8")

                        from_ = msg.get("From", 'Không xác định') 
                        date = msg.get("Date", 'Không có ngày') 

                        email_list.append({
                            'subject': subject,
                            'from': from_,
                            'date': date,
                            'body': self.get_email_body(msg)
                        })

            imap.logout()
            return email_list
        except Exception as e:
            print(f"Không thể lấy email: {e}")
        return []


    def get_email_body(self, msg):
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))

                if content_type == "text/plain" and "attachment" not in content_disposition:
                    return part.get_payload(decode=True).decode()
        else:
            return msg.get_payload(decode=True).decode()
        return None

    # Xử lý email, kiểm tra xem có deadline nào bị trễ hạn không
    def process_email_data(self, emails):
        missed_deadlines = []
        current_date = datetime.now()

        for email_data in emails:
            body = email_data.get('body', None)
            deadline = self.extract_deadline_from_body(body)
            class_info = self.extract_class_from_body(body)  

            print(f"Email Body: {body}, Deadline: {deadline}, Current Date: {current_date}, Class: {class_info}")  
            if deadline is not None: 
                if current_date > deadline:
                    missed_deadlines.append({
                        **email_data,
                        'class_info': class_info  
                    })

        return missed_deadlines


    def extract_deadline_from_body(self, body):
        if not isinstance(body, str): 
            return None

        match = re.search(r"Deadline: (\d{2}/\d{2}/\d{4})", body)
        if match:
            deadline_str = match.group(1)
            return datetime.strptime(deadline_str, "%d/%m/%Y")
        return None

    def extract_class_from_body(self, body):
        if not isinstance(body, str): 
            return "Không xác định"

        match = re.search(r"Lớp: (.+)", body)
        if match:
            return match.group(1).strip()  
        return "Không xác định"  


    def send_missed_deadline_report(self, missed_deadlines):
        if missed_deadlines:
            report_body = "Báo cáo các deadline bị bỏ lỡ:\n\n"
            for email_data in missed_deadlines:
                report_body += (
                    f"- Từ: {email_data.get('from', 'Không xác định')}\n"
                    f"  Tiêu đề: {email_data.get('subject', 'Không có tiêu đề')}\n"
                    f"  Ngày: {email_data.get('date', 'Không xác định')}\n"
                    f"  Lớp: {email_data.get('class_info', 'Không xác định')}\n\n"
                )

            # Gửi email báo cáo
            self.send_email("quanly@example.com", "Báo cáo deadline bị bỏ lỡ", report_body)
            messagebox.showinfo("Thông báo", "Đã gửi báo cáo các deadline bị bỏ lỡ!")


    def send_email(self, to_email, subject, body):
            from_email = "leduyquan2574@gmail.com"  
            from_password = "lajewrwnkozpmulx"         

            to_email = str(to_email)
            
            msg = MIMEMultipart()
            msg['From'] = from_email
            msg['To'] = to_email
            msg['Subject'] = subject

            msg.attach(MIMEText(body, 'plain'))

            try:
                with smtplib.SMTP('smtp.gmail.com', 587) as server:  # SMTP server
                    server.starttls()
                    server.login(from_email, from_password)
                    server.send_message(msg)
                print(f"Email sent to {to_email}")
            except Exception as e:
                print(f"Failed to send email to {to_email}: {e}")

    def open_email_info_frame(self, emails):
        info_window = tk.Toplevel()
        info_window.title("Thông tin Email Gần Nhất")

        tree = ttk.Treeview(info_window, columns=("Sender", "Subject", "Date", "Body", "Class"), show='headings')
        tree.heading("Sender", text="Người Gửi")
        tree.heading("Subject", text="Tiêu Đề")
        tree.heading("Date", text="Ngày")
        tree.heading("Body", text="Nội Dung")
        tree.heading("Class", text="Lớp")

        tree.column("Sender", anchor="w", width=200)
        tree.column("Subject", anchor="w", width=300)
        tree.column("Date", anchor="w", width=100)
        tree.column("Body", anchor="w", width=300)
        tree.column("Class", anchor="w", width=100)

        tree.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        for email_data in emails:
            class_info = self.extract_class_from_body(email_data.get('body', ""))  
            tree.insert("", tk.END, values=(
                email_data.get('from', "Không xác định"),
                email_data.get('subject', "Không có tiêu đề"),
                email_data.get('date', "Không xác định"),  
                email_data.get('body', "Không có nội dung"),  
                class_info
            ))

        scrollbar = ttk.Scrollbar(info_window, orient="vertical", command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)


    def open_missed_deadline_frame(self, missed_deadlines):
        missed_window = tk.Toplevel()
        missed_window.title("Email Trễ Hạn")

        tree = ttk.Treeview(missed_window, columns=("Sender", "Subject", "Date", "Body", "Class"), show='headings')
        tree.heading("Sender", text="Người Gửi")
        tree.heading("Subject", text="Tiêu Đề")
        tree.heading("Date", text="Ngày")
        tree.heading("Body", text="Nội Dung")
        tree.heading("Class", text="Lớp")

        tree.column("Sender", anchor="w", width=200)
        tree.column("Subject", anchor="w", width=300)
        tree.column("Date", anchor="w", width=100)
        tree.column("Body", anchor="w", width=300)
        tree.column("Class", anchor="w", width=100)

        tree.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        for email_data in missed_deadlines:
            tree.insert("", tk.END, values=(
                email_data.get('from', 'Không xác định'), 
                email_data.get('subject', 'Không có tiêu đề'), 
                email_data.get('date', 'Không xác định'), 
                email_data.get('body', 'Không có nội dung'), 
                email_data.get('class_info', 'Không xác định')
            ))

        scrollbar = ttk.Scrollbar(missed_window, orient="vertical", command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)


# Assignment 3 
# 
# 

        self.staff_accounts = {
            "staff1": ("leduyquan2574@gmail.com", "lajewrwnkozpmulx"),
        }
    def check_staff_emails(self):
        for name, (email_address, password) in self.staff_accounts.items():
            imap = imaplib.IMAP4_SSL("imap.gmail.com")
            try:
                imap.login(email_address, password)

                # Chọn hộp thư 'inbox'
                status, _ = imap.select("inbox")
                if status != "OK":
                    print(f"Không thể chọn hộp thư 'inbox' cho {name} ({email_address})")
                    continue
                
                print(f"Đăng nhập thành công vào tài khoản của nhân viên {name}")

                # Tìm tất cả email trong thư mục đã chọn
                status, messages = imap.search(None, 'UNSEEN')
                if status != "OK":
                    print(f"Không thể tìm kiếm email cho {name} ({email_address})")
                    continue

                email_ids = messages[0].split()[-10:]  # Lấy 10 email mới nhất

                for email_id in email_ids:
                    res, msg_data = imap.fetch(email_id, "(RFC822)")
                    for response_part in msg_data:
                        if isinstance(response_part, tuple):
                            msg = email.message_from_bytes(response_part[1])
                            subject, encoding = decode_header(msg["Subject"])[0]

                            if isinstance(subject, bytes):
                                subject = subject.decode(encoding if encoding else "utf-8")

                            from_ = msg.get("From")
                            sender_email = email.utils.parseaddr(from_)[1]

                            # Kiểm tra xem email có phải là của sinh viên
                            student_email = self.extract_student_email(sender_email)
                            if student_email:
                                body = self.get_email_body(msg)

                                # Kiểm tra thời gian gửi email
                                send_time = msg["Date"]
                                
                                # Thay thế 'GMT' bằng '+0000' và xóa '(UTC)' 
                                send_time = send_time.replace("GMT", "+0000").replace("(UTC)", "").strip()
                                send_time = datetime.strptime(send_time, '%a, %d %b %Y %H:%M:%S %z')

                                # Nếu đã 24 giờ không nhận được phản hồi, gửi nhắc nhở
                                if datetime.now(tz=send_time.tzinfo) - send_time > timedelta(minutes=1):
                                    self.send_reminder_to_management(email_address, msg)

                imap.logout()
            except Exception as e:
                print(f"Không thể kiểm tra email cho {name} ({email_address}): {e}")
            finally:
                try:
                    imap.logout()
                except:
                    pass




    def extract_student_email(self, sender):
        # Kiểm tra xem email có phải là của sinh viên không
        if "@gmail.com" in sender:
            return sender.strip()  # Trả về địa chỉ email sinh viên
        return None  # Nếu không phải email sinh viên


    def send_reminder_to_management(self, staff_email, msg):
        # Lấy phần nội dung chính của email
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))

                if "attachment" not in content_disposition and content_type == "text/plain":
                    payload = part.get_payload(decode=True)
                    break
            else:
                payload = None  # Nếu không có phần text/plain
        else:
            payload = msg.get_payload(decode=True)

        if payload is not None:
            try:
                body = payload.decode()
            except UnicodeDecodeError:
                body = "Không thể giải mã nội dung email."
        else:
            body = "Không có nội dung email hoặc không thể giải mã."

        # Gửi nhắc nhở đến quản lý
        print(f"Gửi nhắc nhở đến quản lý cho nhân viên {staff_email}")
        
        reminder_body = f"Nhắc nhở: Nhân viên {staff_email} chưa trả lời email từ sinh viên.\n\nNội dung email:\n{body}"
        self.send_email("quanly@example.com", "Nhắc nhở về email chưa được trả lời", reminder_body)
        
        # Gửi nhắc nhở cho nhân viên
        staff_reminder_body = f"Chào bạn,\n\nBạn có một email mới từ sinh viên chưa được trả lời. Vui lòng kiểm tra hộp thư của bạn để xem chi tiết.\n\nNội dung email:\n{body}"
        self.send_email(staff_email, "Nhắc nhở: Kiểm tra email mới", staff_reminder_body)


    def process_new_emails(self):
        try:
            email_list = self.fetch_emails()
            processed_count = 0
            
            for email_data in email_list:
                # Kiểm tra nội dung email và trả lời tự động nếu có thể
                if self.process_single_email(email_data):
                    processed_count += 1
            
            messagebox.showinfo("Thông báo", f"Đã xử lý {processed_count} email")

        except Exception as e:
            messagebox.showerror("Lỗi", f"Có lỗi xảy ra: {str(e)}")

    def fetch_emails(self):
        imap = imaplib.IMAP4_SSL("imap.gmail.com")
        email_address = "leduyquan2574@gmail.com"
        password = "lajewrwnkozpmulx"

        try:
            imap.login(email_address, password)
            imap.select("inbox")

            status, messages = imap.search(None, 'UNSEEN')
            email_ids = messages[0].split()
            email_ids = email_ids[-10:]  
            email_list = []

            for email_id in email_ids:
                res, msg_data = imap.fetch(email_id, "(RFC822)")
                for response_part in msg_data:
                    if isinstance(response_part, tuple):
                        msg = email.message_from_bytes(response_part[1])
                        subject, encoding = decode_header(msg["Subject"])[0]

                        if isinstance(subject, bytes):
                            subject = subject.decode(encoding if encoding else "utf-8")

                        from_ = msg.get("From")
                        sender_email = email.utils.parseaddr(from_)[1]
                        body = self.get_email_body(msg)

                        email_list.append({
                            'subject': subject,
                            'from': from_,
                            'sender_email': sender_email,
                            'body': body
                        })

            imap.logout()
            return email_list
        except Exception as e:
            print(f"Không thể lấy email: {e}")
            return []

    def process_single_email(self, email_data):
        # Kiểm tra và chuyển đổi body thành chuỗi nếu cần
        body = email_data.get('body', '')
        if body is not None:
            body = str(body).lower().strip()
        else:
            body = ''
        
        sender_email = email_data['sender_email']  # Email người gửi câu hỏi
        response = ""

        # Truy vấn Q&A từ cơ sở dữ liệu
        query = "SELECT question, answer FROM qa"
        rows = self.db.connection.execute(query).fetchall()

        # Kiểm tra câu hỏi trong cơ sở dữ liệu
        for keyword, answer in rows:
            if keyword.lower() in body: 
                response += answer + "\n"

        if response:
            # Nếu tìm thấy câu trả lời, gửi email trả lời cho người gửi
            self.send_response_email(sender_email, response)
            return True
        else:
            # Nếu không tìm thấy, gửi email cho người phụ trách
            responsible_email = "phutrach@example.com"  # Thay đổi địa chỉ email người phụ trách
            subject = f"Câu hỏi từ {sender_email} chưa có trong cơ sở dữ liệu Q&A"
            content = f"Kính gửi người phụ trách,\n\nCâu hỏi từ {sender_email} chưa được tìm thấy trong cơ sở dữ liệu Q&A.\nVui lòng xem xét và phản hồi.\n\nNội dung câu hỏi:\n{body}"
            
            self.send_email(responsible_email, subject, content)
            return False




    def get_email_body(self, msg):
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":  
                    return part.get_payload(decode=True).decode("utf-8")
        else:
            return msg.get_payload(decode=True).decode("utf-8")


    def send_response_email(self, recipient, response):
        sender_email = "leduyquan2574@gmail.com"
        password = "lajewrwnkozpmulx"
        subject = "Phản hồi từ hệ thống"

        msg = MIMEText(response)
        msg['Subject'] = subject
        msg['From'] = sender_email
        msg['To'] = recipient

        server = smtplib.SMTP_SSL("smtp.gmail.com", 465)
        server.login(sender_email, password)
        server.sendmail(sender_email, recipient, msg.as_string())
        server.quit()


    def open_input_window(self):
        input_window = tk.Toplevel(self.classification_frame)
        input_window.title("Thêm Q&A / Chat Box")

        # Set window size and centering
        input_window_width = 600
        input_window_height = 400
        screen_width = input_window.winfo_screenwidth()
        screen_height = input_window.winfo_screenheight()
        x = (screen_width // 2) - (input_window_width // 2)
        y = (screen_height // 2) - (input_window_height // 2)
        input_window.geometry(f"{input_window_width}x{input_window_height}+{x}+{y}")

        # Create a frame for the Treeview
        treeview_frame = tk.Frame(input_window)
        treeview_frame.pack(padx=10, pady=10, fill="both", expand=True)

        # Treeview for displaying Q&A
        treeview = ttk.Treeview(treeview_frame, columns=("Câu hỏi", "Câu trả lời"), show='headings')
        treeview.heading("Câu hỏi", text="Câu hỏi")
        treeview.heading("Câu trả lời", text="Câu trả lời")
        treeview.pack(fill="both", expand=True)

        # Cập nhật Treeview từ database
        self.update_qa_display(treeview)

        # Create a frame for the input fields and buttons
        input_frame = tk.Frame(input_window)
        input_frame.pack(padx=10, pady=10, fill="x")

        tk.Label(input_frame, text="Câu hỏi:", font=("Helvetica", 12)).grid(row=0, column=0, sticky="w", padx=5, pady=5)
        question_entry = tk.Entry(input_frame, width=50, font=("Helvetica", 12))
        question_entry.grid(row=0, column=1, padx=5, pady=5)

        tk.Label(input_frame, text="Câu trả lời:", font=("Helvetica", 12)).grid(row=1, column=0, sticky="w", padx=5, pady=5)
        answer_entry = tk.Entry(input_frame, width=50, font=("Helvetica", 12))
        answer_entry.grid(row=1, column=1, padx=5, pady=5)

        # Create a frame for the buttons
        button_frame = tk.Frame(input_window)
        button_frame.pack(padx=10, pady=10)

        # Nút Lưu Q&A
        save_button = tk.Button(button_frame, text="Lưu Q&A", width=20, bg="lightgrey", fg="black", font=("Helvetica", 12),
                                command=lambda: self.save_qa(question_entry.get(), answer_entry.get(), treeview))
        save_button.grid(row=0, column=0, padx=5, pady=5)

        # Nút Kiểm tra Email
        check_email_button = tk.Button(button_frame, text="Kiểm tra Email", width=20, bg="lightblue", fg="black", font=("Helvetica", 12),
                                        command=self.process_new_emails)
        check_email_button.grid(row=0, column=1, padx=5, pady=5)

        # Nút Xóa Q&A
        delete_button = tk.Button(button_frame, text="Xóa Q&A", width=20, bg="red", fg="white", font=("Helvetica", 12),
                                command=lambda: self.delete_selected_qa(treeview))
        delete_button.grid(row=0, column=2, padx=5, pady=5)

        input_window.grid_columnconfigure(0, weight=1)
        input_window.grid_columnconfigure(1, weight=1)
        input_window.grid_rowconfigure(0, weight=1)

        # Ensure the input window remains on top
        input_window.transient(self.classification_frame)
        input_window.grab_set()
        self.classification_frame.wait_window(input_window)


    def delete_selected_qa(self, treeview):
        # Lấy dòng đang chọn
        selected_item = treeview.selection()
        if not selected_item:
            tk.messagebox.showwarning("Cảnh báo", "Vui lòng chọn dòng để xóa.")
            return

        # Lấy thông tin câu hỏi từ dòng đã chọn
        item_values = treeview.item(selected_item[0])["values"]
        question = item_values[0]

        # Xóa câu hỏi từ cơ sở dữ liệu
        query = "DELETE FROM qa WHERE question = ?"
        self.db.connection.execute(query, (question,))
        self.db.connection.commit()

        # Cập nhật lại Treeview sau khi xóa
        treeview.delete(selected_item[0])
        tk.messagebox.showinfo("Thông báo", f"Đã xóa Q&A với câu hỏi: {question}")





    def save_qa(self, question, answer, treeview):
        if question and answer:
            with self.db.connection:
                self.db.connection.execute("INSERT INTO qa (question, answer) VALUES (?, ?)", (question, answer))

            messagebox.showinfo("Thông báo", "Đã lưu Q&A thành công!")
        else:
            messagebox.showwarning("Cảnh báo", "Vui lòng nhập cả câu hỏi và câu trả lời.")


    def update_qa_display(self, treeview):
        for row in treeview.get_children():
            treeview.delete(row)

        query = "SELECT question, answer FROM qa"
        rows = self.db.connection.execute(query).fetchall()

        for row in rows:
            treeview.insert("", "end", values=row)


    def sort_classification_data(self):
        # Lấy dữ liệu từ cây hiển thị
        students = []
        for row in self.classification_tree.get_children():
            values = self.classification_tree.item(row)["values"]
            students.append(values)

        # Định nghĩa thứ tự cho các thể loại
        grade_order = {"A+": 1, "A": 2, "B+": 3, "B": 4, "C+": 5, "C": 6, "D": 7}

        # Sắp xếp theo thể loại và họ tên
        sorted_students = sorted(students, key=lambda x: (grade_order.get(x[-1], 8), x[0]))  # x[-1] là thể loại, x[0] là họ tên

        # Cập nhật cây hiển thị với dữ liệu đã sắp xếp
        self.classification_tree.delete(*self.classification_tree.get_children())  # Xóa dữ liệu cũ
        for student in sorted_students:
            self.classification_tree.insert("", "end", values=student)

    def search_students(self):
        search_term = self.search_entry.get().strip().lower()  # Lấy từ khóa tìm kiếm và chuyển thành chữ thường
        if not search_term:
            messagebox.showwarning("Cảnh báo", "Vui lòng nhập từ khóa tìm kiếm.")
            return

        # Lấy dữ liệu từ cây hiển thị
        students = []
        for row in self.classification_tree.get_children():
            values = self.classification_tree.item(row)["values"]
            students.append(values)

        # Lọc danh sách sinh viên theo từ khóa tìm kiếm
        filtered_students = [
            student for student in students 
            if search_term in student[0].lower() or search_term in str(student[1])  # Chuyển MSSV thành chuỗi
        ]

        # Cập nhật cây hiển thị với dữ liệu đã lọc
        self.classification_tree.delete(*self.classification_tree.get_children())  # Xóa dữ liệu cũ
        for student in filtered_students:
            self.classification_tree.insert("", "end", values=student)

        if not filtered_students:
            messagebox.showinfo("Kết quả tìm kiếm", "Không tìm thấy sinh viên nào.")

    def filter_classification_data(self):
        selected_grade = self.filter_combobox.get()  # Lấy phân loại đã chọn

        # Lấy dữ liệu từ cây hiển thị
        students = []
        for row in self.classification_tree.get_children():
            values = self.classification_tree.item(row)["values"]
            students.append(values)

        # Lọc dữ liệu theo phân loại
        if selected_grade == "Tất cả":
            filtered_students = students  # Hiển thị tất cả
        else:
            filtered_students = [student for student in students if student[-1] == selected_grade]  # Lọc theo phân loại

        # Cập nhật cây hiển thị với dữ liệu đã lọc
        self.classification_tree.delete(*self.classification_tree.get_children())  # Xóa dữ liệu cũ
        for student in filtered_students:
            self.classification_tree.insert("", "end", values=student)

        if not filtered_students:
            messagebox.showinfo("Kết quả lọc", "Không tìm thấy sinh viên nào với phân loại đã chọn.")

    def close(self):
        self.db.close()



# Source gmailnguye

#         Kính gửi các bạn sinh viên,

# Chúng tôi xin thông báo về bài tập mới dành cho lớp:

# Lớp: Lớp 15A1
# Tiêu đề bài tập: **Bài tập về lập trình Python
# Deadline: 20/10/2024 (trễ hạn)

# Rất tiếc vì đã không nhận được bài tập từ các bạn đúng hạn. Các bạn hãy nhanh chóng hoàn thành và gửi lại bài tập cho tôi.

# Nếu có bất kỳ câu hỏi nào, hãy liên hệ với giáo viên chủ nhiệm.

# Cảm ơn các bạn đã chú ý!

# Trân trọng,  
# Giáo viên bộ môn

