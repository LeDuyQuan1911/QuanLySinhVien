import sqlite3

class Database:
    def __init__(self):
        # Kết nối với cơ sở dữ liệu
        self.connection = sqlite3.connect("students.db")
        # Tạo bảng students và absences
        self.create_table()
        self.connection.execute("DROP TABLE IF EXISTS absences;")
        self.create_absences_table()
        self.create_qa_table()

    def create_table(self):
        # Tạo bảng students nếu chưa tồn tại
        with self.connection:
            self.connection.execute("""
            CREATE TABLE IF NOT EXISTS students (
                mssv TEXT PRIMARY KEY,
                ho_ten TEXT,
                lop TEXT,
                mon_hoc TEXT,
                so_buoi_vang INTEGER,
                ngay_nghi INTEGER
            )
            """)

    def create_absences_table(self):
        create_table_query = """
        CREATE TABLE IF NOT EXISTS absences (
            stt INTEGER NOT NULL,
            mssv INTEGER NOT NULL,
            ho_dem TEXT NOT NULL,
            ten TEXT NOT NULL,
            gioi_tinh TEXT NOT NULL,
            ngay_sinh TEXT,
            lop TEXT,  -- Thêm trường lớp
            mon TEXT,
            thu_1_status TEXT,
            thu_2_status TEXT,
            thu_3_status TEXT,
            thu_4_status TEXT,
            thu_5_status TEXT,
            thu_6_status TEXT,
            id INTEGER PRIMARY KEY AUTOINCREMENT
        );
        """

        with self.connection:
            self.connection.execute(create_table_query)

    def create_qa_table(self):
        with self.connection:
            self.connection.execute('''
                CREATE TABLE IF NOT EXISTS qa (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    question TEXT NOT NULL,
                    answer TEXT NOT NULL
                )
            ''')


    def add_absence(self, absence_data):
        # Thêm thông tin vắng mặt vào bảng absences
        query = """
        INSERT INTO absences (stt, mssv, ho_dem, ten, gioi_tinh, ngay_sinh,
            thu_1_status, thu_1_st, thu_1_ld, thu_2_status, thu_2_st, thu_2_ld,
            thu_3_status, thu_3_st, thu_3_ld, thu_4_status, thu_4_st, thu_4_ld,
            thu_5_status, thu_5_st, thu_5_ld, thu_6_status, thu_6_st, thu_6_ld,
            vang_co_phep, vang_khong_phep, tong_so_tiet, phan_tram_vang, id)
        VALUES (:stt, :mssv, :ho_dem, :ten, :gioi_tinh, :ngay_sinh,
            :thu_1_status, :thu_1_st, :thu_1_ld, :thu_2_status, :thu_2_st, :thu_2_ld,
            :thu_3_status, :thu_3_st, :thu_3_ld, :thu_4_status, :thu_4_st, :thu_4_ld,
            :thu_5_status, :thu_5_st, :thu_5_ld, :thu_6_status, :thu_6_st, :thu_6_ld,
            :vang_co_phep, :vang_khong_phep, :tong_so_tiet, :phan_tram_vang, :id)
        """
        with self.connection:
            self.connection.execute(query, absence_data)

    def get_students(self):
        # Lấy danh sách sinh viên
        with self.connection:
            return self.connection.execute("SELECT * FROM students").fetchall()

    def add_student(self, student_data):
        # Thêm sinh viên vào bảng students
        with self.connection:
            self.connection.execute("""
            INSERT OR REPLACE INTO students (mssv, ho_ten, lop, mon_hoc, so_buoi_vang, ngay_nghi)
            VALUES (:mssv, :ho_ten, :lop, :mon_hoc, :so_buoi_vang, :ngay_nghi)
            """, student_data)

    def delete_student(self, mssv):
        # Xóa sinh viên khỏi bảng students
        with self.connection:
            self.connection.execute("DELETE FROM students WHERE mssv = ?", (mssv,))

    def get_absence_info(self, mssv):
        # Lấy thông tin vắng mặt của sinh viên dựa trên mssv
        date_columns = [
            "thu_1_status",
            "thu_2_status",
            "thu_3_status",
            "thu_4_status",
            "thu_5_status",
            "thu_6_status",
        ]
        
        query = f"SELECT {', '.join(date_columns)} FROM absences WHERE mssv = ?"
        
        with self.connection:
            result = self.connection.execute(query, (mssv,)).fetchone()

        absence_info = {}
        if result:
            for i, status in enumerate(result):
                if status in ['K', 'P']:  # Chỉ lấy những ngày vắng mặt có phép hoặc không phép
                    absence_info[date_columns[i]] = status

        return absence_info

        

    def close(self):
        # Đóng kết nối cơ sở dữ liệu
        self.connection.close()
