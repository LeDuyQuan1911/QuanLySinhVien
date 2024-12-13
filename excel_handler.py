import pandas as pd

class ExcelHandler:
    def export_absence_report(self, students, file_name):
        df = pd.DataFrame(students, columns=["Lớp", "Môn học", "Họ tên", "MSSV", "Số buổi vắng", "Ngày vắng"])
        df.to_excel(file_name, index=False)

    def import_data(self, file_name):
        df = pd.read_excel(file_name)
        return df.to_dict('records')
