import os
from datetime import datetime, timedelta
import xlsxwriter


def get_last_month():
    today = datetime.today()
    first = today.replace(day=1)
    last_month = first - timedelta(days=1)
    last_month_name = last_month.strftime("%B")
    return last_month_name


def read_file(path, data):
    try:
        with open(path, "r", encoding="UTF-8") as file:
            lines = file.readlines()
            for line in lines:
                row = [cell.strip() for cell in line.strip().split(",")]
                if any(row):
                    try:
                        start_datetime = datetime.strptime(f"{row[4]} {row[5]}", "%d/%m/%Y %H:%M:%S")
                        end_datetime = datetime.strptime(f"{row[6]} {row[7]}", "%d/%m/%Y %H:%M:%S")

                        duration = end_datetime - start_datetime
                        duration_seconds = duration.total_seconds()
                        duration_hours = int(duration_seconds // 3600)
                        duration_minutes = int((duration_seconds % 3600) // 60)

                        duration_str = f"{duration_hours}:{duration_minutes:02d}"

                        row.append(duration_str)
                    except ValueError:
                        row.append('Invalid date or time format')
                    data.append(row)
    except FileNotFoundError:
        print(f"Error: The file {path} does not exist.")
    return data


def excel_util(last_month, data):
    try:
        directory = f"D://optitecha//ataskaita//{last_month}"
        if not os.path.exists(directory):
            os.makedirs(directory)

        workbook = xlsxwriter.Workbook(f"{directory}//{last_month}.xlsx")
        worksheet = workbook.add_worksheet()

        headers = ['Company', 'Project', 'Task', 'Worker', 'Start Date', 'Start time', 'End date', 'End Time',
                   'Duration (h:m)']

        header_format = workbook.add_format({
            'bold': True,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#D9D9D9'
        })

        border_format = workbook.add_format({
            'border': 1
        })

        sum_label_format = workbook.add_format({
            'bold': True,
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
        })

        sum_value_format = workbook.add_format({
            'bold': True,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
        })

        worksheet.write_row(0, 0, headers, header_format)

        for row_idx, row in enumerate(data):
            padded_row = row + [''] * (len(headers) - len(row))
            for col_idx, cell_value in enumerate(padded_row):
                worksheet.write(row_idx + 1, col_idx, cell_value, border_format)

        total_duration_seconds = 0
        for row in data:
            if isinstance(row[-1], str) and ':' in row[-1]:
                hours, minutes = map(int, row[-1].split(':'))
                total_duration_seconds += (hours * 3600) + (minutes * 60)

        total_duration_hours = total_duration_seconds // 3600
        total_duration_minutes = (total_duration_seconds % 3600) // 60
        total_duration_str = f"{total_duration_hours}:{total_duration_minutes:02d}"

        summary_row = len(data) + 1
        worksheet.merge_range(summary_row, 0, summary_row, 7, "Sum:", sum_label_format)
        worksheet.write(summary_row, 8, total_duration_str, sum_value_format)

        for col_idx in range(len(headers)):
            max_length = len(headers[col_idx])
            for row in data:
                if col_idx < len(row):
                    max_length = max(max_length, len(str(row[col_idx])))
            worksheet.set_column(col_idx, col_idx, max_length + 2)

        workbook.close()
        print(f"Excel file saved at: {directory}//{last_month}.xlsx")
    except Exception as e:
        print(f"Error occurred while creating Excel file: {e}")
        import traceback
        traceback.print_exc()


def main():
    data = []
    last_month = get_last_month().lower()

    path = f"D://optitecha//ataskaita//{last_month}"
    path_txt = f"D://optitecha//ataskaita//{last_month}//{last_month}.txt"

    if not os.path.exists(path_txt):
        print(f"Warning: The file {path_txt} does not exist.")
        if not os.path.exists(path):
            os.makedirs(path)
            print(f"Created directory: {path}")
        print("No data to process. Empty Excel file will be created.")

    read_file(path_txt, data)
    excel_util(last_month, data)


if __name__ == "__main__":
    main()