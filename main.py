import os
import shutil
from datetime import datetime, timedelta
import schedule
import time
import oracledb
from dotenv import load_dotenv
import argparse
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import numbers
from copy import copy


class EnvConfig:
    def __init__(self, args=None):
        # 如果提供了命令行参数，优先使用命令行参数
        if args:
            self.host = args.host or os.getenv('DB_HOST', 'localhost')
            self.user = args.user or os.getenv('DB_USER', 'system')
            self.password = args.password or os.getenv('DB_PASSWORD', '')
            self.service_name = args.dbname or os.getenv('DB_NAME', 'orcl')
            self.port = args.port or int(os.getenv('DB_PORT', '1521'))
            self.encoding = args.encoding or os.getenv('DB_ENCODING', 'UTF-8')
            self.excel_path = args.excel_path or os.getenv('EXCEL_PATH', 'report.xlsx')
            self.sheet_name = args.sheet_name or os.getenv('SHEET_NAME', 'Sheet1')
            self.schedule_time = args.schedule_time or os.getenv('SCHEDULE_TIME', '09:00')
            # Excel column names
            self.col_start_time = args.col_start_time or os.getenv('COL_START_TIME', '开始时间')
            self.col_end_time = args.col_end_time or os.getenv('COL_END_TIME', '结束时间')
            self.col_guest_count = args.col_guest_count or os.getenv('COL_GUEST_COUNT', '入住旅客数')
            self.col_late_upload = args.col_late_upload or os.getenv('COL_LATE_UPLOAD', '15分上传不及时数')
            self.col_completion_rate = args.col_completion_rate or os.getenv('COL_COMPLETION_RATE', '完成率')
            # 时间参数
            self.start_time = args.start_time or os.getenv('START_TIME', '')
            self.end_time = args.end_time or os.getenv('END_TIME', '')
        else:
            self.host = os.getenv('DB_HOST', 'localhost')
            self.user = os.getenv('DB_USER', 'system')
            self.password = os.getenv('DB_PASSWORD', '')
            self.service_name = os.getenv('DB_NAME', 'orcl')
            self.port = int(os.getenv('DB_PORT', '1521'))
            self.encoding = os.getenv('DB_ENCODING', 'UTF-8')
            self.excel_path = os.getenv('EXCEL_PATH', 'report.xlsx')
            self.sheet_name = os.getenv('SHEET_NAME', 'Sheet1')
            self.schedule_time = os.getenv('SCHEDULE_TIME', '09:00')
            # Excel column names
            self.col_start_time = os.getenv('COL_START_TIME', '开始时间')
            self.col_end_time = os.getenv('COL_END_TIME', '结束时间')
            self.col_guest_count = os.getenv('COL_GUEST_COUNT', '入住旅客数')
            self.col_late_upload = os.getenv('COL_LATE_UPLOAD', '15分上传不及时数')
            self.col_completion_rate = os.getenv('COL_COMPLETION_RATE', '完成率')
            # 时间参数
            self.start_time = os.getenv('START_TIME', '')
            self.end_time = os.getenv('END_TIME', '')


class ExcelProcessor:
    def __init__(self, config):
        self.config = config
        self.excel_path = config.excel_path
        self.sheet_name = config.sheet_name

    def backup_excel(self):
        """备份Excel文件"""
        backup_path = f"{self.excel_path}.backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        shutil.copy2(self.excel_path, backup_path)
        return backup_path

    def update_excel(self, start_time, end_time, count1, count2):
        """更新Excel文件，保留原有格式和其他工作表"""
        try:
            # 使用openpyxl加载整个工作簿，保持数据和格式
            workbook = openpyxl.load_workbook(self.excel_path, data_only=False)
            sheet = workbook[self.sheet_name]

            # 找到最后一行（不包括合计行）
            last_row = sheet.max_row - 1  # 减1是为了排除合计行

            # 获取列标题和对应的列号
            headers = {}
            for col in range(1, sheet.max_column + 1):
                cell_value = sheet.cell(row=1, column=col).value
                if cell_value in [self.config.col_start_time, self.config.col_end_time,
                                self.config.col_guest_count, self.config.col_late_upload,
                                self.config.col_completion_rate]:
                    headers[cell_value] = col

            # 插入新行在倒数第二行（保持合计行在最后）
            sheet.insert_rows(last_row + 1)
            new_row_num = last_row + 1

            # 复制上一行的格式到新行
            for col in range(1, sheet.max_column + 1):
                source_cell = sheet.cell(row=last_row, column=col)
                target_cell = sheet.cell(row=new_row_num, column=col)
                if source_cell.has_style:
                    target_cell._style = copy(source_cell._style)
                    target_cell.number_format = source_cell.number_format

            # 写入新数据并设置适当的格式
            col_data = {
                self.config.col_start_time: (start_time, 'date'),
                self.config.col_end_time: (end_time, 'date'),
                self.config.col_guest_count: (count1, 'number'),
                self.config.col_late_upload: (count2, 'number')
            }

            for header, (value, data_type) in col_data.items():
                if header in headers:
                    cell = sheet.cell(row=new_row_num, column=headers[header])
                    # 确保时间值作为字符串写入，防止Excel自动合并
                    if data_type == 'date':
                        cell.value = str(value)
                        cell.number_format = '@'  # 设置为文本格式
                    else:
                        cell.value = value
                        if data_type == 'number':
                            cell.number_format = '0'

            # 更新完成率列的公式
            if self.config.col_completion_rate in headers:
                completion_col = headers[self.config.col_completion_rate]
                count1_col = headers[self.config.col_guest_count]
                count2_col = headers[self.config.col_late_upload]
                
                # 获取完成率的计算公式
                formula_cell = sheet.cell(row=last_row, column=completion_col)
                formula_value = str(formula_cell.value)
                if formula_value.startswith('='):
                    # 如果存在公式，复制并更新行号
                    new_formula = formula_value.replace(str(last_row), str(new_row_num))
                    sheet.cell(row=new_row_num, column=completion_col).value = new_formula
                else:
                    # 如果没有公式，创建新的公式
                    count1_letter = openpyxl.utils.get_column_letter(count1_col)
                    count2_letter = openpyxl.utils.get_column_letter(count2_col)
                    formula = f"={count1_letter}{new_row_num}/({count1_letter}{new_row_num}+{count2_letter}{new_row_num})"
                    sheet.cell(row=new_row_num, column=completion_col).value = formula
                    sheet.cell(row=new_row_num, column=completion_col).number_format = '0.00%'

            # 更新合计行的公式范围
            total_row = sheet.max_row
            for col in range(1, sheet.max_column + 1):
                cell = sheet.cell(row=total_row, column=col)
                cell_value = str(cell.value)
                if cell_value.startswith('='):
                    # 替换SUM公式范围
                    if 'SUM' in cell_value.upper():
                        start_ref = cell_value[cell_value.find('(') + 1:cell_value.find(':')]
                        # 获取当前列号
                        col_letter = get_column_letter(cell.column)
                        end_ref = f"{col_letter}{total_row-1}"
                        new_formula = f"=SUM({start_ref}:{end_ref})"
                    # 最后一列的完成率公式
                    elif col == completion_col:
                        count1_letter = openpyxl.utils.get_column_letter(col-2)
                        count2_letter = openpyxl.utils.get_column_letter(col-1)
                        new_formula = f"={count1_letter}{total_row}/({count1_letter}{total_row}+{count2_letter}{total_row})"
                        cell.value = new_formula
                        cell.number_format = '0.00%'
            # 保存工作簿
            workbook.save(self.excel_path)
            print(f"已更新Excel文件 {self.excel_path}")

        except Exception as e:
            print(f"更新Excel时发生错误: {str(e)}")
            raise


class DatabaseQuery:
    def __init__(self, config):
        self.config = config

    def execute_queries(self, start_time, end_time):
        """执行SQL查询"""
        try:
            # 转换时间格式
            try:
                # 尝试解析为YYYY-MM-DD HH:MM:SS格式
                start_dt = datetime.strptime(start_time, '%Y-%m-%d %H:%M:%S')
                end_dt = datetime.strptime(end_time, '%Y-%m-%d %H:%M:%S')
            except ValueError:
                try:
                    # 尝试解析为YYYY-MM-DD格式
                    start_dt = datetime.strptime(start_time, '%Y-%m-%d')
                    end_dt = datetime.strptime(end_time, '%Y-%m-%d')
                except ValueError:
                    # 尝试解析为YYYYMMDDHHMISS格式
                    start_dt = datetime.strptime(start_time, '%Y%m%d%H%M%S')
                    end_dt = datetime.strptime(end_time, '%Y%m%d%H%M%S')
            
            # 转换为数据库需要的格式
            start_time = start_dt.strftime('%Y%m%d%H%M%S')
            end_time = end_dt.strftime('%Y%m%d%H%M%S')
            
            dsn = f"{self.config.host}:{self.config.port}/{self.config.service_name}"
            connection = oracledb.connect(
                user=self.config.user,
                password=self.config.password,
                dsn=dsn
            )

            with connection.cursor() as cursor:
                # 示例SQL查询1 - 待替换为实际SQL
                # sql1 = """
                # SELECT COUNT(*) 
                # FROM t_user 
                # WHERE create_time >= TO_TIMESTAMP(:start_time, 'YYYY-MM-DD HH24:MI:SS')
                # AND create_time < TO_TIMESTAMP(:end_time, 'YYYY-MM-DD HH24:MI:SS')
                # """

                sql1 = """
                SELECT COUNT(*) 
                FROM v_gnlkxx_50 v
                WHERE v.rzsj >= TO_DATE(:start_time, 'YYYYMMDDHH24MISS')
                AND v.rzsj < TO_DATE(:end_time, 'YYYYMMDDHH24MISS')
                """
                cursor.execute(sql1, start_time=start_time, end_time=end_time)
                count1 = cursor.fetchone()[0]

                # 示例SQL查询2 - 待替换为实际SQL
                # sql2 = """
                # SELECT COUNT(*) 
                # FROM t_user 
                # WHERE create_time >= TO_TIMESTAMP(:start_time, 'YYYY-MM-DD HH24:MI:SS')
                # AND create_time < TO_TIMESTAMP(:end_time, 'YYYY-MM-DD HH24:MI:SS')
                # """
                sql2 = """
                SELECT COUNT(*) 
                FROM v_gnlkxx_50 v
                WHERE v.rzsj >= TO_DATE(:start_time, 'YYYYMMDDHH24MISS')
                AND v.rzsj < TO_DATE(:end_time, 'YYYYMMDDHH24MISS') 
                AND ((v.tfsj is null and (v.gxsj-v.rzsj)*24*60 >= 15) or (v.tfsj is not null and (v.gxsj-v.tfsj)*24*60 >= 15))
                """
                cursor.execute(sql2, start_time=start_time, end_time=end_time)
                count2 = cursor.fetchone()[0]

            return count1, count2

        except Exception as e:
            print(f"执行数据库查询时发生错误: {str(e)}")
            raise
        finally:
            if 'connection' in locals():
                connection.close()


def job(config):
    """定时任务主函数"""
    try:
        # 设置时间范围
        # 优先使用配置的时间参数，如果没有配置则使用默认时间范围
        if config.start_time and config.end_time:
            start_time = datetime.strptime(config.start_time, '%Y%m%d%H%M%S')
            end_time = datetime.strptime(config.end_time, '%Y%m%d%H%M%S')
        else:
            end_time = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
            start_time = end_time - timedelta(days=1)
        
        # 初始化处理器
        db_query = DatabaseQuery(config)
        excel_processor = ExcelProcessor(config)
        
        # 备份Excel
        backup_path = excel_processor.backup_excel()
        print(f"Excel已备份到: {backup_path}")
        
        # 执行查询
        count1, count2 = db_query.execute_queries(
            start_time.strftime('%Y%m%d%H%M%S'),
            end_time.strftime('%Y%m%d%H%M%S')
        )
        
        # 更新Excel
        excel_processor.update_excel(start_time.strftime('%Y%m%d%H%M%S'), end_time.strftime('%Y%m%d%H%M%S'), count1, count2)
        print("Excel更新成功")
        
    except Exception as e:
        print(f"执行任务时发生错误: {str(e)}")


def main():
    # 创建命令行参数解析器
    parser = argparse.ArgumentParser(description='SQL查询结果导出到Excel工具')
    
    # 添加数据库连接参数
    parser.add_argument('--host', help='数据库主机地址')
    parser.add_argument('--port', type=int, help='数据库端口')
    parser.add_argument('--user', help='数据库用户名')
    parser.add_argument('--password', help='数据库密码')
    parser.add_argument('--dbname', help='数据库名称')
    parser.add_argument('--encoding', help='数据库编码')
    
    # 添加Excel相关参数
    parser.add_argument('--excel-path', help='Excel文件路径')
    parser.add_argument('--sheet-name', help='Excel工作表名称')
    
    # 添加Excel列名参数
    parser.add_argument('--col-start-time', help='开始时间列名')
    parser.add_argument('--col-end-time', help='结束时间列名')
    parser.add_argument('--col-guest-count', help='入住旅客数列名')
    parser.add_argument('--col-late-upload', help='15分上传不及时数列名')
    parser.add_argument('--col-completion-rate', help='完成率列名')
    
    # 添加时间参数
    parser.add_argument('--start-time', help='开始时间，支持格式：YYYY-MM-DD HH:MM:SS 或 YYYY-MM-DD 或 YYYYMMDDHHMISS')
    parser.add_argument('--end-time', help='结束时间，支持格式：YYYY-MM-DD HH:MM:SS 或 YYYY-MM-DD 或 YYYYMMDDHHMISS')
    
    # 添加运行模式参数
    parser.add_argument('--run-now', action='store_true', help='立即执行一次，不等待定时任务')
    parser.add_argument('--config', help='指定配置文件路径')

    # 添加调度参数
    parser.add_argument('--schedule-time', help='定时任务时间，格式为 HH:MM')
    
    args = parser.parse_args()
    
    # 如果指定了配置文件，加载配置文件
    if args.config:
        load_dotenv(args.config)
    else:
        load_dotenv()

     
    # 创建配置对象
    config = EnvConfig(args)
    
    # 初始化Oracle thick mode
    oracledb.init_oracle_client()
    # 根据参数决定是立即执行还是运行定时任务
    if args.run_now:
        job(config)
    else:
        print("每天{}点执行任务".format(config.schedule_time))
        # 设置定时任务
        schedule.every().day.at(config.schedule_time).do(job, config)
        
        while True:
            schedule.run_pending()
            time.sleep(60)


if __name__ == "__main__":
    main()
