import os
import shutil
from datetime import datetime, timedelta
import schedule
import time
import pandas as pd
import oracledb
from dotenv import load_dotenv

# 加载环境变量
load_dotenv()

class DatabaseConfig:
    def __init__(self):
        self.host = os.getenv('DB_HOST', 'localhost')
        self.user = os.getenv('DB_USER', 'system')
        self.password = os.getenv('DB_PASSWORD', '')
        self.service_name = os.getenv('DB_NAME', 'orcl')
        self.port = int(os.getenv('DB_PORT', '1521'))
        self.encoding = os.getenv('DB_ENCODING', 'UTF-8')

class ExcelProcessor:
    def __init__(self, excel_path, sheet_name):
        self.excel_path = excel_path
        self.sheet_name = sheet_name

    def backup_excel(self):
        """备份Excel文件"""
        backup_path = f"{self.excel_path}.backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        shutil.copy2(self.excel_path, backup_path)
        return backup_path

    def update_excel(self, start_time, end_time, count1, count2):
        """更新Excel文件"""
        try:
            df = pd.read_excel(self.excel_path, sheet_name=self.sheet_name)
            
            # 计算完成率
            completion_rate = count1 / (count1 + count2) if (count1 + count2) > 0 else 0
            
            # 创建新行数据
            new_row = {
                '开始时间': start_time,
                '结束时间': end_time,
                '记录数1': count1,
                '记录数2': count2,
                '完成率': completion_rate
            }
            
            # 保存最后一行（合计行）
            last_row = df.iloc[-1:].copy()
            
            # 删除最后一行
            df = df.iloc[:-1]
            
            # 添加新行
            df.loc[len(df)] = new_row
            
            # 添加回合计行
            df = pd.concat([df, last_row], ignore_index=True)
            
            # 保存Excel
            df.to_excel(self.excel_path, sheet_name=self.sheet_name, index=False)
            
        except Exception as e:
            print(f"更新Excel文件时发生错误: {str(e)}")
            raise

class DatabaseQuery:
    def __init__(self, config):
        self.config = config

    def execute_queries(self, start_time, end_time):
        """执行SQL查询"""
        try:
            dsn = f"{self.config.host}:{self.config.port}/{self.config.service_name}"
            connection = oracledb.connect(
                user=self.config.user,
                password=self.config.password,
                dsn=dsn,
                encoding=self.config.encoding
            )

            with connection.cursor() as cursor:
                # 示例SQL查询1 - 待替换为实际SQL
                sql1 = f"""
                SELECT COUNT(*) 
                FROM t_user 
                WHERE create_time >= TO_TIMESTAMP(:start_time, 'YYYY-MM-DD HH24:MI:SS')
                AND create_time < TO_TIMESTAMP(:end_time, 'YYYY-MM-DD HH24:MI:SS')
                """
                cursor.execute(sql1, start_time=start_time, end_time=end_time)
                count1 = cursor.fetchone()[0]

                # 示例SQL查询2 - 待替换为实际SQL
                sql2 = f"""
                SELECT COUNT(*) 
                FROM t_user 
                WHERE create_time >= TO_TIMESTAMP(:start_time, 'YYYY-MM-DD HH24:MI:SS')
                AND create_time < TO_TIMESTAMP(:end_time, 'YYYY-MM-DD HH24:MI:SS')
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

def job():
    """定时任务主函数"""
    try:
        # 设置时间范围
        end_time = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        start_time = end_time - timedelta(days=1)
        
        # 初始化配置
        db_config = DatabaseConfig()
        excel_path = os.getenv('EXCEL_PATH', 'report.xlsx')
        sheet_name = os.getenv('SHEET_NAME', 'Sheet1')
        
        # 初始化处理器
        db_query = DatabaseQuery(db_config)
        excel_processor = ExcelProcessor(excel_path, sheet_name)
        
        # 备份Excel
        backup_path = excel_processor.backup_excel()
        print(f"Excel已备份到: {backup_path}")
        
        # 执行查询
        count1, count2 = db_query.execute_queries(
            start_time.strftime('%Y-%m-%d %H:%M:%S'),
            end_time.strftime('%Y-%m-%d %H:%M:%S')
        )
        
        # 更新Excel
        excel_processor.update_excel(start_time, end_time, count1, count2)
        print("Excel更新成功")
        
    except Exception as e:
        print(f"执行任务时发生错误: {str(e)}")

def main():
    # 设置每天10点执行任务
    # schedule.every().day.at("10:00").do(job)
    
    # print("定时任务已启动，将在每天10:00执行...")
    # while True:
        # schedule.run_pending()
        # time.sleep(60)
    job()

if __name__ == "__main__":
    main()
