import os
import shutil
from datetime import datetime, timedelta
import schedule
import time
import pandas as pd
import oracledb
from dotenv import load_dotenv
import argparse


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
                '入住旅客数': count1,
                '15分上传不及时数': count2,
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
                AND ((v.tfsj is null and (v.gxsj-v.rzsj)*24*60 >= 15) or (v.tfsj is not null and (v.tfsj-v.rzsj)*24*60 >= 15))
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
        end_time = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        start_time = end_time - timedelta(days=1)
        
        # 初始化处理器
        db_query = DatabaseQuery(config)
        excel_processor = ExcelProcessor(config.excel_path, config.sheet_name)
        
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
