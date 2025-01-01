# SQL自动执行与Excel报告生成器

这个Python项目用于自动执行SQL查询并将结果写入Excel文件。程序会在每天10点自动执行，并在写入Excel之前进行备份。

## 功能特点

- 每天10点自动执行指定的SQL查询
- 查询时间范围为昨天0点到今天0点
- 自动备份Excel文件
- 将查询结果写入Excel指定worksheet的倒数第二行
- 自动计算完成率
- 支持MySQL数据库（可配置为Oracle）
- 支持打包成Windows可执行文件，无需安装Python环境

## 开发环境配置

1. 克隆项目到本地
2. 安装依赖：
   ```bash
   pip install -r requirements.txt
   ```
3. 复制环境变量配置文件并修改：
   ```bash
   cp .env.example .env
   ```
4. 修改 `.env` 文件中的配置信息

## 打包为Windows可执行文件

1. 确保已安装所有依赖：
   ```bash
   pip install -r requirements.txt
   ```

2. 运行打包脚本：
   ```bash
   python build_exe.py
   ```

3. 打包完成后，可执行文件将在 `dist` 目录中生成：
   - `SQLExcelReporter.exe`: 主程序可执行文件

4. 部署步骤：
   - 将 `SQLExcelReporter.exe` 复制到目标机器
   - 在同目录下创建 `.env` 文件并配置数据库连接信息
   - 双击运行 `SQLExcelReporter.exe`

## 配置说明

在 `.env` 文件中配置以下参数：

- DB_HOST: 数据库主机地址
- DB_USER: 数据库用户名
- DB_PASSWORD: 数据库密码
- DB_NAME: 数据库名称
- DB_PORT: 数据库端口
- EXCEL_PATH: Excel文件路径
- SHEET_NAME: 工作表名称

## Excel文件格式

Excel文件应包含以下列：
- 开始时间
- 结束时间
- 记录数1
- 记录数2
- 完成率

## 运行程序

开发环境运行：
```bash
python main.py
```

生产环境运行：
- 双击 `SQLExcelReporter.exe`

## 注意事项

1. 确保Excel文件存在且具有正确的列名
2. 确保数据库连接信息正确
3. 程序会自动在写入Excel之前创建备份
4. 如需切换到Oracle数据库，需要修改数据库连接相关代码
5. 打包后的程序包含完整的运行环境，无需安装Python
6. 程序会在后台运行，可以在任务管理器中查看
