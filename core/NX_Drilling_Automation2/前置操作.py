import pymysql
import json
import os
from decimal import Decimal
from datetime import datetime


class CustomJSONEncoder(json.JSONEncoder):
    """自定义 JSON 编码器，处理 Decimal、datetime 等特殊类型"""

    def default(self, obj):
        if isinstance(obj, Decimal):
            return float(obj)  # 或 str(obj) 保持精度
        if isinstance(obj, datetime):
            return obj.strftime('%Y-%m-%d %H:%M:%S')
        return super().default(obj)


def export_drill_table_to_json(
        host="localhost",
        port=3306,
        user="root",
        password="root",
        database="tool_database"
):
    """
    从 MySQL 数据库读取 drill_table 表数据并保存为 JSON 文件。
    """
    conn = None
    cursor = None
    try:
        # 连接 MySQL
        conn = pymysql.connect(
            host=host,
            port=port,
            user=user,
            password=password,
            database=database,
            charset="utf8mb4"
        )
        cursor = conn.cursor()

        table_list = ["knife_table","drill_table"]
        for table_name in table_list:
            output_json_path = rf"E:\{table_name}.json"
            # 查询 drill_table
            cursor.execute(f"SELECT * FROM {table_name}")
            columns = [desc[0] for desc in cursor.description]
            rows = cursor.fetchall()

            # 转换为 JSON 格式
            data_list = [dict(zip(columns, row)) for row in rows]

            # 保存为 JSON 文件（使用自定义编码器）
            output_dir = os.path.dirname(output_json_path)
            if output_dir:  # 防止路径为空
                os.makedirs(output_dir, exist_ok=True)

            with open(output_json_path, "w", encoding="utf-8") as f:
                json.dump(data_list, f, ensure_ascii=False, indent=4, cls=CustomJSONEncoder)

            print(f"✅ 数据已成功导出到: {output_json_path}")
            print(f"📊 共导出 {len(data_list)} 条记录")

    except pymysql.Error as e:
        print(f"❌ 数据库错误：{e}")
    except IOError as e:
        print(f"❌ 文件写入错误：{e}")
    except Exception as e:
        print(f"❌ 未知错误：{e}")
    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()


# ---------------- 主程序 ----------------
if __name__ == "__main__":
    export_drill_table_to_json()