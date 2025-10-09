from pathlib import Path
from datetime import datetime, timedelta


def find_newest_zip_time(directory):
    """
    获取指定目录中最近一次创建的zip文件时间
    """
    newest_time = None

    for item in Path(directory).rglob('*'):
        print('item', item)
        if item.is_file():
            try:
                ctime = item.stat().st_ctime

                # 更新最新时间
                if newest_time is None or ctime > newest_time:
                    newest_time = ctime

            except (OSError, PermissionError):
                continue

    if newest_time:
        print(f"创建时间: {datetime.fromtimestamp(newest_time)}")
        # return datetime.fromtimestamp(newest_time)
    else:
        # return None
        print('创建时间为不存在')


if __name__ == '__main__':

    directory = "./软件"
    new_files = find_newest_zip_time(directory)