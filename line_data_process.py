import re
from datetime import datetime

from constant import COL_NAME_LIST, USELESS_COL_NAME_LIST, AGE_COL_NAME, NEW_AGE_COL_NAME


def modify_sample_source(line):
    """
    根据实验室编号改写标本来源
    """
    source = ""
    for line_kv in line:
        k = list(line_kv.keys())[0]
        v = line_kv[k]
        if k == "实验室编号":
            if "S" in v:
                source = "SARI病例"
            else:
                source = "ILI病例"
    for line_index, line_kv in enumerate(line):
        k = list(line_kv.keys())[0]
        if k == "标本来源":
            line[line_index][k] = source
    return line


def modify_flu_a(line):
    """
    改写甲型流感通用型
    """
    change_yang = False
    num_yin = 0
    for line_kv in line:
        k = list(line_kv.keys())[0]
        v = line_kv[k]
        if k in ["甲型流感H3", "甲型流感H1 2009"]:
            if v == "阳性":
                change_yang = 1
                break
            elif v == "阴性":
                num_yin += 1
    for line_index, line_kv in enumerate(line):
        k = list(line_kv.keys())[0]
        if k == "甲型流感通用型":
            if change_yang:
                line[line_index][k] = "阳性"
            if num_yin == 2:
                line[line_index][k] = "阴性"
    return line


def modify_flu_b(line):
    """
    改写乙型流感通用型
    """
    change_yang = False
    num_yin = 0
    for line_kv in line:
        k = list(line_kv.keys())[0]
        v = line_kv[k]
        if k in ["乙型流感BV", "乙型流感BY"]:
            if v == "阳性":
                change_yang = 1
                break
            elif v == "阴性":
                num_yin += 1
    for line_index, line_kv in enumerate(line):
        k = list(line_kv.keys())[0]
        if k == "乙型流感通用型":
            if change_yang:
                line[line_index][k] = "阳性"
            if num_yin == 2:
                line[line_index][k] = "阴性"
    return line


def modify_rsv(line):
    """
    改写呼吸道合胞病毒通用型
    """
    change_yang = False
    num_yin = 0
    for line_kv in line:
        k = list(line_kv.keys())[0]
        v = line_kv[k]
        if k in ["呼吸道合胞病毒A", "呼吸道合胞病毒B"]:
            if v == "阳性":
                change_yang = 1
                break
            elif v == "阴性":
                num_yin += 1
    for line_index, line_kv in enumerate(line):
        k = list(line_kv.keys())[0]
        if k == "呼吸道合胞病毒通用型":
            if change_yang:
                line[line_index][k] = "阳性"
            if num_yin == 2:
                line[line_index][k] = "阴性"
    return line


def modify_hpiv(line):
    """
    改写副流感病毒通用型
    """
    change_yang = False
    num_yin = 0
    for line_kv in line:
        k = list(line_kv.keys())[0]
        v = line_kv[k]
        if k in ["副流感病毒1", "副流感病毒2", "副流感病毒3", "副流感病毒4"]:
            if v == "阳性":
                change_yang = 1
                break
            elif v == "阴性":
                num_yin += 1
    for line_index, line_kv in enumerate(line):
        k = list(line_kv.keys())[0]
        if k == "副流感病毒通用型":
            if change_yang:
                line[line_index][k] = "阳性"
            if num_yin == 4:
                line[line_index][k] = "阴性"
    return line


def modify_coronavirus(line):
    """
    改写普通冠状病毒通用型
    """
    change_yang = False
    num_yin = 0
    for line_kv in line:
        k = list(line_kv.keys())[0]
        v = line_kv[k]
        if k in ["冠状病毒229E", "冠状病毒NL63", "冠状病毒OC43", "冠状病毒HKU1"]:
            if v == "阳性":
                change_yang = 1
                break
            elif v == "阴性":
                num_yin += 1
    for line_index, line_kv in enumerate(line):
        k = list(line_kv.keys())[0]
        if k == "普通冠状病毒通用型":
            if change_yang:
                line[line_index][k] = "阳性"
            if num_yin == 4:
                line[line_index][k] = "阴性"
    return line


def line_value_modify(line):
    """
    按照预设逻辑，调整行内数值
    """
    line = modify_sample_source(line)
    line = modify_flu_a(line)
    line = modify_flu_b(line)
    line = modify_rsv(line)
    line = modify_hpiv(line)
    line = modify_coronavirus(line)
    return line


def calibrate_number(value):
    """
    校准实验室编号，返回是否合规、校准后的值
    """
    value = str(value)
    if value is None or ("/" not in value and len(value) != 6):
        return False, value
    else:
        value = value.replace("s", "S")
        value = value.replace(" ", "")
        return True, value


def calibrate_carrier(value):
    """
    校准职业，返回是否合规、校准后的值
    """
    value = str(value)
    if value == "女":
        return False, value
    if value in ["干部职员", "干部职工"]:
        value = "干部职工"
    if value in ["职工", "职员"]:
        value = "职工"
    if value in ["待业", "家务", "家务及待业", "无业", "无"]:
        value = "无业"
    if value in ["不详", "未知", "None"]:
        value = "未知"
    if value in ["其他", "其它"]:
        value = "其他"
    if value in ["幼托儿童", "托幼儿童", ]:
        value = "幼托儿童"
    if value in ["医生", "护士", "医务人员", "医护人员"]:
        value = "医护人员"
    if value in ["离退人员", "退休", ]:
        value = "退休"
    return True, value


def calibrate_gender_and_hospital(value):
    """
    校准性别，返回是否合规、校准后的值
    """
    value = str(value)
    if value == "" or value == "None":
        return False, value
    return True, value


def parse_real_age(age_value):
    """
    解析原年龄字段，获得实际的岁数，单位为年
    :param age_value: 原年龄字段
    :return: 岁数
    """
    if isinstance(age_value, int) or isinstance(age_value, float):
        return age_value
    match = re.fullmatch(r'(?:(?P<year>\d+)岁)?(?:(?P<month>\d+)月)?(?:(?P<day>\d+)天)?', str(age_value))
    age = 0
    if match:
        y, m, d = match.group('year'), match.group('month'), match.group('day')
        age += int(y) if y else 0 + int(m) / 12 if m else 0 + (int(d) // 30) / 12 if d else 0
    return age


def calibrate_age(value):
    """
    校准年龄，返回是否合规、校准后的值
    """
    value = str(value)
    if value == "" or value == "None":
        return False, value
    if "岁" in value or "月" in value or "天" in value:
        value = str(value)
        value = value.replace(" ", "")
        value = value.replace("个", "")
        age = parse_real_age(value)
        if age >= 1:
            age = round(age, 1)
            return True, age
        return True, value
    else:
        value = int(value)
        if value > 150:
            return False, value
        return True, value


def calibrate_date(value):
    """
    校准日期，返回是否合规、校准后的值
    """
    if isinstance(value, datetime):
        return True, value
    else:
        value = str(value)
        if value == "" or value == "None":
            return False, value
        if "-" in value or "/" in value:
            if "-" in value:
                date = value.split("-")
            else:
                date = value.split("/")
            value = datetime(year=int(date[0]), month=int(date[1]), day=int(date[2]))
        return True, value


def calibrate_sample_category(value):
    """
    校准标本种类，返回是否合规、校准后的值
    """
    value = str(value)
    if value == "" or value == "None":
        return False, value
    if value == "痰液":
        value = "痰"
    return True, value


def calibrate_pathogen(value):
    """
    校准病原检测结果，返��是否合规、校准后的值
    """
    value = str(value)
    if value == "" or value == "None":
        value = "阴性"
    if value == "是":
        value = "阳性"
    assert value in ["未检测", "阴性", "阳性"]
    return True, value


def line_value_calibrate(col_name, value):
    """
    按照数值自身定义，校准行内数值
    """
    if col_name == "实验室编号":
        calibrate_result, value = calibrate_number(value)
    elif col_name == "职业":
        calibrate_result, value = calibrate_carrier(value)
    elif col_name == "性别" or col_name == "采集医院":
        calibrate_result, value = calibrate_gender_and_hospital(value)
    elif col_name == "年龄":
        calibrate_result, value = calibrate_age(value)
    elif col_name in ["发病日期", "采集日期", "送检日期"]:
        calibrate_result, value = calibrate_date(value)
    elif col_name == "标本种类":
        calibrate_result, value = calibrate_sample_category(value)
    # 来源不做校准，在modify模块，使用实验室编号的信息进行覆盖
    elif col_name == "标本来源":
        calibrate_result, value = True, ""
    # 病原检测结果
    else:
        calibrate_result, value = calibrate_pathogen(value)
    return calibrate_result, value


def write_line_data(write_row_index, ws_write, line):
    """
    写入工作簿
    :param write_row_index: 写的工作簿的行号
    :param ws_write: 工作簿
    :param line: 要写的行信息
    :return: 写的工作簿
    """
    # 新增一列年龄，用于将所有数都转为浮点数，此处写第一行的列名
    ws_write.cell(row=1, column=len(COL_NAME_LIST) + 1, value=NEW_AGE_COL_NAME)
    # 遍历所有kv对
    for line_kv in line:
        k = list(line_kv.keys())[0]
        # 无用的跳过
        if k in USELESS_COL_NAME_LIST:
            continue
        v = line_kv[k]
        # 写入工作簿
        ws_write.cell(row=write_row_index, column=COL_NAME_LIST.index(k) + 1, value=v)
        # 如果是日期，调整单元格格式
        if isinstance(v, datetime):
            ws_write.cell(row=write_row_index, column=COL_NAME_LIST.index(k) + 1).style = "date_style"
        # 在新增的年龄列中计算并写入
        if k == AGE_COL_NAME:
            age = parse_real_age(v)
            assert age != 0
            ws_write.cell(row=write_row_index, column=len(COL_NAME_LIST) + 1, value=round(age, 1))
    return ws_write
