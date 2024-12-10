import pandas as pd
import json


def read_sheet_data(sheet_name, file_path):
    """读取指定 sheet 的数据并生成相应层级的结构，跳过无意义的 'target' 数据."""
    try:
        sheet_df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

        fourth_fifth_level = []
        for idx, row in sheet_df.iterrows():
            fourth_level_value = row[4] if not pd.isna(row[4]) else None
            fifth_level_value = row[5] if not pd.isna(row[5]) else None

            # 跳过无意义的 'target' 数据
            if (fourth_level_value is None or fourth_level_value == "target") and \
                    (fifth_level_value is None or fifth_level_value == "target"):
                continue

            fourth_child = {
                "name": f"target{idx + 1}",
                "children": []
            }

            if fourth_level_value and fourth_level_value != "target":
                fourth_child["value"] = fourth_level_value

            if fifth_level_value and fifth_level_value != "target":
                fourth_child["children"].append({
                    "name": f"target{idx + 1}",
                    "value": fifth_level_value
                })

            fourth_fifth_level.append(fourth_child)
        return fourth_fifth_level
    except Exception as e:
        print(f"读取 {sheet_name} 出现错误: {e}")
        return []


def excel_to_json(file_path):
    df = pd.read_excel(file_path, sheet_name=0, header=None)

    # 初始化顶层字典结构
    json_data = {"name": "root", "children": []}

    for idx, row in df.iterrows():
        if idx == 0:
            continue  # 跳过第一行

        # 一级结构：A 列为 name，拼接 B 和 C 列为 value
        first_level_name = row[0] if not pd.isna(row[0]) else None
        if first_level_name is None or first_level_name == "target":
            continue  # 跳过无效的 name 值

        first_level = {"name": first_level_name}
        if not pd.isna(row[2]) or not pd.isna(row[1]):
            first_level_value = f"{row[2]}\n{row[1]}" if not pd.isna(row[2]) and not pd.isna(row[1]) else (
                row[2] if not pd.isna(row[2]) else row[1])
            if first_level_value != "target":
                first_level["value"] = first_level_value

        # 二级结构：D 列（第4列）
        targetIdx = 0
        second_level_name = f"HRS-RAP1 Target{targetIdx + 1}"
        second_level = {"name": second_level_name}
        if not pd.isna(row[3]) and row[3] != "target":
            second_level["value"] = row[3]

        # 三级结构
        third_level_names = ["HRS2", "PAC", "HRS3", "HRS4", "HRS9"]
        third_level_children = []
        for third_name in third_level_names:
            third_child = {"name": third_name, "children": []}
            # 根据三级名称读取第四层及第五层数据
            third_child["children"] = read_sheet_data(third_name, file_path)
            third_level_children.append(third_child)

        second_level["children"] = third_level_children
        first_level["children"] = [second_level]
        json_data["children"].append(first_level)

    # 写入 JSON 文件
    with open('output.json', 'w', encoding='utf-8') as f:
        json.dump(json_data, f, ensure_ascii=False, indent=4)

    return json_data


# 使用文件路径调用函数
file_path = 'Leadsership Target dashboard.xlsx'
output_json = excel_to_json(file_path)
print(json.dumps(output_json, ensure_ascii=False, indent=4))
