#!/usr/bin/env python3
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import sys
import os

def modify_schedule(input_path, output_path):
    # 黄色填充
    yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')

    # 需要替换成日班的关键词列表（添加结，凝血改为凝）
    daily_words = {'血', '急', '凝', '服', '接', '细', '肿', '小', '白', '骨', '生', '免', 'P日', '抽', '早', '精', '科', '结', '公'}

    # 完全匹配替换表
    exact_match_replace = {
        '前1': '中/前',
        '前2': '中/前',
        '急帮': '日/帮',
        '急帮休': '日/帮',
        '后1': '上午班/休',
        '后2': '上午班/休',
        '备': '休',
        '休': '休',
        '产假': '产假',
    }

    for word in daily_words:
        exact_match_replace[word] = '日班'

    # 使用openpyxl直接处理，保留原格式
    wb = load_workbook(input_path)

    modified_count = 0
    highlight_count = 0
    triangle_changed = 0

    # 我们需要跟踪哪些单元格已经被替换过（包括三角形替换）
    replaced_cells = set()

    # 先处理三角形替换：前1/前2、后1/后2、后1/中、后2/中后面的三角形
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        print(f"第一步：处理三角形替换 - {sheet_name}")
        
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None:
                    continue
                cell_text = str(cell.value).strip()
                # 判断当前单元格是否需要处理其右边的三角形
                if cell_text in ['前1', '前2', '后1', '后2', '后1/中', '后2/中']:
                    # 处理当前单元格右边的单元格
                    next_cell = ws.cell(row=cell.row, column=cell.column + 1)
                    if next_cell.value is None:
                        continue
                    next_text = str(next_cell.value).strip()
                    if next_text == '△' or next_text == '▲':
                        if cell_text in ['前1', '前2']:
                            next_cell.value = '值休'
                            print(f"  三角形替换: {next_cell.coordinate} △ → 值休 (前面是{cell_text})")
                        else:
                            # 后1、后2、后1/中、后2/中 后面的三角形都替换为后夜
                            next_cell.value = '后夜'
                            print(f"  三角形替换: {next_cell.coordinate} △ → 后夜 (前面是{cell_text})")
                        # 记录这个单元格已经被替换
                        replaced_cells.add((sheet_name, next_cell.row, next_cell.column))
                        triangle_changed += 1
                        modified_count += 1

    print(f"\n三角形替换完成，共 {triangle_changed} 处")

    # 处理每个单元格的其他替换规则
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        print(f"\n第二步：处理其他替换规则 - {sheet_name}")
        
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None:
                    continue
                
                cell_text = str(cell.value).strip()
                original = cell_text
                was_replaced = False
                
                # 检查是否已经被三角形替换过
                if (sheet_name, cell.row, cell.column) in replaced_cells:
                    was_replaced = True
                
                # 如果已经替换过，跳过替换逻辑，直接不标黄
                if was_replaced:
                    continue
                
                # 规则1: 包含 /中 → 整个替换为日班
                if '/中' in cell_text:
                    cell.value = '日班'
                    print(f"  {cell.coordinate}: {original} → 日班 (包含/中)")
                    modified_count += 1
                    was_replaced = True
                    continue
                
                # 规则2: 以中开头或以中结尾 → 整个替换为日班
                if cell_text.startswith('中') or cell_text.endswith('中'):
                    cell.value = '日班'
                    print(f"  {cell.coordinate}: {original} → 日班 (*中或中*)")
                    modified_count += 1
                    was_replaced = True
                    continue
                
                # 规则3: 正好两个字 且 以休结尾 → 上午班/休
                if len(cell_text) == 2 and cell_text.endswith('休'):
                    cell.value = '上午班/休'
                    print(f"  {cell.coordinate}: {original} → 上午班/休 (两字，*休结尾)")
                    modified_count += 1
                    was_replaced = True
                    continue
                
                # 规则4: 正好两个字 且 以休开头 → 休/下午班
                if len(cell_text) == 2 and cell_text.startswith('休'):
                    cell.value = '休/下午班'
                    print(f"  {cell.coordinate}: {original} → 休/下午班 (两字，休开头*)")
                    modified_count += 1
                    was_replaced = True
                    continue
                
                # 规则5: 完全匹配替换
                if cell_text in exact_match_replace:
                    cell.value = exact_match_replace[cell_text]
                    print(f"  {cell.coordinate}: {original} → {cell.value} (完全匹配)")
                    modified_count += 1
                    was_replaced = True
                    continue
                
                # 规则6: 包含两个或更多日班关键词 → 整个替换为日班
                count = 0
                for word in daily_words:
                    if word in cell_text:
                        count += 1
                if count >= 2:
                    cell.value = '日班'
                    print(f"  {cell.coordinate}: {original} → 日班 (包含{count}个关键词组合)")
                    modified_count += 1
                    was_replaced = True
                    continue
                
                # 没替换的单元格标黄
                # 跳过标题行（第一行是标题，跳过不标黄）
                # 只有从未被替换的才标黄
                if not was_replaced and cell.row > 1 and cell_text.strip() != '':
                    cell.fill = yellow_fill
                    highlight_count += 1

    # 保存修改后的文件
    wb.save(output_path)
    print(f"\n✅ 修改完成:")
    print(f"  替换单元格数: {modified_count}")
    print(f"  其中三角形替换: {triangle_changed}")
    print(f"  标黄单元格数: {highlight_count}")
    print(f"  保存到: {output_path}")
    
    return modified_count, highlight_count

def main():
    if len(sys.argv) == 3:
        input_path = sys.argv[1]
        output_path = sys.argv[2]
        modify_schedule(input_path, output_path)
    elif len(sys.argv) == 2:
        input_path = sys.argv[1]
        # 默认输出文件名在输入文件名基础上加 _modified
        name, ext = os.path.splitext(input_path)
        output_path = f"{name}_modified{ext}"
        modify_schedule(input_path, output_path)
    else:
        print("排班自动替换工具")
        print("用法:")
        print("  python run_modify_schedule.py <输入文件.xlsx>")
        print("  python run_modify_schedule.py <输入文件.xlsx> <输出文件.xlsx>")
        print("\n输出会自动在输入文件名后加上 _modified 后缀，不会覆盖原文件")

if __name__ == "__main__":
    main()
