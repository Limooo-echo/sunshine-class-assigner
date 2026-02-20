import pandas as pd
import os
import sys


def analyze_class_balance(file_path):
    # 1. 读取文件
    if not os.path.exists(file_path):
        print("❌ 文件不存在")
        return

    try:
        df = pd.read_excel(file_path)
        print(f"✅ 成功读取文件，共 {len(df)} 条数据")
    except Exception as e:
        print(f"❌ 读取 Excel 失败: {e}")
        return

    # 2. 检查必要列名
    required_cols = ['班级', '性别']
    # 检查是否有'城乡'列 (兼容旧数据)
    has_rural = '城乡' in df.columns

    # 检查基础列是否存在
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        print(f"❌ 缺少必要列: {missing}")
        return

    print("🔄 正在进行统计分析...")

    # 3. 核心统计逻辑
    # 使用 groupby + apply 一次性计算所有指标
    def calc_stats(x):
        stats = {
            '男生': (x['性别'] == '男').sum(),
            '女生': (x['性别'] == '女').sum(),
            '总人数': len(x),
        }

        # 如果有总分，计算平均分
        if '总分' in x.columns:
            stats['平均分'] = round(x['总分'].mean(), 2)

        # 如果有城乡，计算城乡分布
        if has_rural:
            stats['城区'] = (x['城乡'] == '城区').sum()
            stats['乡下'] = (x['城乡'] == '乡下').sum()

        return pd.Series(stats)

    # 按班级分组计算
    result = df.groupby('班级').apply(calc_stats)

    # 4. 计算“全年级平均”行
    # 计算列的平均值
    avg_row = result.mean()

    # 将结果转换为 DataFrame 并添加平均行
    final_df = result.copy()
    final_df.loc['平均'] = avg_row

    # 5. 格式化数据（保留小数位）
    # 人数类指标保留1位小数 (为了看平均值的 .4 这种)，或者取整
    # 这里为了模仿截图，平均行保留1位小数，其他行取整

    cols_order = ['男生', '女生', '总人数', '平均分']
    if has_rural:
        cols_order += ['城区', '乡下']

    # 过滤掉不存在的列（防止没总分的情况）
    cols_order = [c for c in cols_order if c in final_df.columns]
    final_df = final_df[cols_order]

    # 6. 打印预览表格
    print("\n" + "=" * 50)
    print(" 📊 分班统计报告 ")
    print("=" * 50)
    print(final_df.round(1).to_string())  # 控制台打印保留1位小数
    print("=" * 50)

    # 7. 导出 Excel
    input_dir = os.path.dirname(file_path)
    base_name = os.path.splitext(os.path.basename(file_path))[0]
    output_path = os.path.join(input_dir, f"{base_name}_统计报告.xlsx")

    try:
        final_df.round(2).to_excel(output_path)
        print(f"\n✅ 统计完成！详细报告已生成:\n👉 {output_path}")
    except PermissionError:
        print("\n❌ 保存失败！请关闭正在打开的 Excel 文件。")


# ================= 主程序 =================
if __name__ == "__main__":
    while True:
        print("\n请输入【分班结果 Excel】的路径:")
        raw_path = input("> ").strip().replace('"', '').replace("'", "")

        if os.path.exists(raw_path):
            analyze_class_balance(raw_path)
            break
        else:
            print("❌ 路径无效，请重新输入")

    input("\n按回车键退出...")