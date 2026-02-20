import pandas as pd
import numpy as np
import random
import sys
import os


class SunshineDistributor:
    def __init__(self, n_classes, top_n_per_class=10, bottom_n_per_class=10):
        self.K = int(n_classes)
        self.top_N = top_n_per_class * self.K
        self.bot_N = bottom_n_per_class * self.K
        # 修复：初始化必需列名
        self.required_cols = ['姓名', '性别', '总分']
        self.has_rural = False

    def load_data(self, file_path):
        if not os.path.exists(file_path):
            print(f"❌ 错误：找不到文件！\n路径: {file_path}")
            sys.exit(1)

        try:
            df = pd.read_excel(file_path)

            # 检查列名
            if '城乡' in df.columns:
                self.has_rural = True
                self.required_cols.append('城乡')
                print("ℹ️ 检测到'城乡'列，启用四维均衡分班。")
            else:
                print("ℹ️ 未检测到'城乡'列，启用二维均衡分班。")

            # 修复：检查缺失列
            missing_cols = [col for col in self.required_cols if col not in df.columns]
            if missing_cols:
                print(f"❌ 错误：Excel 中缺少以下列: {missing_cols}")
                print(f"   请确保表头包含: {self.required_cols}")
                sys.exit(1)

            # 清理数据：如果源表有空的'班级'或'年级排名'列，先删掉
            # 修复：补全列表
            for col in ['班级', '年级排名']:
                if col in df.columns:
                    df = df.drop(columns=[col])

            original_len = len(df)

            # 修复：只将'总分'转为数字，避免姓名变成NaN
            df['总分'] = pd.to_numeric(df['总分'], errors='coerce')

            # 修复：删除总分无效的行
            df = df.dropna(subset=['总分'])

            print(f"✅ 读取成功: 共 {original_len} 行，有效数据 {len(df)} 行")
            return df

        except Exception as e:
            print(f"❌ 读取发生错误: {e}")
            sys.exit(1)

    def _distribute_sub_group(self, sub_df, assigned_list):
        if sub_df.empty: return

        # 组内按成绩降序
        sub_df = sub_df.sort_values(by='总分', ascending=False).copy()

        # 盲选打乱班级序列 (防止每次都是1班拿第一名)
        class_indices = list(range(1, self.K + 1))
        random.shuffle(class_indices)

        # 修复：蛇形填充逻辑 (S型：正序 -> 逆序 -> 正序...)
        snake_pattern = []
        while len(snake_pattern) < len(sub_df):
            snake_pattern.extend(class_indices)  # 正序 [1, 2, 3...]
            snake_pattern.extend(class_indices[::-1])  # 逆序 [3, 2, 1...]

        # 截取对应长度并赋值
        sub_df['班级'] = snake_pattern[:len(sub_df)]
        assigned_list.append(sub_df)

    def run(self, file_path):
        print("🔄 正在计算最优分班方案...")
        df = self.load_data(file_path)

        # === 🌟 新增：计算全年级排名 ===
        # 修复：只对总分排名
        df['年级排名'] = df['总分'].rank(method='min', ascending=False).astype(int)

        # 确保整体按总分降序排列
        df = df.sort_values(by='总分', ascending=False).reset_index(drop=True)

        if len(df) < self.top_N + self.bot_N:
            print(f"❌ 错误：学生总数不足 (需至少 {self.top_N + self.bot_N} 人)")
            sys.exit(1)

        # 修复：三段式切片参数
        df_top = df.iloc[:self.top_N].copy()
        df_bot = df.iloc[-self.bot_N:].copy()
        # 中间层是去掉头尾剩下的
        df_mid = df.iloc[self.top_N: -self.bot_N].copy()

        assigned_list = []
        # 修复：定义层级列表
        layers = [df_top, df_mid, df_bot]

        # 分组处理
        for layer in layers:
            if self.has_rural:
                # 四维分组：性别 + 城乡
                # 修复：Pandas 筛选语法
                groups = [
                    layer[(layer['性别'] == '男') & (layer['城乡'] == '城区')],
                    layer[(layer['性别'] == '男') & (layer['城乡'] == '乡下')],
                    layer[(layer['性别'] == '女') & (layer['城乡'] == '城区')],
                    layer[(layer['性别'] == '女') & (layer['城乡'] == '乡下')]
                ]
            else:
                # 二维分组：性别
                groups = [
                    layer[layer['性别'] == '男'],
                    layer[layer['性别'] == '女']
                ]

            for group in groups:
                self._distribute_sub_group(group, assigned_list)

        result = pd.concat(assigned_list)

        # === 关键步骤：调整列顺序和排序 ===
        # 1. 排序：先按班级排序(1,2,3...)，再按总分降序
        # 修复：填充排序参数
        result = result.sort_values(by=['班级', '总分'], ascending=[True, False])

        # 2. 调整列位置：把"班级"和"年级排名"挪到最前面
        # 修复：列表生成
        cols = ['班级', '年级排名'] + [c for c in result.columns if c not in ['班级', '年级排名']]
        result = result[cols]

        return result

    def export_excel(self, df, output_path):
        # 动态构建统计指标
        agg_funcs = {
            '姓名': 'count',
            '总分': 'mean',
            '性别': lambda x: (x == '男').sum()
        }
        rename_cols = {'姓名': '总人数', '总分': '平均分', '性别': '男生数'}

        if self.has_rural:
            # 修复：添加城乡统计
            agg_funcs['城乡'] = lambda x: (x == '城区').sum()
            rename_cols['城乡'] = '城区数'

        stats = df.groupby('班级').agg(agg_funcs).rename(columns=rename_cols)

        # 计算剩余列
        # 修复：计算女生数
        stats['女生数'] = stats['总人数'] - stats['男生数']
        stats = stats.round(2)

        final_cols = ['总人数', '平均分', '男生数', '女生数']

        if self.has_rural:
            # 修复：计算乡下数
            stats['乡下数'] = stats['总人数'] - stats['城区数']
            final_cols.extend(['城区数', '乡下数'])
            stats = stats[final_cols]
        else:
            stats = stats[final_cols]

        print("\n📈 质量报告:")
        print(f"   平均分最大分差: {stats['平均分'].max() - stats['平均分'].min():.2f}")

        try:
            with pd.ExcelWriter(output_path) as writer:
                df.to_excel(writer, sheet_name='详细名单', index=False)
                stats.to_excel(writer, sheet_name='统计报表')
            print(f"\n🎉 成功！结果文件已生成:\n👉 {output_path}")
        except PermissionError:
            print(f"\n❌ 保存失败！文件被占用，请确保已经把生成的Excel文件关掉。")


# ================= 主程序 =================

if __name__ == "__main__":
    print("=" * 50)
    print("      SJTU 阳光分班系统 (修复完整版)")
    print("=" * 50)

    # 1. 获取路径
    input_path = ""
    while True:
        print("\n请输入Excel文件的完整路径 (例如 D:\\data\\student.xlsx):")
        raw_input = input("> ").strip()
        # 去除引号 (兼容直接拖拽文件)
        clean_path = raw_input.replace('"', '').replace("'", "")
        clean_path = os.path.normpath(clean_path)

        if os.path.exists(clean_path):
            input_path = clean_path
            break
        else:
            print("❌ 路径无效，请重新输入")

    # 2. 获取班级数
    CLASS_COUNT = 0
    while True:
        try:
            val = input("\n请输入班级数量 (例如 16):\n> ").strip()
            CLASS_COUNT = int(val)
            if CLASS_COUNT > 0: break
        except ValueError:
            print("⚠️ 请输入数字")  # <--- 注意这里要有缩进

    # 3. 运行
    distributor = SunshineDistributor(CLASS_COUNT, top_n_per_class=10, bottom_n_per_class=10)
    final_df = distributor.run(input_path)

    # 4. 导出
    input_dir = os.path.dirname(input_path)
    base_name = os.path.splitext(os.path.basename(input_path))[0]
    output_path = os.path.join(input_dir, f"{base_name}_分班结果.xlsx")

    distributor.export_excel(final_df, output_path)

    input("\n✅ 所有步骤完成，按回车键退出...")