import pandas as pd
import streamlit as st
from datetime import datetime
import plotly.express as px
import locale

# 设置页面布局为宽屏
st.set_page_config(layout="wide")

# 设置中文语言环境，使日期选择器显示中文
locale.setlocale(locale.LC_ALL, 'zh_Hans_CN.UTF-8')

# 自定义 CSS 来调整容器宽度
def add_custom_css():
    st.markdown(
        """
        <style>
        .main > div {
            max-width: 100%;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

# 读取 Excel 数据（使用新版缓存装饰器）
@st.cache_data
def load_data(file):
    df = pd.read_excel(file, sheet_name="sheet1")
    original_length = len(df)
    # 处理可能的缺失值，删除领卡时间为 NaT 的行
    df = df.dropna(subset=["领卡时间"])
    removed_count = original_length - len(df)
    if removed_count > 0:
        st.warning(f"由于领卡时间存在缺失值，已删除 {removed_count} 行数据。")
    df["领卡时间"] = pd.to_datetime(df["领卡时间"])

    # 检查必要列名是否存在
    required_columns = ["手机号", "推荐人手机号", "姓名", "等级", "直推订单数", "直推订单金额",
                        "自购订单数", "自购订单金额", "自购订单实体卡", "团队订单数",
                        "团队订单金额", "团队订单实体卡"]
    for col in required_columns:
        if col not in df.columns:
            st.error(f"Excel 文件中缺少 '{col}' 列，请检查文件。")
            return None

    # 确保手机号和推荐人手机号为字符串类型
    df["手机号"] = df["手机号"].astype(str).str.strip()
    df["推荐人手机号"] = df["推荐人手机号"].astype(str).str.strip()
    # 处理推荐人手机号为空的情况
    df = df[df["推荐人手机号"] != '']
    empty_referrer_count = original_length - len(df)
    if empty_referrer_count > 0:
        st.warning(f"由于推荐人手机号为空，已删除 {empty_referrer_count} 行数据。")
    return df

# 查找所有下级（递归函数），添加递归深度限制
def find_all_subordinates(df, user_phone, all_subs=None, depth=0, max_depth=100):
    if all_subs is None:
        all_subs = []
    if depth > max_depth:
        st.warning("递归深度超过限制，可能存在循环引用，请检查数据。")
        return all_subs
    direct_subs = df[df["推荐人手机号"] == user_phone]
    if not direct_subs.empty:
        all_subs.extend(direct_subs["手机号"].tolist())
        for sub in direct_subs["手机号"]:
            find_all_subordinates(df, sub, all_subs, depth + 1, max_depth)
    return all_subs

# 计算统计指标
def calculate_metrics(df, level):
    if level == "direct":
        sub_count = len(df)
        black_card_count = df[df["等级"] == "黑金卡"].shape[0]
        order_count = df["直推订单数"].sum() if "直推订单数" in df.columns else 0
        order_amount = df["直推订单金额"].sum() if "直推订单金额" in df.columns else 0
    else:
        sub_count = len(df)
        black_card_count = df[df["等级"] == "黑金卡"].shape[0]
        order_count = df["团队订单数"].sum() if "团队订单数" in df.columns else 0
        order_amount = df["团队订单金额"].sum() if "团队订单金额" in df.columns else 0
    return sub_count, black_card_count, order_count, order_amount

# 构建仪表盘
def main():
    add_custom_css()  # 添加自定义 CSS
    st.title("用户关系分析仪表盘")

    # 文件选择器，修改提示文字为中文
    file = st.file_uploader("请选择 Excel 文件", type=["xlsx"])

    if file is not None:
        df = load_data(file)
        if df is None:
            return

        # 侧边栏日期筛选
        st.sidebar.header("筛选条件")
        # 确保领卡时间列不为空
        if not df["领卡时间"].empty:
            start_date = st.sidebar.date_input("开始日期", df["领卡时间"].min().date())
            end_date = st.sidebar.date_input("结束日期", df["领卡时间"].max().date())
        else:
            st.warning("领卡时间列无有效数据，请检查文件。")
            return

        # 主界面搜索用户
        search_name = st.text_input("输入用户姓名搜索")

        if search_name:
            target_user = df[df["姓名"] == search_name]
            if target_user.empty:
                st.warning(f"未找到姓名为 '{search_name}' 的用户，请检查姓名拼写是否正确。")
                return

            user_phone = target_user["手机号"].values[0]

            # 按日期筛选数据
            filtered_df = df[(df["领卡时间"] >= pd.Timestamp(start_date)) &
                             (df["领卡时间"] <= pd.Timestamp(end_date))]
            st.write(f"筛选前数据量: {len(df)}, 筛选后数据量: {len(filtered_df)}")

            # 查找直推下级
            direct_subs = filtered_df[filtered_df["推荐人手机号"] == user_phone]

            # 查找所有下级
            all_subs_phones = find_all_subordinates(filtered_df, user_phone)
            all_subs = filtered_df[filtered_df["手机号"].isin(all_subs_phones)]

            # 显示统计指标
            col1, col2 = st.columns(2)
            with col1:
                direct_sub_count, direct_black_card_count, direct_order_count, direct_order_amount = calculate_metrics(direct_subs, "direct")
                st.write(f"直推下级人数计算结果: {direct_sub_count}")
                st.metric("直推下级人数", direct_sub_count)
                st.write(f"直推黑金卡数计算结果: {direct_black_card_count}")
                st.metric("直推黑金卡数", direct_black_card_count)
                st.write(f"直推订单总数计算结果: {direct_order_count}")
                st.metric("直推订单总数", direct_order_count)
                st.write(f"直推订单金额计算结果: {direct_order_amount}")
                st.metric("直推订单金额", f"¥{direct_order_amount:,.2f}")

            with col2:
                all_sub_count, all_black_card_count, team_order_count, team_order_amount = calculate_metrics(all_subs, "all")
                st.write(f"所有下级人数计算结果: {all_sub_count}")
                st.metric("所有下级人数", all_sub_count)
                st.write(f"所有黑金卡数计算结果: {all_black_card_count}")
                st.metric("所有黑金卡数", all_black_card_count)
                st.write(f"团队订单总数计算结果: {team_order_count}")
                st.metric("团队订单总数", team_order_count)
                st.write(f"团队订单金额计算结果: {team_order_amount}")
                st.metric("团队订单金额", f"¥{team_order_amount:,.2f}")

            # 显示下级名单，调整布局
            # 第一行显示直推下级名单和直推下级推广情况
            row1_col1, row1_col2 = st.columns(2)
            with row1_col1:
                st.subheader("直推下级名单")
                st.dataframe(direct_subs[["姓名", "手机号", "等级", "自购订单数", "自购订单金额", "自购订单实体卡"]])

            with row1_col2:
                st.subheader("直推下级推广情况")
                st.dataframe(direct_subs[["姓名", "手机号", "等级", "直推订单数", "直推订单金额", "直推订单实体卡"]])

            # 第二行显示所有下级名单
            st.subheader("所有下级名单")
            st.dataframe(all_subs[["姓名", "手机号", "等级", "团队订单数", "团队订单金额", "团队订单实体卡"]])

            # 绘制订单金额分布图，设置为竖版
            st.subheader("订单金额分布")
            if "团队订单金额" in all_subs.columns and not all_subs["团队订单金额"].empty:
                fig = px.bar(all_subs, y="姓名", x="团队订单金额", orientation='h')  # 设置为竖版
                st.plotly_chart(fig)
            else:
                st.warning("团队订单金额列无有效数据，无法绘制订单金额分布图。")

if __name__ == "__main__":
    main()