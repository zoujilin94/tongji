from io import BytesIO
import pandas as pd
import streamlit as st
from datetime import datetime
import plotly.express as px
import locale

# 设置页面布局为宽屏
st.set_page_config(layout="wide")

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
    try:
        sub_count = len(df)
        black_card_count = df[df["等级"] == "黑金卡"].shape[0] if not df.empty else 0
        order_count = df["直推订单数"].sum() if (level == "direct" and not df.empty) else df["团队订单数"].sum()
        order_amount = df["直推订单金额"].sum() if (level == "direct" and not df.empty) else df["团队订单金额"].sum()
        return sub_count, black_card_count, order_count, order_amount
    except Exception as e:
        st.error(f"计算指标时发生错误: {str(e)}")
        return 0, 0, 0, 0

# 构建仪表盘
def main():
    add_custom_css()
    st.title("用户关系分析仪表盘")

    file = st.file_uploader("请选择 Excel 文件", type=["xlsx"])

    if file is not None:
        df = load_data(file)
        if df is None:
            return

        # 获取所有用户姓名列表（去重）
        all_users = df["姓名"].unique().tolist()

        # 创建多选组件
        selected_names = st.multiselect(
            "选择或搜索多个用户（支持拼音首字母搜索）",
            options=all_users,
            format_func=lambda x: x,
            key="user_selector"
        )

        # 侧边栏日期筛选
        st.sidebar.header("筛选条件")
        start_date = st.sidebar.date_input("开始日期", df["领卡时间"].min().date())
        end_date = st.sidebar.date_input("结束日期", df["领卡时间"].max().date())

        # 按日期筛选数据
        filtered_df = df[(df["领卡时间"] >= pd.Timestamp(start_date)) &
                        (df["领卡时间"] <= pd.Timestamp(end_date))]

        # 存储所有用户数据的列表
        all_users_data = []

        # 进度条
        progress_bar = st.progress(0)
        total_users = len(selected_names)

        for index, name in enumerate(selected_names):
            progress = (index + 1) / total_users
            progress_bar.progress(progress)

            target_user = filtered_df[filtered_df["姓名"] == name]
            if target_user.empty:
                st.warning(f"跳过无效用户：{name}")
                continue

            user_phone = target_user["手机号"].values[0]

            # 查找直推下级
            direct_subs = filtered_df[filtered_df["推荐人手机号"] == user_phone]
            
            # 查找所有下级
            all_subs_phones = find_all_subordinates(filtered_df, user_phone)
            all_subs = filtered_df[filtered_df["手机号"].isin(all_subs_phones)]

            # 计算指标
            direct_metrics = calculate_metrics(direct_subs, "direct")
            all_metrics = calculate_metrics(all_subs, "all")

            # 构建用户数据字典
            user_data = {
                "姓名": name,
                "手机号": user_phone,
                "直推下级人数": direct_metrics[0],
                "直推黑金卡数": direct_metrics[1],
                "直推订单总数": direct_metrics[2],
                "直推订单金额": direct_metrics[3],
                "所有下级人数": all_metrics[0],
                "所有黑金卡数": all_metrics[1],
                "团队订单总数": all_metrics[2],
                "团队订单金额": all_metrics[3],
                "直推下级名单": direct_subs["手机号"].tolist(),
                "所有下级名单": all_subs["手机号"].tolist()
            }
            all_users_data.append(user_data)

        # 生成汇总表格
        summary_df = pd.DataFrame([{
            "姓名": data["姓名"],
            "直推下级人数": data["直推下级人数"],
            "直推黑金卡数": data["直推黑金卡数"],
            "直推订单总数": data["直推订单总数"],
            "直推订单金额": data["直推订单金额"],
            "所有下级人数": data["所有下级人数"],
            "所有黑金卡数": data["所有黑金卡数"],
            "团队订单总数": data["团队订单总数"],
            "团队订单金额": data["团队订单金额"]
        } for data in all_users_data])

        # 显示汇总表格
        st.subheader("多用户汇总统计")
        st.dataframe(summary_df)

        # 导出功能
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            summary_df.to_excel(writer, sheet_name='汇总统计', index=False)
            
            # 添加详细数据
            for data in all_users_data:
                user_sheet_name = f"{data['姓名']}-详情"
                
                # 直推下级名单
                direct_df = filtered_df[filtered_df["手机号"].isin(data["直推下级名单"])]
                direct_df.to_excel(writer, sheet_name=f"{user_sheet_name}-直推", index=False)
                
                # 所有下级名单
                all_df = filtered_df[filtered_df["手机号"].isin(data["所有下级名单"])]
                all_df.to_excel(writer, sheet_name=f"{user_sheet_name}-所有下级", index=False)

        output.seek(0)
        st.download_button(
            label="下载报表",
            data=output,
            file_name=f"用户统计报表_{datetime.now().strftime('%Y%m%d%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


if __name__ == "__main__":
    main()