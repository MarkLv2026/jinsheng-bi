# -*- coding: utf-8 -*-
"""
=============================================================================
电商数据看板 — 空白模板
=============================================================================
基于 Streamlit + Plotly，6 大模块，开箱即用。
换品牌/渠道只需修改 config.py，无需改动此文件。

架构：
  config.py          → 品牌、店铺、配色等配置
  dashboard_core.py  → Excel 数据解析引擎
  app.py             → 前端看板（本文件）

部署：参考 config.py 中的 _PUSH_REPO 配置，推送到 GitHub 后
      在 share.streamlit.io 一键部署。
=============================================================================
"""
from __future__ import annotations
import datetime
import hashlib
import io
import json
import os
import pathlib
import sys as _sys
import traceback as _tb

import streamlit as st

# ── 配置导入 ──────────────────────────────────────────────────
from config import (
    BRAND_NAME, BRAND_SHORT, PAGE_TITLE, HERO_TITLE, HERO_SUBTITLE,
    KNOWN_SHOPS, _SALT, THEME,
)

# ── 第三方库 ──────────────────────────────────────────────────
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from dashboard_core import parse_sales_workbook, rows_to_csv

# ── 缓存目录 ──────────────────────────────────────────────────
_CACHE_DIR = pathlib.Path(__file__).parent / '.data_cache'
_CACHE_DIR.mkdir(exist_ok=True)

# ── 常量 ──────────────────────────────────────────────────────
METRICS = ['商品访客数', '商品浏览量', '商品加购人数', '商品加购件数',
           '支付买家数', '支付件数', '支付金额', '成功退款金额']


# ═══════════════════════════════════════════════════════════════
# 页面基础设置
# ═══════════════════════════════════════════════════════════════
st.set_page_config(
    page_title=PAGE_TITLE,
    page_icon='📊',
    layout='wide',
    initial_sidebar_state='expanded',
)

# ── 自定义 CSS ────────────────────────────────────────────────
CSS = '''
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', -apple-system, sans-serif; }
    
    /* 顶部横幅 */
    .hero {
        background: linear-gradient(135deg, #1e3a5f 0%, #2F6DF6 100%);
        border-radius: 16px; padding: 28px 36px; margin-bottom: 24px; color: #fff;
    }
    .hero .badge {
        display: inline-block; background: rgba(255,255,255,0.2);
        padding: 3px 12px; border-radius: 20px; font-size: 12px;
        margin-right: 8px; margin-bottom: 10px;
    }
    .hero h1 { font-size: 28px; font-weight: 700; margin: 8px 0 4px; }
    .hero .hero-sub { font-size: 14px; opacity: 0.85; }
    
    /* KPI 卡片 */
    .kpi-card {
        background: #fff; border-radius: 12px; padding: 18px 20px;
        box-shadow: 0 1px 4px rgba(0,0,0,0.06); text-align: center;
    }
    .kpi-card .label { font-size: 12px; color: #64748b; margin-bottom: 4px; }
    .kpi-card .value { font-size: 26px; font-weight: 700; color: #0f172a; }
    .kpi-card .change { font-size: 12px; margin-top: 4px; }
    .change-up { color: #16a34a; }
    .change-down { color: #dc2626; }

    /* 区块标题 */
    .section-title {
        font-size: 18px; font-weight: 700; color: #0f172a;
        border-left: 4px solid #2F6DF6; padding-left: 12px;
        margin: 28px 0 16px;
    }
    
    /* 空状态 */
    .empty-state {
        text-align: center; padding: 60px 20px; color: #94a3b8;
    }
    .empty-state .icon { font-size: 48px; margin-bottom: 12px; }
    .empty-state .title { font-size: 18px; font-weight: 600; color: #64748b; }
    .empty-state .desc { font-size: 13px; margin-top: 6px; }
    
    /* 隐藏 Streamlit 默认元素 */
    #MainMenu { visibility: hidden; }
    footer { visibility: hidden; }
    .stDeployButton { display: none; }
</style>
'''
st.markdown(CSS, unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════
# 登录系统
# ═══════════════════════════════════════════════════════════════
def init_session():
    """初始化 session_state"""
    defaults = {
        'authenticated': False, 'username': '', 'role': 'viewer',
        'sales_data': None, 'promo_data': None, 'target_data': None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def load_users():
    """加载用户配置"""
    users_path = pathlib.Path(__file__).parent / 'users.json'
    if users_path.exists():
        with open(users_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {'users': {}}


def check_password(username, password):
    """验证密码"""
    users = load_users()
    if username not in users.get('users', {}):
        return False
    user = users['users'][username]
    h = hashlib.sha256((_SALT + password).encode()).hexdigest()
    return h == user.get('password_hash', '')


def login_page():
    """登录页面"""
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown(f'''
        <div class="hero">
            <div>
                <span class="badge">电商数据看板</span>
                <span class="badge">空白模板</span>
            </div>
            <h1>{HERO_TITLE}</h1>
            <div class="hero-sub">请输入账号密码登录</div>
        </div>
        ''', unsafe_allow_html=True)

        with st.form('login_form'):
            username = st.text_input('账号', placeholder='请输入账号')
            password = st.text_input('密码', type='password', placeholder='请输入密码')
            submitted = st.form_submit_button('登 录', use_container_width=True)

            if submitted:
                if check_password(username, password):
                    users = load_users()
                    user_info = users['users'].get(username, {})
                    st.session_state.authenticated = True
                    st.session_state.username = username
                    st.session_state.role = user_info.get('role', 'viewer')
                    st.rerun()
                else:
                    st.error('账号或密码错误')


# ═══════════════════════════════════════════════════════════════
# 侧边栏 — 数据上传 & 筛选
# ═══════════════════════════════════════════════════════════════
def render_sidebar():
    """渲染侧边栏"""
    with st.sidebar:
        role_label = '管理员' if st.session_state.role == 'admin' else '浏览者'
        st.markdown(f'👤 {st.session_state.username}（{role_label}）')

        if st.session_state.role == 'admin':
            st.divider()
            st.header('📤 数据上传')

            # 销售数据
            with st.expander('📦 销售数据', expanded=not st.session_state.sales_data):
                sales_file = st.file_uploader(
                    '上传销售 Excel',
                    type=['xlsx', 'xls'],
                    key='sales_upload',
                    help='需包含"天猫数据源"或"京东抖音数据源"工作表',
                )
                if sales_file:
                    try:
                        st.session_state.sales_data = parse_sales_workbook(sales_file)
                        st.success(f'✅ 已加载 {st.session_state.sales_data["meta"]["rows"]} 条记录')
                    except Exception as e:
                        st.error(f'解析失败：{e}')

            # 推广数据
            with st.expander('📢 推广数据', expanded=True):
                promo_file = st.file_uploader(
                    '上传推广 Excel',
                    type=['xlsx', 'xls'],
                    key='promo_upload',
                    help='需包含"京东推广数据源"或"天猫推广数据源"工作表',
                )
                if promo_file:
                    try:
                        st.session_state.promo_data = promo_file
                        st.success('✅ 推广数据已上传')
                    except Exception as e:
                        st.error(f'解析失败：{e}')

            # 目标数据
            with st.expander('🎯 目标数据', expanded=True):
                target_file = st.file_uploader(
                    '上传目标 Excel',
                    type=['xlsx', 'xls'],
                    key='target_upload',
                    help='含"X年X月目标拆解及登记"工作表',
                )
                if target_file:
                    st.session_state.target_data = target_file
                    st.success('✅ 目标数据已上传')

        # 全局筛选器
        st.divider()
        st.header('🔍 筛选器')

        # 时间范围
        if st.session_state.sales_data:
            dates = st.session_state.sales_data.get('meta', {}).get('dateRange', [])
            if dates and len(dates) == 2:
                date_range = st.date_input(
                    '时间范围',
                    value=[datetime.date.fromisoformat(dates[0]),
                           datetime.date.fromisoformat(dates[1])],
                    key='date_filter',
                )
                st.session_state.date_start = date_range[0].isoformat() if len(date_range) > 0 else dates[0]
                st.session_state.date_end = date_range[1].isoformat() if len(date_range) > 1 else dates[1]

        # 渠道筛选
        if st.session_state.sales_data:
            channels = st.session_state.sales_data.get('filters', {}).get('channels', [])
            st.multiselect('渠道', channels, default=channels, key='channel_filter')

        # 店铺筛选
        if st.session_state.sales_data:
            stores = st.session_state.sales_data.get('filters', {}).get('stores', [])
            st.multiselect('店铺', stores, default=stores, key='store_filter')

        # 退出登录
        st.divider()
        if st.button('🚪 退出登录', use_container_width=True):
            st.session_state.authenticated = False
            st.session_state.username = ''
            st.rerun()


# ═══════════════════════════════════════════════════════════════
# 工具函数
# ═══════════════════════════════════════════════════════════════
def fmt_num(n):
    """智能格式化数字"""
    if n is None:
        return '—'
    n = float(n)
    if abs(n) >= 1e8:
        return f'{n/1e8:.2f}亿'
    if abs(n) >= 1e4:
        return f'{n/1e4:.1f}万'
    if n == int(n):
        return f'{int(n):,}'
    return f'{n:,.2f}'


def fmt_pct(n):
    """格式化百分比"""
    if n is None or n == 0:
        return '—'
    return f'{float(n)*100:.1f}%'


def fmt_change(curr, prev):
    """计算变化率"""
    if not prev or prev == 0:
        return '', ''
    pct = (curr - prev) / prev
    color = 'change-up' if pct >= 0 else 'change-down'
    arrow = '↑' if pct >= 0 else '↓'
    return f'{arrow} {abs(pct)*100:.1f}%', color


def empty_state(icon, title, desc=''):
    """空状态提示"""
    st.markdown(f'''
    <div class="empty-state">
        <div class="icon">{icon}</div>
        <div class="title">{title}</div>
        <div class="desc">{desc}</div>
    </div>
    ''', unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════
# 模块 1：经营总览
# ═══════════════════════════════════════════════════════════════
def tab_overview(data):
    """经营总览页 — KPI 卡片 + 渠道分布 + 店铺排行 + 月度趋势"""
    if not data:
        empty_state('📊', '暂无销售数据',
                     '请在左侧边栏上传销售 Excel 文件，系统将自动生成经营总览。')
        return

    totals = data.get('totals', {})
    st.markdown('<div class="section-title">📈 核心指标概览</div>', unsafe_allow_html=True)

    # KPI 卡片行
    kpi_items = [
        ('支付金额', 'GMV', '¥'),
        ('支付买家数', '买家数', ''),
        ('支付件数', '销量', ''),
        ('客单价', '客单价', '¥'),
        ('支付转化率', '转化率', ''),
    ]
    cols = st.columns(len(kpi_items))
    for i, (key, label, prefix) in enumerate(kpi_items):
        val = totals.get(key, 0)
        display = f'{prefix}{fmt_num(val)}' if prefix else fmt_num(int(val))
        with cols[i]:
            st.markdown(f'''
            <div class="kpi-card">
                <div class="label">{label}</div>
                <div class="value">{display}</div>
            </div>
            ''', unsafe_allow_html=True)

    # 两列布局：渠道饼图 + 店铺柱状图
    col_left, col_right = st.columns(2)

    with col_left:
        st.markdown('<div class="section-title">📡 渠道销售分布</div>', unsafe_allow_html=True)
        channels = data.get('channels', [])
        if channels:
            df_ch = pd.DataFrame(channels)
            fig = px.pie(df_ch, values='支付金额', names='渠道',
                         hole=0.4, color_discrete_sequence=px.colors.qualitative.Set2)
            fig.update_traces(textinfo='percent+label', textfont_size=12)
            fig.update_layout(height=360, margin=dict(t=10, b=10, l=10, r=10))
            st.plotly_chart(fig, use_container_width=True)

    with col_right:
        st.markdown('<div class="section-title">🏪 店铺销售排行</div>', unsafe_allow_html=True)
        stores = data.get('stores', [])
        if stores:
            df_st = pd.DataFrame(stores).head(10)
            fig = px.bar(df_st, x='店铺', y='支付金额', color='支付金额',
                         color_continuous_scale='Blues', text_auto='.2s')
            fig.update_layout(height=360, margin=dict(t=10, b=10, l=10, r=10))
            st.plotly_chart(fig, use_container_width=True)

    # 月度趋势
    st.markdown('<div class="section-title">📅 月度销售趋势</div>', unsafe_allow_html=True)
    all_months = data.get('all_months', [])
    if all_months:
        df_m = pd.DataFrame(all_months)
        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=df_m['月份'], y=df_m['支付金额'], mode='lines+markers',
            fill='tozeroy', fillcolor='rgba(47,109,246,0.1)',
            line=dict(color='#2F6DF6', width=2),
            name='GMV'
        ))
        fig.update_layout(height=320, margin=dict(t=10, b=10, l=10, r=10),
                          hovermode='x unified')
        st.plotly_chart(fig, use_container_width=True)

    # 品类排行
    st.markdown('<div class="section-title">📦 品类销售额排行</div>', unsafe_allow_html=True)
    categories = data.get('categories', [])
    if categories:
        df_cat = pd.DataFrame(categories)
        st.dataframe(df_cat, use_container_width=True, hide_index=True,
                     column_config={'支付金额': st.column_config.NumberColumn(format='¥%.0f')})


# ═══════════════════════════════════════════════════════════════
# 模块 2：推广分析
# ═══════════════════════════════════════════════════════════════
def tab_promotion(data):
    """推广分析页 — 花费 vs 成交 + ROI + 效率矩阵"""
    if not data:
        empty_state('📢', '暂无推广数据',
                     '请在左侧边栏上传推广 Excel 文件，系统将自动生成推广分析。')
        return

    st.markdown('<div class="section-title">📢 推广数据概览</div>', unsafe_allow_html=True)
    st.info('上传推广数据后，此处将展示：推广花费趋势、ROI 分析、效率矩阵、TOP 推广计划等。')
    st.markdown('''
    ### 即将支持的模块：
    - 📊 推广花费 & 成交金额趋势图
    - 📈 ROI 趋势分析
    - 🎯 推广效率矩阵（花费 vs 成交金额）
    - 🏪 店铺/渠道推广矩阵
    - 📦 单品推广分析
    - 🔝 TOP10 推广计划排行

    > 💡 推广 Excel 需包含「京东推广数据源」和「天猫推广数据源」工作表。
    ''')


# ═══════════════════════════════════════════════════════════════
# 模块 3：时间段对比
# ═══════════════════════════════════════════════════════════════
def tab_comparison(data):
    """时间段对比页 — 两个时间段的指标对比"""
    if not data:
        empty_state('📊', '暂无销售数据',
                     '上传销售数据后，可选择两个时间段进行对比分析。')
        return

    st.markdown('<div class="section-title">📊 时间段对比分析</div>', unsafe_allow_html=True)
    st.info('选择两个时间段，系统将自动对比：GMV、订单量、转化率、渠道分布等核心指标。')

    col1, col2 = st.columns(2)
    with col1:
        st.selectbox('基准期', ['本月', '上月', '自定义'], key='cmp_period1')
    with col2:
        st.selectbox('对比期', ['上月', '去年同期', '自定义'], key='cmp_period2')

    st.markdown('### 即将支持的对比维度：')
    st.markdown('''
    - 📈 核心 KPI 对比（GMV、买家数、客单价、转化率）
    - 📡 渠道结构变化
    - 🏪 店铺表现对比
    - 📦 品类/型号增长排行
    - 📢 推广效率对比
    ''')


# ═══════════════════════════════════════════════════════════════
# 模块 4：趋势分析
# ═══════════════════════════════════════════════════════════════
def tab_trends(data):
    """趋势分析页 — 日度/月度趋势 + 同比增长"""
    if not data:
        empty_state('📈', '暂无销售数据',
                     '上传销售数据后，可查看销售趋势和同比分析。')
        return

    st.markdown('<div class="section-title">📈 趋势分析</div>', unsafe_allow_html=True)

    # 日度趋势
    daily = data.get('daily', [])
    if daily:
        st.markdown('### 📅 日度销售趋势')
        df_d = pd.DataFrame(daily)
        df_agg = df_d.groupby('日期').agg({'支付金额': 'sum'}).reset_index()

        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=df_agg['日期'], y=df_agg['支付金额'],
            mode='lines', line=dict(color='#2F6DF6', width=2),
            fill='tozeroy', fillcolor='rgba(47,109,246,0.08)',
        ))
        fig.update_layout(height=340, margin=dict(t=10, b=10, l=10, r=10),
                          hovermode='x unified')
        st.plotly_chart(fig, use_container_width=True)

    # 月度趋势
    monthly = data.get('monthly', [])
    if monthly:
        st.markdown('### 📊 月度销售趋势')
        df_m = pd.DataFrame(monthly)
        df_m_agg = df_m.groupby('月份').agg({'支付金额': 'sum'}).reset_index()

        fig = px.bar(df_m_agg, x='月份', y='支付金额',
                     color='支付金额', color_continuous_scale='Blues',
                     text_auto='.2s')
        fig.update_layout(height=340, margin=dict(t=10, b=10, l=10, r=10))
        st.plotly_chart(fig, use_container_width=True)

    # 同比趋势（需要至少12个月数据展示占位）
    st.markdown('### 📊 同比增长趋势')
    st.info('当数据跨度超过 12 个月时，自动展示同比增长曲线。')

    # 数据明细
    with st.expander('📋 查看数据明细', expanded=False):
        tab_day, tab_month = st.tabs(['日度汇总', '月度汇总'])
        with tab_day:
            if daily:
                st.dataframe(pd.DataFrame(daily).head(100), use_container_width=True, hide_index=True)
        with tab_month:
            if monthly:
                st.dataframe(pd.DataFrame(monthly), use_container_width=True, hide_index=True)


# ═══════════════════════════════════════════════════════════════
# 模块 5：智能诊断
# ═══════════════════════════════════════════════════════════════
def tab_diagnosis(data):
    """智能诊断页 — 产品排行 + 转化分析"""
    if not data:
        empty_state('🔍', '暂无销售数据',
                     '上传销售数据后，系统将自动进行智能诊断分析。')
        return

    st.markdown('<div class="section-title">🔍 智能诊断</div>', unsafe_allow_html=True)

    diag_tabs = st.tabs(['📦 产品分析', '📊 转化漏斗', '⚡ 异常检测'])

    # 产品分析
    with diag_tabs[0]:
        st.markdown('#### TOP 20 热销型号')
        models = data.get('models', [])
        if models:
            df_model = pd.DataFrame(models[:20])
            fig = px.bar(df_model, x='型号', y='支付金额', color='支付金额',
                         color_continuous_scale='Blues')
            fig.update_layout(height=400, margin=dict(t=10, b=10, l=10, r=10))
            st.plotly_chart(fig, use_container_width=True)

            st.markdown('#### 型号明细表')
            st.dataframe(df_model, use_container_width=True, hide_index=True,
                         column_config={'支付金额': st.column_config.NumberColumn(format='¥%.0f')})

    # 转化漏斗
    with diag_tabs[1]:
        st.markdown('#### 全域转化漏斗')
        totals = data.get('totals', {})
        visit = int(totals.get('商品访客数', 0))
        view = int(totals.get('商品浏览量', 0))
        cart = int(totals.get('商品加购人数', 0))
        buyer = int(totals.get('支付买家数', 0))
        if visit > 0:
            stages = ['访客', '浏览', '加购', '支付']
            values = [visit, view, cart, buyer]
            fig = go.Figure(go.Funnel(
                y=stages, x=values,
                textposition='inside',
                textinfo='value+percent initial',
                marker=dict(color=['#93c5fd', '#60a5fa', '#3b82f6', '#1d4ed8']),
            ))
            fig.update_layout(height=360, margin=dict(t=10, b=10, l=10, r=10))
            st.plotly_chart(fig, use_container_width=True)

    # 异常检测
    with diag_tabs[2]:
        st.info('异常检测模块将在数据充足时自动识别：GMV 骤降、转化率异常、客单价波动等。')
        st.markdown('### 即将支持的诊断维度：')
        st.markdown('''
        - 📉 销售额异常波动检测
        - 🔻 转化率骤降型号预警
        - 💰 客单价异常变化
        - 📦 库存周转风险提示
        - 🎯 单品健康度评分
        ''')


# ═══════════════════════════════════════════════════════════════
# 模块 6：目标达成
# ═══════════════════════════════════════════════════════════════
def tab_targets(data):
    """目标达成页 — 目标进度 + 完成率"""
    if not data:
        empty_state('🎯', '暂无目标数据',
                     '请在左侧边栏上传目标 Excel 文件（含"X年X月目标拆解及登记"工作表），系统将自动追踪目标达成情况。')
        return

    totals = data.get('totals', {})
    st.markdown('<div class="section-title">🎯 目标达成追踪</div>', unsafe_allow_html=True)

    cols = st.columns(4)
    with cols[0]:
        st.metric('月目标 GMV', '¥ —', '点击设置')
    with cols[1]:
        st.metric('当月实际', f'¥ {fmt_num(totals.get("支付金额", 0))}', '—')
    with cols[2]:
        st.metric('达成率', '—%', '—')
    with cols[3]:
        st.metric('缺口', '¥ —', '—')

    st.markdown('### 即将支持的功能：')
    st.markdown('''
    - 🎯 月度/年度目标设定与追踪
    - 📊 目标达成进度可视化（仪表盘）
    - 🏪 店铺维度目标分解
    - 📦 单品目标达成明细
    - 📅 日度目标完成趋势
    - ⚠️ 目标偏离预警
    ''')


# ═══════════════════════════════════════════════════════════════
# 主程序
# ═══════════════════════════════════════════════════════════════
def main():
    init_session()

    # 未登录 → 显示登录页
    if not st.session_state.authenticated:
        login_page()
        return

    # 已登录 → 渲染侧边栏
    render_sidebar()

    # ── 顶部横幅 ──
    st.markdown(f'''
    <div class="hero">
        <div>
            <span class="badge">{BRAND_NAME}</span>
            <span class="badge">电商数据看板</span>
            <span class="badge">空白模板</span>
        </div>
        <h1>{HERO_TITLE}</h1>
        <div class="hero-sub">{HERO_SUBTITLE}</div>
    </div>
    ''', unsafe_allow_html=True)

    # ── 数据状态栏 ──
    data_loaded = bool(st.session_state.sales_data)
    promo_loaded = bool(st.session_state.promo_data)
    target_loaded = bool(st.session_state.target_data)
    status_cols = st.columns(3)
    with status_cols[0]:
        st.metric('销售数据', '✅ 已加载' if data_loaded else '❌ 未加载')
    with status_cols[1]:
        st.metric('推广数据', '✅ 已加载' if promo_loaded else '❌ 未加载')
    with status_cols[2]:
        st.metric('目标数据', '✅ 已加载' if target_loaded else '❌ 未加载')

    # ── 6 大标签页 ──
    tabs = st.tabs([
        '📊 经营总览',
        '📢 推广分析',
        '📊 时间段对比',
        '📈 趋势分析',
        '🔍 智能诊断',
        '🎯 目标达成',
    ])

    sales = st.session_state.sales_data
    promo = st.session_state.promo_data

    with tabs[0]:
        tab_overview(sales)
    with tabs[1]:
        tab_promotion(promo)
    with tabs[2]:
        tab_comparison(sales)
    with tabs[3]:
        tab_trends(sales)
    with tabs[4]:
        tab_diagnosis(sales)
    with tabs[5]:
        tab_targets(sales)

    # ── 底部 ──
    st.divider()
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        st.caption(f'{BRAND_NAME} 电商数据看板 · 空白模板')
    with col_f2:
        st.caption(f'数据更新时间：{datetime.datetime.now().strftime("%Y-%m-%d %H:%M")}')


if __name__ == '__main__':
    main()
