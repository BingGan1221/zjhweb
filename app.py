# 标准库导入
import sys
import logging
import html
from pathlib import Path
from collections import Counter

# 第三方库导入
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import jieba
from wordcloud import WordCloud
import requests
import shutil

# 页面配置（必须是第一个 Streamlit 命令）
st.set_page_config(
    page_title="Excel评论分析工具",
    page_icon="📊",
    layout="wide"
)

# 日志配置
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# 停用词列表
STOP_WORDS = {
    '的', '了', '和', '是', '就', '都', '而', '及', '与', '着',
    '之', '用', '于', '把', '等', '去', '又', '能', '好', '在',
    '或', '这', '那', '有', '很', '只', '些', '为', '呢', '啊',
    '并', '给', '跟', '还', '个', '之类', '各种', '没有', '非常',
    '可以', '因为', '因此', '所以', '但是', '但', '然后', '如果',
    '虽然', '这样', '这些', '那些', '如此', '只是', '真的', '一个',
}

# 用户自定义停用词
USER_STOP_WORDS = set()

# 建议相关词汇
SUGGESTION_WORDS = {
    '议', '觉得', '希望', '调', '换', '改', '改进', '完善',
    '优化', '提议', '期望', '最好', '应该', '不如', '要是',
    '可以', '或许', '建议', '推荐', '提醒'
}

# 负面情绪词汇
NEGATIVE_WORDS = {
    '累', '无聊', '难受', '差', '糟糕', '失望', '不满', '不好',
    '不行', '垃圾', '烦', '恶心', '坑', '不值', '贵', '慢',
    '差劲', '敷衍', '态度差', '脏', '乱', '吵', '挤', '冷',
    '热', '差评', '退款', '投诉', '举报', '骗', '坑'
}

# 在文件开头添加版本常量
VERSION = "2.0.0"  # 更新版本号
CHART_COLORS = ['#153f36', '#d88b2d', '#337b87', '#c8503e', '#547a44', '#8a6f3a']

# 在文件开头，USER_STOP_WORDS 定义后添加
if 'user_stop_words' not in st.session_state:
    st.session_state.user_stop_words = set()

# 在文件开头添加自定义样式
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,650;9..144,800&family=Noto+Sans+SC:wght@400;500;700;900&display=swap');

    :root {
        --ink: #17201b;
        --muted: #68756f;
        --paper: #fbf7ed;
        --panel: rgba(255, 252, 244, 0.92);
        --panel-strong: #fffdf6;
        --line: rgba(39, 52, 44, 0.14);
        --pine: #153f36;
        --moss: #547a44;
        --amber: #d88b2d;
        --coral: #c8503e;
        --sky: #337b87;
        --shadow: 0 18px 50px rgba(23, 32, 27, 0.10);
        --soft-shadow: 0 8px 24px rgba(23, 32, 27, 0.07);
        --radius: 8px;
    }

    html, body, [data-testid="stAppViewContainer"] {
        color: var(--ink);
        font-family: 'Noto Sans SC', sans-serif !important;
        background:
            radial-gradient(circle at 12% 10%, rgba(216, 139, 45, 0.16), transparent 30%),
            radial-gradient(circle at 88% 6%, rgba(51, 123, 135, 0.13), transparent 28%),
            linear-gradient(135deg, #fbf7ed 0%, #f3ead9 48%, #edf3ea 100%) !important;
    }

    .stApp::before {
        content: "";
        position: fixed;
        inset: 0;
        pointer-events: none;
        z-index: 0;
        opacity: 0.22;
        background-image:
            linear-gradient(rgba(23, 32, 27, 0.06) 1px, transparent 1px),
            linear-gradient(90deg, rgba(23, 32, 27, 0.06) 1px, transparent 1px);
        background-size: 34px 34px;
        mask-image: linear-gradient(to bottom, black, transparent 75%);
    }

    .main .block-container {
        max-width: 1320px;
        padding: 2rem 2rem 4rem;
        position: relative;
        z-index: 1;
    }

    h1, h2, h3 {
        letter-spacing: 0 !important;
    }

    .stTextInput input,
    .stTextArea textarea,
    .stSelectbox select {
        border: 1px solid var(--line) !important;
        border-radius: var(--radius) !important;
        padding: 0.75rem 0.9rem !important;
        background-color: rgba(255, 253, 246, 0.98) !important;
        color: var(--ink) !important;
        font-size: 0.94rem !important;
        box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.55) !important;
    }

    .stTextInput input:focus,
    .stTextArea textarea:focus {
        border-color: var(--amber) !important;
        box-shadow: 0 0 0 3px rgba(216, 139, 45, 0.16) !important;
    }

    .stMultiSelect [data-baseweb="select"] > div {
        background-color: rgba(255, 253, 246, 0.98) !important;
        border: 1px solid var(--line) !important;
        border-radius: var(--radius) !important;
        min-height: 46px !important;
    }

    .stSelectbox label,
    .stMultiSelect label,
    .stTextInput label,
    .stTextArea label {
        color: var(--pine) !important;
        font-weight: 700 !important;
        font-size: 0.9rem !important;
    }

    .stButton button {
        background: var(--pine) !important;
        color: #fffdf6 !important;
        border: 1px solid rgba(255, 255, 255, 0.18) !important;
        padding: 0.62rem 1.1rem !important;
        border-radius: var(--radius) !important;
        font-weight: 800 !important;
        box-shadow: var(--soft-shadow) !important;
        transition: transform 160ms ease, box-shadow 160ms ease, background 160ms ease !important;
    }

    .stButton button:hover {
        background: #1d5448 !important;
        transform: translateY(-1px);
        box-shadow: 0 12px 28px rgba(21, 63, 54, 0.18) !important;
    }

    [data-testid="stMetric"] {
        background: var(--panel-strong);
        border: 1px solid var(--line);
        border-radius: var(--radius);
        padding: 1rem 1.1rem;
        box-shadow: var(--soft-shadow);
    }

    [data-testid="stMetricLabel"] {
        color: var(--muted);
        font-size: 0.78rem;
        font-weight: 700;
    }

    [data-testid="stMetricValue"] {
        color: var(--pine);
        font-family: 'Fraunces', 'Noto Sans SC', serif;
        font-size: 2rem;
        font-weight: 800;
    }

    [data-testid="stFileUploader"] {
        background: var(--panel);
        border: 1px dashed rgba(21, 63, 54, 0.34);
        border-radius: var(--radius);
        padding: 1rem;
        box-shadow: var(--soft-shadow);
    }

    .streamlit-expanderHeader {
        background: rgba(255, 253, 246, 0.9) !important;
        border: 1px solid var(--line) !important;
        border-radius: var(--radius) !important;
        color: var(--pine) !important;
        font-weight: 800 !important;
    }

    .stTabs [data-baseweb="tab-list"] {
        gap: 0.5rem !important;
        padding: 0.4rem !important;
        background: rgba(21, 63, 54, 0.08) !important;
        border: 1px solid rgba(21, 63, 54, 0.12) !important;
        border-radius: var(--radius) !important;
        margin-bottom: 1.25rem !important;
    }

    .stTabs [data-baseweb="tab"] {
        height: 44px !important;
        padding: 0 1rem !important;
        border-radius: 6px !important;
        color: var(--pine) !important;
        font-weight: 800 !important;
        background: transparent !important;
    }

    .stTabs [aria-selected="true"] {
        background: var(--panel-strong) !important;
        color: var(--coral) !important;
        box-shadow: 0 6px 18px rgba(23, 32, 27, 0.08) !important;
    }

    .workspace-hero {
        position: relative;
        overflow: hidden;
        border-radius: var(--radius);
        padding: 2.1rem;
        margin-bottom: 1.25rem;
        color: #fffdf6;
        background:
            linear-gradient(120deg, rgba(21, 63, 54, 0.96), rgba(38, 82, 62, 0.92)),
            repeating-linear-gradient(135deg, transparent 0, transparent 16px, rgba(255,255,255,0.06) 16px, rgba(255,255,255,0.06) 17px);
        box-shadow: var(--shadow);
    }

    .workspace-hero::after {
        content: "";
        position: absolute;
        right: -8rem;
        top: -8rem;
        width: 22rem;
        height: 22rem;
        border: 1px solid rgba(255, 253, 246, 0.22);
        transform: rotate(18deg);
    }

    .hero-kicker {
        color: rgba(255, 253, 246, 0.76);
        font-size: 0.82rem;
        font-weight: 800;
        letter-spacing: 0.08em;
        text-transform: uppercase;
        margin-bottom: 0.55rem;
    }

    .hero-title {
        font-family: 'Fraunces', 'Noto Sans SC', serif;
        font-size: clamp(2.2rem, 5vw, 4.7rem);
        line-height: 0.95;
        font-weight: 800;
        margin: 0;
        max-width: 760px;
    }

    .hero-copy {
        max-width: 680px;
        margin: 1rem 0 0;
        color: rgba(255, 253, 246, 0.82);
        font-size: 1rem;
        line-height: 1.7;
    }

    .hero-meta {
        display: flex;
        gap: 0.65rem;
        flex-wrap: wrap;
        margin-top: 1.4rem;
    }

    .hero-pill {
        border: 1px solid rgba(255, 253, 246, 0.22);
        background: rgba(255, 253, 246, 0.10);
        border-radius: 999px;
        padding: 0.45rem 0.75rem;
        font-size: 0.82rem;
        font-weight: 800;
    }

    .steps-grid {
        display: grid;
        grid-template-columns: repeat(3, minmax(0, 1fr));
        gap: 0.85rem;
        margin-bottom: 1.25rem;
    }

    .step-card,
    .notice-card,
    .control-card,
    .comment-card {
        background: var(--panel);
        border: 1px solid var(--line);
        border-radius: var(--radius);
        box-shadow: var(--soft-shadow);
    }

    .step-card {
        padding: 1rem;
        border-top: 4px solid var(--amber);
    }

    .step-card strong {
        display: block;
        color: var(--pine);
        margin-bottom: 0.28rem;
        font-size: 0.98rem;
    }

    .step-card span {
        color: var(--muted);
        font-size: 0.88rem;
        line-height: 1.55;
    }

    .notice-card {
        padding: 1.1rem 1.2rem;
        margin-bottom: 1.25rem;
        border-left: 5px solid var(--sky);
    }

    .notice-card h4 {
        margin: 0 0 0.35rem;
        color: var(--pine);
        font-size: 1.02rem;
    }

    .notice-card p {
        color: var(--muted);
        margin: 0;
        line-height: 1.65;
        font-size: 0.92rem;
    }

    .designer-mark {
        color: var(--muted);
        text-align: right;
        font-size: 0.82rem;
        font-weight: 700;
        margin: 0.5rem 0 1.1rem;
    }

    .filter-box {
        background: rgba(255, 253, 246, 0.94);
        border: 1px solid var(--line);
        border-radius: var(--radius);
        box-shadow: var(--soft-shadow);
        padding: 1rem;
        margin: 1rem 0 1.25rem;
    }

    .filter-title {
        color: var(--pine);
        font-weight: 900;
        margin-bottom: 0.75rem;
        padding-bottom: 0.65rem;
        border-bottom: 1px solid var(--line);
    }

    .filter-item,
    .control-card {
        padding: 1rem;
    }

    .file-summary {
        background: rgba(84, 122, 68, 0.08);
        border: 1px solid rgba(84, 122, 68, 0.16);
        border-radius: var(--radius);
        padding: 0.85rem;
        margin-bottom: 1rem;
        color: var(--pine);
        line-height: 1.55;
    }

    .comment-card {
        padding: 1rem;
        margin-bottom: 0.85rem;
        min-height: 120px;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        transition: transform 160ms ease, box-shadow 160ms ease;
    }

    .comment-card:hover {
        transform: translateY(-2px);
        box-shadow: var(--shadow);
    }

    .comment-text {
        color: var(--ink);
        line-height: 1.7;
        font-size: 0.94rem;
    }

    .comment-source {
        margin-top: 0.75rem;
        color: var(--muted);
        font-size: 0.78rem;
        text-align: right;
        font-weight: 700;
    }

    .highlight {
        color: var(--coral);
        background: rgba(216, 139, 45, 0.16);
        border: 1px solid rgba(216, 139, 45, 0.22);
        border-radius: 4px;
        padding: 0 0.2rem;
        font-weight: 900;
    }

    @media (max-width: 820px) {
        .main .block-container {
            padding: 1rem 0.8rem 3rem;
        }

        .workspace-hero {
            padding: 1.35rem;
        }

        .steps-grid {
            grid-template-columns: 1fr;
        }
    }
</style>
""", unsafe_allow_html=True)

# 工具函数
def ensure_font():
    """确保字体文件存在并返回字体路径"""
    font_dir = Path(__file__).parent / 'fonts'
    font_path = font_dir / 'simhei.ttf'
    
    if font_path.exists():
        return str(font_path)
    
    font_dir.mkdir(exist_ok=True)
    
    # 检查系统字体路径
    system_fonts = [
        Path('C:/Windows/Fonts/simhei.ttf'),
        Path('/usr/share/fonts/truetype/simhei.ttf'),
        Path('/System/Library/Fonts/simhei.ttf')
    ]
    
    for system_font in system_fonts:
        if system_font.exists():
            try:
                shutil.copy(system_font, font_path)
                return str(font_path)
            except Exception:
                pass
    
    # 从网络下载字体
    try:
        font_url = "https://cdn.jsdelivr.net/gh/googlefonts/noto-cjk@main/Sans/OTF/SimplifiedChinese/NotoSansCJKsc-Regular.otf"
        response = requests.get(font_url)
        response.raise_for_status()
        
        with open(font_path, 'wb') as f:
            f.write(response.content)
        
        return str(font_path)
        
    except Exception as e:
        st.error(f"获取字体文件失败: {str(e)}")
        return None

def highlight_words(text, selected_word, source=None):
    """高亮显示选中的关键词，并添加来源标注"""
    safe_text = html.escape(str(text))
    safe_word = html.escape(str(selected_word))
    highlighted_text = safe_text.replace(
        safe_word,
        f'<span class="highlight">{safe_word}</span>'
    )
    if source:
        highlighted_text += f'<div class="comment-source">来源: {html.escape(str(source))}</div>'
    return highlighted_text

def render_comment_card(comment, selected_word, source):
    """渲染统一风格的评论卡片"""
    st.markdown(
        f"""
        <div class="comment-card">
            <div class="comment-text">{highlight_words(comment, selected_word)}</div>
            <div class="comment-source">来源: {html.escape(str(source))}</div>
        </div>
        """,
        unsafe_allow_html=True
    )

def style_bar_chart(fig, color='#153f36'):
    """统一 Plotly 柱状图视觉语言"""
    fig.update_layout(
        plot_bgcolor='rgba(255,253,246,0.7)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(color='#17201b', family='Noto Sans SC'),
        showlegend=False
    )
    fig.update_traces(
        marker_color=color,
        marker_line_color='rgba(23,32,27,0.28)',
        marker_line_width=1,
        opacity=0.9
    )
    fig.update_yaxes(showgrid=True, gridwidth=1, gridcolor='rgba(23,32,27,0.10)')
    fig.update_xaxes(showgrid=False)
    return fig

def get_most_complete_comment(comments):
    """从相似评论中选择最完整的一条"""
    normalized_comments = {}
    for comment in comments:
        normalized = ' '.join(comment.split())
        if normalized in normalized_comments:
            if len(comment) > len(normalized_comments[normalized]):
                normalized_comments[normalized] = comment
        else:
            normalized_comments[normalized] = comment
    return list(normalized_comments.values())

# UI组件函数
def render_header():
    """渲染页面标题和说明"""
    st.markdown(f"""
    <div class="designer-mark">v{VERSION} · Designed with 🐑 @小羊</div>
    <section class="workspace-hero">
        <div class="hero-kicker">Comment Intelligence Workspace</div>
        <h1 class="hero-title">Excel 评论分析工具</h1>
        <p class="hero-copy">
            面向运营复盘的评论分析工作台。上传评论表后，快速合并多文件、定位差评线索、抽取建议关键词，并把高频反馈变成可读的图表和词云。
        </p>
        <div class="hero-meta">
            <span class="hero-pill">多文件合并</span>
            <span class="hero-pill">差评定位</span>
            <span class="hero-pill">建议词识别</span>
            <span class="hero-pill">词云可视化</span>
        </div>
    </section>
    <div class="steps-grid">
        <div class="step-card">
            <strong>上传 Excel</strong>
            <span>支持一次选择多个文件，并自动保留来源，方便横向比较。</span>
        </div>
        <div class="step-card">
            <strong>筛选评论</strong>
            <span>按文件、关键词、建议类型或负面情绪快速收窄样本。</span>
        </div>
        <div class="step-card">
            <strong>查看洞察</strong>
            <span>在总体、差评、建议、负面四个视角里追踪真实反馈。</span>
        </div>
    </div>
    <div class="notice-card">
        <h4>系统更新公告 v2.2.0</h4>
        <p>本版重点优化图表阅读、筛选控制和评论卡片层级，让高频词、文件来源和具体原文之间的跳转更直接。</p>
    </div>
    """, unsafe_allow_html=True)

# 主函数
def main():
    logger.info(f"Python 版本: {sys.version}")
    
    # 渲染页面主要内容
    render_header()
    
    uploaded_files = st.file_uploader(
        "选择Excel文件上传（可多选）",
        type=['xlsx'],
        accept_multiple_files=True,
        help="请上传包含评论数据的Excel文件（.xlsx格式）"
    )
    
    if uploaded_files:
        # 创建筛选框
        st.markdown("""
        <div class="filter-box">
            <div class="filter-title">筛选条件</div>
        """, unsafe_allow_html=True)
        
        # 使用列布局来组织筛选条件
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.markdown('<div class="filter-item">', unsafe_allow_html=True)
            selected_files = st.multiselect(
                "📁 选择文件",
                options=uploaded_files,
                format_func=lambda x: x.name,
                help="可以选择多个文件进行合并分析"
            )
            st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('</div>', unsafe_allow_html=True)
        
        if selected_files:
            try:
                # 初始化筛选条件变量
                filter_keyword = ""
                comment_type = ["全部评论"]
                
                # 创建两列布局：左侧为控制面板，右侧为分析结果
                control_col, result_col = st.columns([1, 2])
                
                with control_col:
                    # 控制面板容器
                    st.markdown("""
                    <div class="control-card">
                        <strong>控制面板</strong>
                        <p style="margin: 0.35rem 0 0; color: var(--muted); font-size: 0.88rem;">调整筛选口径并管理停用词。</p>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    # 文件信息
                    st.markdown(f"""
                    <div class="file-summary">
                        分析文件<br><strong>{html.escape(', '.join(f.name for f in selected_files))}</strong>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    # 筛选条件（折叠面板）
                    with st.expander("🔍 筛选条件", expanded=True):
                        filter_keyword = st.text_input("关键词筛选", help="输入关键词筛选评论")
                        comment_type = st.multiselect(
                            "评论类型",
                            options=["全部评论", "建议评论", "负面评论"],
                            default=["全部评论"]
                        )
                    
                    # 词汇管理（折叠面板）
                    with st.expander("⚙️ 词汇管理", expanded=False):
                        # 添加停用词
                        new_stop_words = st.text_area(
                            "添加需要过滤的词（每行一个）",
                            help="输入需要从分析中排除的词，每行输入一个词",
                            key="stop_word_input",
                            height=100
                        )
                        if st.button("添加到停用词", key="add_stop_word"):
                            if new_stop_words:
                                words_to_add = {word.strip() for word in new_stop_words.split('\n') if word.strip()}
                                st.session_state.user_stop_words.update(words_to_add)
                                if words_to_add:
                                    st.success(f"已添加 {len(words_to_add)} 个停用词")
                        
                        # 管理停用词
                        if st.session_state.user_stop_words:
                            selected_words = st.multiselect(
                                "当前停用词（可多选删除）",
                                options=sorted(list(st.session_state.user_stop_words)),
                                default=[],
                                key="stop_word_select"
                            )
                            c1, c2 = st.columns(2)
                            with c1:
                                if st.button("删除选中", key="delete_stop_word"):
                                    for word in selected_words:
                                        st.session_state.user_stop_words.remove(word)
                            with c2:
                                if st.button("清空全部", key="clear_stop_words"):
                                    st.session_state.user_stop_words.clear()
                    
                    # 数据处理
                    all_dfs = []
                    for file in selected_files:
                        df = pd.read_excel(file, header=5)
                        df['数据来源'] = file.name
                        all_dfs.append(df)
                    
                    # 合并所有数据
                    df = pd.concat(all_dfs, ignore_index=True)
                    
                    # 查找路线列和评分列
                    route_col = None
                    score_col = None
                    for col in df.columns:
                        if '路线名称' in str(col) or '产品名称' in str(col):
                            route_col = col
                        if '总安排打分' in str(col):
                            score_col = col
                    
                    if score_col is None:
                        st.error('在上传的文件中未找到"总安排打分"列')
                        return
                    
                    # 应用筛选条件
                    mask = pd.Series(True, index=df.index)
                    
                    if filter_keyword:
                        mask = mask & df.iloc[:, 0].str.contains(filter_keyword, na=False)
                    
                    if "全部评论" not in comment_type:
                        if "建议评论" in comment_type:
                            suggestion_mask = df.iloc[:, 0].str.contains('|'.join(SUGGESTION_WORDS), na=False)
                            mask = mask & suggestion_mask
                        if "负面评论" in comment_type:
                            negative_mask = df.iloc[:, 0].str.contains('|'.join(NEGATIVE_WORDS), na=False)
                            mask = mask & negative_mask
                    
                    filtered_df = df[mask]
                    comments = filtered_df.iloc[:, 0]
                    scores = filtered_df[score_col]
                    
                    # 计算统计数据
                    total_comments = len(comments)
                    low_score_comments = sum(1 for score in scores if score <= 3)
                    
                    # 数据统计
                    st.metric("总评论数", total_comments)
                    st.metric("差评数", low_score_comments)
                    
                    # 在进入标签页之前，先初始化所有计数器
                    # 词频统计
                    word_freq = Counter()
                    word_freq_low = Counter()
                    word_comments = {}
                    word_comments_low = {}
                    suggestion_freq = Counter()
                    negative_freq = Counter()
                    suggestion_comments = {}
                    negative_comments = {}
                    
                    # 处理评论数据
                    for comment, source, score in zip(comments, filtered_df['数据来源'], scores):
                        if pd.isna(comment) or pd.isna(score):
                            continue
                        
                        comment = str(comment).strip()
                        if not comment:
                            continue
                        
                        words = jieba.cut(comment)
                        comment_words = set()
                        
                        for word in words:
                            word = word.strip()
                            if (len(word) > 1 and
                                word not in STOP_WORDS and
                                word not in st.session_state.user_stop_words and
                                not word.isdigit()):
                                
                                # 总体词频统计
                                word_freq[word] += 1
                                if word not in word_comments:
                                    word_comments[word] = set()
                                word_comments[word].add((comment, source))
                                
                                # 差评词频统计
                                if score <= 3:
                                    word_freq_low[word] += 1
                                    if word not in word_comments_low:
                                        word_comments_low[word] = set()
                                    word_comments_low[word].add((comment, source))
                                
                                # 建议词统计
                                if word in SUGGESTION_WORDS:
                                    suggestion_freq[word] += 1
                                    if word not in suggestion_comments:
                                        suggestion_comments[word] = set()
                                    suggestion_comments[word].add((comment, source))
                                
                                # 负面词统计
                                if word in NEGATIVE_WORDS:
                                    negative_freq[word] += 1
                                    if word not in negative_comments:
                                        negative_comments[word] = set()
                                    negative_comments[word].add((comment, source))
                    
                    # 然后在标签页中使用这些
                    with result_col:
                        # 分析结果标签页
                        tab1, tab2, tab3, tab4 = st.tabs([
                            "📈 总体分析", "📉 差评分析", 
                            "💡 建议分析", "😟 负面分析"
                        ])

                        # 总体分析
                        with tab1:
                            if word_freq:
                                # 创建一个容器来包裹所有内容
                                with st.container():
                                    # 第一行：数据来源和词云图
                                    col1, col2 = st.columns(2)
                                    with col1:
                                        st.subheader("📊 数据来源分布")
                                        source_counts = filtered_df['数据来源'].value_counts()
                                        fig_source = px.pie(
                                            values=source_counts.values,
                                            names=source_counts.index,
                                            height=300,  # 固定高度
                                            width=None,  # 自动应宽度
                                        )
                                        # 调整饼图布局
                                        fig_source.update_layout(
                                            margin=dict(l=20, r=20, t=20, b=20),  # 减小边距
                                            showlegend=True,  # 显示图例
                                            legend=dict(
                                                orientation="h",  # 水平图例
                                                yanchor="bottom",
                                                y=1.02,  # 图例位置
                                                xanchor="right",
                                                x=1
                                            ),
                                            # 调整饼图大小
                                            autosize=True,  # 自动调整���小
                                            height=300,  # 固定高度
                                        )
                                        # 调整饼图样式
                                        fig_source.update_traces(
                                            textposition='inside',  # 文字位置
                                            textinfo='percent+label',  # 显示百分比和标签
                                            hole=0.3,  # 添加环形效果
                                            pull=[0.05] * len(source_counts),  # 轻微分离扇形
                                            marker=dict(
                                                colors=CHART_COLORS,
                                                line=dict(color='white', width=2)  # 添加白色边框
                                            )
                                        )
                                        st.plotly_chart(fig_source, use_container_width=True, config={
                                            'displayModeBar': False  # 隐藏plotly工具栏
                                        })
                                    
                                    with col2:
                                        st.subheader("☁️ 词云图")
                                        wc = WordCloud(
                                            font_path=ensure_font(),
                                            width=400,
                                            height=300,
                                            background_color='white'
                                        )
                                        wc.generate_from_frequencies(word_freq)
                                        st.image(wc.to_array())
                                    
                                    # 第二行：词频统计
                                    st.subheader("📈 词频统计")
                                    top_words = dict(sorted(word_freq.items(), key=lambda x: x[1], reverse=True)[:20])
                                    fig = px.bar(
                                        x=list(top_words.keys()),
                                        y=list(top_words.values()),
                                        labels={'x': '关键词', 'y': '出现次数'},
                                        height=400  # 固定高度
                                    )
                                    fig.update_layout(
                                        margin=dict(l=20, r=20, t=20, b=80),  # 增加底部边距，为倾斜的标签��出空间
                                        xaxis_tickangle=-45,  # 标签倾斜角度
                                        xaxis=dict(
                                            tickmode='array',
                                            ticktext=list(top_words.keys()),
                                            tickvals=list(range(len(top_words))),
                                            tickfont=dict(size=11)  # 调整字体大小
                                        ),
                                        yaxis=dict(
                                            title=dict(
                                                text='出现次数',
                                                font=dict(size=12)
                                            ),
                                            tickfont=dict(size=11)
                                        ),
                                        bargap=0.2,  # 调整柱子之间的间距
                                        plot_bgcolor='white',  # 设置背景色为白色
                                        showlegend=False
                                    )
                                    style_bar_chart(fig, CHART_COLORS[0])
                                    st.plotly_chart(fig, use_container_width=True, config={
                                        'displayModeBar': False  # 隐藏plotly工具栏
                                    })
                                    
                                    # 第三行：评论详情
                                    st.subheader("💬 评论详情")
                                    
                                    selected_word = st.selectbox(
                                        "选择关键词查看相关评论",
                                        options=list(top_words.keys())
                                    )
                                    
                                    if selected_word:
                                        relevant_comments = word_comments.get(selected_word, set())
                                        unique_comments = []
                                        seen_texts = set()
                                        for comment, source in relevant_comments:
                                            normalized = ' '.join(comment.split())
                                            if normalized not in seen_texts:
                                                seen_texts.add(normalized)
                                                unique_comments.append((comment, source))
                                        
                                        # 使用网格布局显示评论
                                        cols = st.columns(2)
                                        for i, (comment, source) in enumerate(unique_comments):
                                            with cols[i % 2]:
                                                render_comment_card(comment, selected_word, source)
                            else:
                                st.info("没有找到有效的评论数据")

                        # 差评分析
                        with tab2:
                            if word_freq_low:
                                # 创建一个容器来包裹所有内容
                                with st.container():
                                    # 第一行：词云图和词频统计
                                    viz_col1, viz_col2 = st.columns(2)
                                    
                                    with viz_col1:
                                        st.subheader("☁️ 差评词云图")
                                        wc = WordCloud(
                                            font_path=ensure_font(),
                                            width=400,
                                            height=300,
                                            background_color='white'
                                        )
                                        wc.generate_from_frequencies(word_freq_low)
                                        st.image(wc.to_array())
                                    
                                    with viz_col2:
                                        st.subheader("📊 差评词频统计")
                                        top_words_low = dict(sorted(word_freq_low.items(), key=lambda x: x[1], reverse=True)[:20])
                                        fig = px.bar(
                                            x=list(top_words_low.keys()),
                                            y=list(top_words_low.values()),
                                            labels={'x': '关键词', 'y': '出现次数'},
                                            height=300  # 固定高度
                                        )
                                        fig.update_layout(
                                            margin=dict(l=20, r=20, t=20, b=20),
                                            xaxis_tickangle=-45
                                        )
                                        style_bar_chart(fig, CHART_COLORS[3])
                                        st.plotly_chart(fig, use_container_width=True)
                                    
                                    # 第二行：差评详情
                                    st.subheader("💬 差评详情")
                                    
                                    selected_word = st.selectbox(
                                        "选择关键词查看相关差评",
                                        options=list(top_words_low.keys()),
                                        key="low_score_select"
                                    )
                                    
                                    if selected_word:
                                        # 使用网格布局显示评论
                                        cols = st.columns(2)
                                        unique_comments = []
                                        seen_texts = set()
                                        for comment, source in word_comments_low.get(selected_word, set()):
                                            normalized = ' '.join(comment.split())
                                            if normalized not in seen_texts:
                                                seen_texts.add(normalized)
                                                unique_comments.append((comment, source))
                                        
                                        for i, (comment, source) in enumerate(unique_comments):
                                            with cols[i % 2]:
                                                render_comment_card(comment, selected_word, source)
                            else:
                                st.info("没有找到差评数据")

                        # 建议分析
                        with tab3:
                            if suggestion_freq:
                                with st.container():
                                    # 第一行：建议词统计
                                    st.subheader("📊 建议关键词统计")
                                    top_suggestions = dict(sorted(suggestion_freq.items(), key=lambda x: x[1], reverse=True)[:20])
                                    fig = px.bar(
                                        x=list(top_suggestions.keys()),
                                        y=list(top_suggestions.values()),
                                        labels={'x': '建议关键词', 'y': '出现次数'},
                                        height=300
                                    )
                                    fig.update_layout(
                                        margin=dict(l=20, r=20, t=20, b=20),
                                        xaxis_tickangle=-45
                                    )
                                    style_bar_chart(fig, CHART_COLORS[1])
                                    st.plotly_chart(fig, use_container_width=True)
                                    
                                    # 第二行：建议详情
                                    st.subheader("💡 建议详情")
                                    
                                    selected_word = st.selectbox(
                                        "选择关键词查看相关建议",
                                        options=list(top_suggestions.keys()),
                                        key="suggestion_select"
                                    )
                                    
                                    if selected_word:
                                        cols = st.columns(2)
                                        unique_comments = []
                                        seen_texts = set()
                                        for comment, source in suggestion_comments.get(selected_word, set()):
                                            normalized = ' '.join(comment.split())
                                            if normalized not in seen_texts:
                                                seen_texts.add(normalized)
                                                unique_comments.append((comment, source))
                                        
                                        for i, (comment, source) in enumerate(unique_comments):
                                            with cols[i % 2]:
                                                render_comment_card(comment, selected_word, source)
                            else:
                                st.info("没有找到建议相关的评论")

                        # 负面分析
                        with tab4:
                            if negative_freq:
                                with st.container():
                                    # 第一行：词云图和词频统计
                                    viz_col1, viz_col2 = st.columns(2)
                                    
                                    with viz_col1:
                                        st.subheader("☁️ 负面情绪词云图")
                                        wc = WordCloud(
                                            font_path=ensure_font(),
                                            width=400,
                                            height=300,
                                            background_color='white'
                                        )
                                        wc.generate_from_frequencies(negative_freq)
                                        st.image(wc.to_array())
                                    
                                    with viz_col2:
                                        st.subheader("📊 负面情绪词统计")
                                        top_negative = dict(sorted(negative_freq.items(), key=lambda x: x[1], reverse=True)[:20])
                                        fig = px.bar(
                                            x=list(top_negative.keys()),
                                            y=list(top_negative.values()),
                                            labels={'x': '负面情绪词', 'y': '出现次数'},
                                            height=300  # 固定高度
                                        )
                                        fig.update_layout(
                                            margin=dict(l=20, r=20, t=20, b=20),
                                            xaxis_tickangle=-45
                                        )
                                        style_bar_chart(fig, CHART_COLORS[3])
                                        st.plotly_chart(fig, use_container_width=True)
                                    
                                    # 第二行：负面评论详情
                                    st.subheader("😟 负面评论详情")
                                    
                                    selected_word = st.selectbox(
                                        "选择关键词查看相关负面评论",
                                        options=list(top_negative.keys()),
                                        key="negative_select"
                                    )
                                    
                                    if selected_word:
                                        cols = st.columns(2)
                                        unique_comments = []
                                        seen_texts = set()
                                        for comment, source in negative_comments.get(selected_word, set()):
                                            normalized = ' '.join(comment.split())
                                            if normalized not in seen_texts:
                                                seen_texts.add(normalized)
                                                unique_comments.append((comment, source))
                                        
                                        for i, (comment, source) in enumerate(unique_comments):
                                            with cols[i % 2]:
                                                render_comment_card(comment, selected_word, source)
                            else:
                                st.info("没有找到负面情绪相关的评论")

            except Exception as e:
                st.error(f"处理文件时出错: {str(e)}")

if __name__ == "__main__":
    main()
