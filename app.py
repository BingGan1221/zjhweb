# 标准库导入
import sys
import logging
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

# 在文件开头，USER_STOP_WORDS 定义后添加
if 'user_stop_words' not in st.session_state:
    st.session_state.user_stop_words = set()

# 在文件开头添加自定义样式
st.markdown("""
<style>
    /* 输入框样式 */
    .stTextInput input {
        border: 2px solid #2E86C1 !important;
        border-radius: 8px !important;
        padding: 10px 15px !important;
        background-color: white !important;
        color: #333 !important;
        font-size: 0.9rem !important;
    }
    .stTextInput input:focus {
        box-shadow: 0 0 0 2px rgba(46,134,193,0.2) !important;
        border-color: #2E86C1 !important;
    }
    
    /* 文本区域样式 */
    .stTextArea textarea {
        border: 2px solid #2E86C1 !important;
        border-radius: 8px !important;
        padding: 10px 15px !important;
        background-color: white !important;
        color: #333 !important;
        font-size: 0.9rem !important;
    }
    .stTextArea textarea:focus {
        box-shadow: 0 0 0 2px rgba(46,134,193,0.2) !important;
        border-color: #2E86C1 !important;
    }
    
    /* 多选框样式 */
    .stMultiSelect {
        background-color: white !important;
        border: 2px solid #2E86C1 !important;
        border-radius: 8px !important;
        padding: 2px !important;
    }
    .stMultiSelect:hover {
        border-color: #2E86C1 !important;
    }
    
    /* 按钮样式 */
    .stButton button {
        background-color: #2E86C1 !important;
        color: white !important;
        border: none !important;
        padding: 0.5rem 1rem !important;
        border-radius: 8px !important;
        font-weight: 500 !important;
        transition: all 0.3s ease !important;
    }
    .stButton button:hover {
        background-color: #2874A6 !important;
        box-shadow: 0 2px 6px rgba(46,134,193,0.3) !important;
    }
    
    /* 标签样式 */
    .stSelectbox label, .stMultiSelect label, .stTextInput label {
        color: #2E86C1 !important;
        font-weight: 500 !important;
        font-size: 0.95rem !important;
        margin-bottom: 0.3rem !important;
    }
    
    /* 下拉框样式 */
    .stSelectbox select {
        border: 2px solid #2E86C1 !important;
        border-radius: 8px !important;
        padding: 10px 15px !important;
        background-color: white !important;
        color: #333 !important;
        font-size: 0.9rem !important;
    }
    
    /* 帮助文本样式 */
    .stMarkdown div.help-text {
        color: #666 !important;
        font-size: 0.85rem !important;
        margin-top: 0.2rem !important;
        font-style: italic !important;
    }
</style>
""", unsafe_allow_html=True)

# 工具函数
def ensure_font():
    """确保字体文件存在并返回字体路径"""
    font_dir = Path('fonts')
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
            shutil.copy(system_font, font_path)
            return str(font_path)
    
    # 从网络下载字体
    try:
        font_url = "https://github.com/microsoft/Windows-Font/raw/master/SimHei.ttf"
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
    highlighted_text = text.replace(
        selected_word, 
        f'<span style="color: red; font-weight: bold;">{selected_word}</span>'
    )
    if source:
        highlighted_text += f'<div style="text-align: right; color: #666666; font-size: 0.8em; margin-top: 0.3em;">来源: {source}</div>'
    return highlighted_text

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
    # 添加全局样式和羊的背景
    st.markdown("""
    <style>
        .stApp {
            background: linear-gradient(to bottom, #F8F9F9, #FFFFFF);
        }
        
        /* 羊背景样式 */
        .sheep-bg {
            position: fixed;
            font-size: var(--size);
            opacity: 0.05;
            transform: rotate(var(--rotate));
            z-index: -999;
            pointer-events: none;
        }
        
        /* 标签页按钮样式 */
        .stTabs [data-baseweb="tab-list"] {
            gap: 30px;
            padding: 1rem;
            background-color: #F5F7FA;
            border-radius: 20px;
            box-shadow: 0 2px 12px rgba(0,0,0,0.04);
            margin-bottom: 2rem;
        }
        
        .stTabs [data-baseweb="tab"] {
            height: 60px;
            padding: 0 30px;
            background-color: white;
            border-radius: 12px;
            color: #1976D2;
            font-weight: 600;
            font-size: 1.1rem;
            border: none;
            box-shadow: 0 2px 8px rgba(0,0,0,0.05);
            transition: all 0.3s ease;
        }
        
        .stTabs [data-baseweb="tab"]:hover {
            background-color: #E3F2FD;
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
        }
        
        .stTabs [aria-selected="true"] {
            background: linear-gradient(135deg, #1976D2, #2196F3) !important;
            color: white !important;
            transform: translateY(-2px);
            box-shadow: 0 4px 15px rgba(25,118,210,0.3) !important;
        }
        
        /* 确保内容区域在最上层 */
        .main .block-container {
            position: relative;
            z-index: 99;
            background: rgba(255, 255, 255, 0.95);
            padding: 1rem;
            border-radius: 10px;
        }
        
        /* 确保所有组件在最上层 */
        .stButton, .stSelectbox, .stFileUploader, .stTabs, 
        .stMarkdown, .stMetric, .element-container {
            position: relative;
            z-index: 100;
        }
    </style>
    """, unsafe_allow_html=True)
    
    # 生成羊的背景
    sheep_styles = []
    rows, cols = 6, 8
    for row in range(rows):
        for col in range(cols):
            left = (col * 100 / cols) + np.random.randint(-15, 15)
            top = (row * 100 / rows) + np.random.randint(-15, 15)
            size = np.random.randint(20, 40)
            rotate = np.random.randint(-45, 45)
            
            sheep_styles.append(
                f"left: {left}%; top: {top}%; "
                f"font-size: {size}px; "
                f"transform: rotate({rotate}deg);"
            )
    
    sheep_divs = '\n'.join([
        f'<div class="sheep-bg" style="{style}">🐑</div>'
        for style in sheep_styles
    ])
    
    # 添加背景羊
    st.markdown(f"""
    <div style="position: fixed; width: 100%; height: 100%; z-index: -999; pointer-events: none;">
        {sheep_divs}
    </div>
    """, unsafe_allow_html=True)
    
    # 添加版本信息到标题区域
    st.markdown("""
    <div style='text-align: right; margin-bottom: 1rem;'>
        <span style='background: rgba(46, 134, 193, 0.1); 
                     padding: 0.3rem 0.8rem; 
                     border-radius: 15px; 
                     font-size: 0.8rem; 
                     color: #2E86C1;'>
             v""" + VERSION + """
        </span>
    </div>
    """, unsafe_allow_html=True)
    
    # 渲染标题
    st.markdown("""
    <div style='background: linear-gradient(135deg, #2E86C1, #3498DB); 
         padding: 2rem; border-radius: 15px; margin-bottom: 2rem; 
         box-shadow: 0 4px 15px rgba(0,0,0,0.1); text-align: center;'>
        <h1 style='color: white; margin-bottom: 0.5rem; font-size: 2.5rem;'>
            📊 Excel评论分析工具
        </h1>
        <p style='color: rgba(255,255,255,0.9); font-size: 1.1rem;'>
            快速分析和可视化您的评论数据
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    # 渲染使用说明
    st.markdown("""
    <div style='background-color: #EBF5FB; padding: 1rem; border-radius: 8px; 
         margin-bottom: 1.5rem; box-shadow: 0 1px 4px rgba(0,0,0,0.05);'>
        <div style='display: flex; gap: 1rem; flex-wrap: wrap;'>
            <div style='background: white; padding: 0.8rem 1rem; border-radius: 6px; 
                 border-left: 3px solid #2E86C1; flex: 1; min-width: 200px;'>
                <p style='margin: 0; color: #2E86C1;'>📤 上传Excel文件</p>
            </div>
            <div style='background: white; padding: 0.8rem 1rem; border-radius: 6px; 
                 border-left: 3px solid #2E86C1; flex: 1; min-width: 200px;'>
                <p style='margin: 0; color: #2E86C1;'>🔍 选择筛选条件</p>
            </div>
            <div style='background: white; padding: 0.8rem 1rem; border-radius: 6px; 
                 border-left: 3px solid #2E86C1; flex: 1; min-width: 200px;'>
                <p style='margin: 0; color: #2E86C1;'>📊 查看分析结果</p>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # 在 render_header 函��中，使用说添加系统公告
    image_url = "https://raw.githubusercontent.com/你的用户名/仓库名/main/feedback.png"

    st.markdown("""
    <div style='background: linear-gradient(135deg, #EBF5FB, #D6EAF8); padding: 1rem; border-radius: 8px; 
         margin-bottom: 1.5rem; border: 1px solid #2E86C1;'>
        <h4 style='color: #2E86C1; margin-bottom: 0.8rem; font-size: 1rem; display: flex; align-items: center;'>
            <span style='margin-right: 0.5rem;'>📢</span> 系统更新公告 v2.1.0
        </h4>
        <div style='color: #2E86C1; font-size: 0.9rem; line-height: 1.5;'>
            <p style='margin: 0 0 0.5rem 0;'>更新内容：</p>
            <ul style='margin: 0 0 1rem 1.5rem;'>
                <li>✨ <strong>新增多文件合并分析功能</strong></li>
                <li>📊 新增数据来源分布统计</li>
                <li>🔍 优化多文件筛选体验</li>
                <li>📈 改进数据分析展示效果</li>
                <li>🐛 修复已知问题，提升稳定性</li>
            </ul>
            <div style='background: white; padding: 0.8rem; border-radius: 6px; margin-top: 0.5rem;'>
                <p style='margin: 0; color: #2E86C1;'>
                    💡 使用提示：现在可以同时选择多个文件进行合并分析，系统会自动整合所有数据示分布情况。
                </p>
                <p style='margin: 0.5rem 0 0 0; color: #666; font-size: 0.85rem; font-style: italic; 
                    padding: 0.5rem; background: #F8F9FA; border-radius: 4px;'>
                    "希望能够同时分析多个文件的数据，这样可以更好地进行整体分析"
                </p>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # 文件上传区域美化
    st.markdown("""
    <div style='background-color: #F8F9FA; padding: 2rem; border-radius: 10px; border: 2px dashed #1E88E5; text-align: center; margin-bottom: 2rem;'>
    """, unsafe_allow_html=True)
    
    # 添加设计者标识
    st.markdown("<br><br>", unsafe_allow_html=True)
    st.markdown("""
    <div style='text-align: center; padding: 2rem; margin-top: 2rem; background: linear-gradient(to right, #F8F9FA, #FFFFFF, #F8F9FA);'>
        <p style='color: #757575; font-size: 0.9rem; margin: 0;'>
            🎨 Designed with 🐑 @小羊
        </p>
    </div>
    """, unsafe_allow_html=True)

# 主函数
def main():
    # 添加一个容器来包裹版本信息
    with st.container():
        col1, col2 = st.columns([3, 1])
        with col2:
            st.markdown(
                f"""
                <div style='background: white; padding: 0.5rem; border-radius: 5px; 
                     border: 1px solid #E0E0E0; margin-bottom: 1rem; text-align: right;
                     font-size: 0.8rem; color: #666;'>
                    Python 版本: {sys.version.split()[0]}
                </div>
                """,
                unsafe_allow_html=True
            )
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
        # 添加筛选框样式
        st.markdown("""
        <style>
            /* 筛选框容器 */
            .filter-box {
                background-color: #EBF5FB;
                padding: 1.2rem;
                border-radius: 12px;
                border: 2px solid #2E86C1;
                box-shadow: 0 2px 8px rgba(46,134,193,0.15);
                margin-bottom: 2rem;
                position: sticky;
                top: 3rem;
                z-index: 1000;
            }
            
            /* 筛选条格布局 */
            .filter-grid {
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
                gap: 1rem;
                margin-top: 0.8rem;
            }
            
            /* 筛选条件项样式 */
            .filter-item {
                background-color: white;
                padding: 0.8rem;
                border-radius: 8px;
                border: 1px solid #E0E0E0;
            }
            
            /* 筛选框标题 */
            .filter-title {
                color: #2E86C1;
                font-size: 1.1rem;
                font-weight: 600;
                margin-bottom: 0.8rem;
                padding-bottom: 0.5rem;
                border-bottom: 2px solid #2E86C1;
            }
        </style>
        """, unsafe_allow_html=True)
        
        # 创建筛选框
        st.markdown("""
        <div class="filter-box">
            <div class="filter-title">🔍 筛选条件</div>
            <div class="filter-grid">
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
                    <div style='background-color: #F8F9FA; padding: 1rem; border-radius: 8px; 
                         border: 1px solid #E0E0E0; height: 100%;'>
                        <h4 style='color: #2E86C1; margin-bottom: 1rem; font-size: 1rem;'>🎮 控制面板</h4>
                    """, unsafe_allow_html=True)
                    
                    # 文件信息
                    st.markdown(f"""
                    <div style='background-color: white; padding: 0.8rem; border-radius: 6px; 
                         margin-bottom: 1rem; border: 1px solid #E0E0E0;'>
                        <p style='margin: 0; font-size: 0.9rem;'>
                            📄 分析文件：<br><strong>{', '.join(f.name for f in selected_files)}</strong>
                        </p>
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
                    st.metric("��评数", low_score_comments)
                    
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
                                            autosize=True,  # 自动调整大小
                                            height=300,  # 固定高度
                                        )
                                        # 调整饼图样式
                                        fig_source.update_traces(
                                            textposition='inside',  # 文字位置
                                            textinfo='percent+label',  # 显示百分比和标签
                                            hole=0.3,  # 添加环形效果
                                            pull=[0.05] * len(source_counts),  # 轻微分离扇形
                                            marker=dict(
                                                colors=['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd'],  # 自定义颜色
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
                                        margin=dict(l=20, r=20, t=20, b=80),  # 增加底部边距，为倾斜的标签留出空间
                                        xaxis_tickangle=-45,  # 标签倾斜角度
                                        xaxis=dict(
                                            tickmode='array',
                                            ticktext=list(top_words.keys()),
                                            tickvals=list(range(len(top_words))),
                                            tickfont=dict(size=11)  # 调整字体大小
                                        ),
                                        yaxis=dict(
                                            title='出现次数',
                                            titlefont=dict(size=12),
                                            tickfont=dict(size=11)
                                        ),
                                        bargap=0.2,  # 调整柱子之间的间距
                                        plot_bgcolor='white',  # 设置背景色为白色
                                        showlegend=False
                                    )
                                    fig.update_traces(
                                        marker_color='#2E86C1',  # 设置柱子颜色
                                        marker_line_color='#2874A6',  # 设置柱子边框颜色
                                        marker_line_width=1,  # 设置柱子边框宽度
                                        opacity=0.8  # 设置透明度
                                    )
                                    fig.update_yaxes(
                                        showgrid=True,
                                        gridwidth=1,
                                        gridcolor='rgba(211,211,211,0.3)'  # 淡灰色网格线
                                    )
                                    st.plotly_chart(fig, use_container_width=True, config={
                                        'displayModeBar': False  # 隐藏plotly工具栏
                                    })
                                    
                                    # 第三行：评论详情
                                    st.markdown("""
                                    <div style="background-color: #f8f9fa; padding: 1rem; border-radius: 8px; margin: 1rem 0;">
                                        <h3 style="color: #2E86C1; font-size: 1.2rem; margin-bottom: 1rem;">💬 评论详情</h3>
                                    """, unsafe_allow_html=True)
                                    
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
                                                st.markdown(
                                                    f"""
                                                    <div style="background-color: white; padding: 1rem; border-radius: 8px; 
                                                         margin-bottom: 0.8rem; border: 1px solid #E0E0E0; height: 100%;">
                                                        <div style="color: #333;">{comment.replace(selected_word, 
                                                            f'<span style="color: red; font-weight: bold;">{selected_word}</span>')}</div>
                                                        <div style="text-align: right; color: #666666; font-size: 0.8em; margin-top: 0.3em;">
                                                            来源: {source}
                                                        </div>
                                                    </div>
                                                    """,
                                                    unsafe_allow_html=True
                                                )
                                    
                                    st.markdown("</div>", unsafe_allow_html=True)
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
                                        st.plotly_chart(fig, use_container_width=True)
                                    
                                    # 第二行：差评详情
                                    st.markdown("""
                                    <div style="background-color: #f8f9fa; padding: 1rem; border-radius: 8px; margin: 1rem 0;">
                                        <h3 style="color: #2E86C1; font-size: 1.2rem; margin-bottom: 1rem;">💬 差评详情</h3>
                                    """, unsafe_allow_html=True)
                                    
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
                                                st.markdown(
                                                    f"""
                                                    <div style="background-color: white; padding: 1rem; border-radius: 8px; 
                                                         margin-bottom: 0.8rem; border: 1px solid #E0E0E0; height: 100%;">
                                                        <div style="color: #333;">{comment.replace(selected_word, 
                                                            f'<span style="color: red; font-weight: bold;">{selected_word}</span>')}</div>
                                                        <div style="text-align: right; color: #666666; font-size: 0.8em; margin-top: 0.3em;">
                                                            来源: {source}
                                                        </div>
                                                    </div>
                                                    """,
                                                    unsafe_allow_html=True
                                                )
                                    
                                    st.markdown("</div>", unsafe_allow_html=True)
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
                                    st.plotly_chart(fig, use_container_width=True)
                                    
                                    # 第二行：建议详情
                                    st.markdown("""
                                    <div style="background-color: #f8f9fa; padding: 1rem; border-radius: 8px; margin: 1rem 0;">
                                        <h3 style="color: #2E86C1; font-size: 1.2rem; margin-bottom: 1rem;">💡 建议详情</h3>
                                    """, unsafe_allow_html=True)
                                    
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
                                                st.markdown(
                                                    f"""
                                                    <div style="background-color: white; padding: 1rem; border-radius: 8px; 
                                                         margin-bottom: 0.8rem; border: 1px solid #E0E0E0; height: 100%;">
                                                        <div style="color: #333;">{comment.replace(selected_word, 
                                                            f'<span style="color: red; font-weight: bold;">{selected_word}</span>')}</div>
                                                        <div style="text-align: right; color: #666666; font-size: 0.8em; margin-top: 0.3em;">
                                                            来源: {source}
                                                        </div>
                                                    </div>
                                                    """,
                                                    unsafe_allow_html=True
                                                )
                                    
                                    st.markdown("</div>", unsafe_allow_html=True)
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
                                        st.plotly_chart(fig, use_container_width=True)
                                    
                                    # 第二行：负面评论详情
                                    st.markdown("""
                                    <div style="background-color: #f8f9fa; padding: 1rem; border-radius: 8px; margin: 1rem 0;">
                                        <h3 style="color: #2E86C1; font-size: 1.2rem; margin-bottom: 1rem;">😟 负面评论详情</h3>
                                    """, unsafe_allow_html=True)
                                    
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
                                                st.markdown(
                                                    f"""
                                                    <div style="background-color: white; padding: 1rem; border-radius: 8px; 
                                                         margin-bottom: 0.8rem; border: 1px solid #E0E0E0; height: 100%;">
                                                        <div style="color: #333;">{comment.replace(selected_word, 
                                                            f'<span style="color: red; font-weight: bold;">{selected_word}</span>')}</div>
                                                        <div style="text-align: right; color: #666666; font-size: 0.8em; margin-top: 0.3em;">
                                                            来源: {source}
                                                        </div>
                                                    </div>
                                                    """,
                                                    unsafe_allow_html=True
                                                )
                                    
                                    st.markdown("</div>", unsafe_allow_html=True)
                            else:
                                st.info("没有找到负面情绪相关的评论")

            except Exception as e:
                st.error(f"处理文件时出错: {str(e)}")

if __name__ == "__main__":
    main()