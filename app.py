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

def highlight_words(text, selected_word):
    """高亮显示选中的关键词"""
    return text.replace(
        selected_word, 
        f'<span style="color: red; font-weight: bold;">{selected_word}</span>'
    )

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
    
    # 在 render_header 函数中，使用说明之前添加系统公告
    image_url = "https://raw.githubusercontent.com/你的用户名/仓库名/main/feedback.png"

    st.markdown("""
    <div style='background: linear-gradient(135deg, #EBF5FB, #D6EAF8); padding: 1rem; border-radius: 8px; 
         margin-bottom: 1.5rem; border: 1px solid #2E86C1;'>
        <h4 style='color: #2E86C1; margin-bottom: 0.8rem; font-size: 1rem; display: flex; align-items: center;'>
            <span style='margin-right: 0.5rem;'>📢</span> 系统更新公告 v2.0.0
        </h4>
        <div style='color: #2E86C1; font-size: 0.9rem; line-height: 1.5;'>
            <p style='margin: 0 0 0.5rem 0;'>更新内容：</p>
            <ul style='margin: 0 0 1rem 1.5rem;'>
                <li>全新界面设计，优化用户体验</li>
                <li>新增筛选框固定显示功能</li>
                <li>优化筛选条件布局和样式</li>
                <li>简化操作流程，提升使用效率</li>
                <li>删除多余的白色框和合并分析功能</li>
            </ul>
            <div style='background: white; padding: 0.8rem; border-radius: 6px; margin-top: 0.5rem;'>
                <p style='margin: 0; color: #2E86C1;'>
                    💡 特别感谢：感谢铭浩提出的宝贵建议，帮助我们实现了多文件上传功能！
                </p>
                <p style='margin: 0.5rem 0 0 0; color: #666; font-size: 0.85rem; font-style: italic; 
                    padding: 0.5rem; background: #F8F9FA; border-radius: 4px;'>
                    "一次只能上传一份，能否同时上传多份文档"
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
            
            /* 筛选条件网格布局 */
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
            selected_file = st.selectbox(
                "📁 选择文件",
                options=uploaded_files,
                format_func=lambda x: x.name
            )
            st.markdown('</div>', unsafe_allow_html=True)
        
        if selected_file:
            try:
                df = pd.read_excel(selected_file, header=5)
                # 查找路线列
                route_col = None
                for col in df.columns:
                    if '路��名称' in str(col) or '产品名称' in str(col):
                        route_col = col
                        routes = df[route_col].dropna().unique()
                        with col2:
                            st.markdown('<div class="filter-item">', unsafe_allow_html=True)
                            selected_routes = st.multiselect(
                                "🛣️ 选择路线",
                                options=routes,
                                default=[],
                                help="选择要分析的路线"
                            )
                            st.markdown('</div>', unsafe_allow_html=True)
                        break
                
                with col3:
                    st.markdown('<div class="filter-item">', unsafe_allow_html=True)
                    filter_keyword = st.text_input(
                        "🔤 关键词",
                        help="输入关键词筛选"
                    )
                    st.markdown('</div>', unsafe_allow_html=True)
                
                with col4:
                    st.markdown('<div class="filter-item">', unsafe_allow_html=True)
                    comment_type = st.multiselect(
                        "📝 评论类型",
                        options=["全部评论", "建议评论", "负面评论"],
                        default=["全部评论"],
                        help="选择评论类型"
                    )
                    st.markdown('</div>', unsafe_allow_html=True)
                
            except Exception as e:
                st.error(f"读取文件时出错: {str(e)}")
        
        st.markdown('</div></div>', unsafe_allow_html=True)
        
        # 主要内容区域
        if selected_file:
            try:
                df = pd.read_excel(selected_file, header=5)
                
                # 查找路线列和评分列
                route_col = None
                score_col = None
                for col in df.columns:
                    if '路线名称' in str(col) or '产品名称' in str(col):
                        route_col = col
                    if '总安排打分' in str(col):
                        score_col = col
                
                if score_col is None:
                    st.error(f'在文件 {selected_file.name} 中未找到"总安排打分"列')
                    return
                
                # 应用筛选条件
                mask = pd.Series(True, index=df.index)  # 初始化全为 True 的掩码
                
                # 添加路线筛选条件
                if route_col and selected_routes:
                    mask = mask & df[route_col].isin(selected_routes)
                
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
                
                # 显示当前文件统计信息
                st.markdown(f"""
                <div style='background-color: #EBF5FB; padding: 0.8rem 1.2rem; border-radius: 8px; 
                     margin: 1rem 0; box-shadow: 0 1px 4px rgba(0,0,0,0.05); display: inline-block;'>
                    <p style='color: #2E86C1; margin: 0; font-size: 0.9rem;'>
                        📄 当前分析文件：<strong>{selected_file.name}</strong>
                    </p>
                </div>
                """, unsafe_allow_html=True)
                
                # 显示筛选结果统计
                total_comments = len(comments)
                low_score_comments = sum(1 for score in scores if score <= 3)
                
                st.markdown("""
                <div style='background-color: #F8F9FA; padding: 1rem; border-radius: 8px; 
                     margin: 1rem 0; border: 1px solid #E0E0E0;'>
                    <h4 style='color: #2E86C1; margin-bottom: 0.8rem; font-size: 1rem;'>筛选结果统计</h4>
                """, unsafe_allow_html=True)
                
                mcol1, mcol2 = st.columns(2)
                with mcol1:
                    st.metric("筛选后评论数", total_comments)
                with mcol2:
                    st.metric("其中差评数", low_score_comments)
                
                st.markdown("</div>", unsafe_allow_html=True)
                
                # 词频统计
                word_freq = Counter()
                word_freq_low = Counter()
                word_comments = {}
                word_comments_low = {}
                suggestion_freq = Counter()
                negative_freq = Counter()
                suggestion_comments = {}
                negative_comments = {}
                
                for comment, score in zip(comments, scores):
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
                            not word.isdigit()):
                            
                            word_freq[word] += 1
                            comment_words.add(word)
                            if word not in word_comments:
                                word_comments[word] = set()
                            word_comments[word].add(comment)
                            
                            if score <= 3:
                                word_freq_low[word] += 1
                                if word not in word_comments_low:
                                    word_comments_low[word] = set()
                                word_comments_low[word].add(comment)
                            
                            # 添加建议词统计
                            if word in SUGGESTION_WORDS:
                                suggestion_freq[word] += 1
                                if word not in suggestion_comments:
                                    suggestion_comments[word] = set()
                                suggestion_comments[word].add(comment)
                            
                            # 添加负面词统计
                            if word in NEGATIVE_WORDS:
                                negative_freq[word] += 1
                                if word not in negative_comments:
                                    negative_comments[word] = set()
                                negative_comments[word].add(comment)
                
                # 修改标签页的显示
                tab1, tab2, tab3, tab4 = st.tabs([
                    "📈  所有评论分析  ",
                    "📉  差评分析  ",
                    "💡  建议分析  ",
                    "😟  负面情绪分析  "
                ])
                
                with tab1:
                    if word_freq:
                        st.subheader("词云图")
                        wc = WordCloud(
                            font_path=ensure_font(),
                            width=800,
                            height=400,
                            background_color='white',
                            max_words=100
                        )
                        wc.generate_from_frequencies(word_freq)
                        st.image(wc.to_array())
                        
                        st.subheader("词频统计")
                        top_words = dict(sorted(word_freq.items(), key=lambda x: x[1], reverse=True)[:20])
                        fig = px.bar(
                            x=list(top_words.keys()),
                            y=list(top_words.values()),
                            labels={'x': '关键词', 'y': '出现次数'}
                        )
                        st.plotly_chart(fig, use_container_width=True)
                        
                        selected_word = st.selectbox(
                            "选择关键词查看相关评论",
                            options=list(top_words.keys())
                        )
                        
                        if selected_word:
                            st.subheader(f"包含 '{selected_word}' 的评论：")
                            relevant_comments = word_comments.get(selected_word, set())
                            unique_comments = get_most_complete_comment(relevant_comments)
                            
                            for comment in unique_comments:
                                st.markdown(
                                    highlight_words(comment, selected_word),
                                    unsafe_allow_html=True
                                )
                    else:
                        st.info("没有找到有效的评论数据")
                
                with tab2:
                    if word_freq_low:
                        st.subheader("差评词云图")
                        wc = WordCloud(
                            font_path=ensure_font(),
                            width=800,
                            height=400,
                            background_color='white',
                            max_words=100
                        )
                        wc.generate_from_frequencies(word_freq_low)
                        st.image(wc.to_array())
                        
                        st.subheader("差评词频统计")
                        top_words = dict(sorted(word_freq_low.items(), key=lambda x: x[1], reverse=True)[:20])
                        fig = px.bar(
                            x=list(top_words.keys()),
                            y=list(top_words.values()),
                            labels={'x': '关键词', 'y': '出现次数'}
                        )
                        st.plotly_chart(fig, use_container_width=True)
                        
                        selected_word = st.selectbox(
                            "选择关键词查看相关差评",
                            options=list(top_words.keys()),
                            key="low_score_select"
                        )
                        
                        if selected_word:
                            st.subheader(f"包含 '{selected_word}' 的差评：")
                            relevant_comments = word_comments_low.get(selected_word, set())
                            unique_comments = get_most_complete_comment(relevant_comments)
                            
                            for comment in unique_comments:
                                st.markdown(
                                    highlight_words(comment, selected_word),
                                    unsafe_allow_html=True
                                )
                    else:
                        st.info("没有找到差评数据")
                
                # 添加建议分析标签页
                with tab3:
                    if suggestion_freq:
                        st.subheader("建议关键词统计")
                        top_suggestions = dict(sorted(suggestion_freq.items(), key=lambda x: x[1], reverse=True)[:20])
                        fig = px.bar(
                            x=list(top_suggestions.keys()),
                            y=list(top_suggestions.values()),
                            labels={'x': '建议关键词', 'y': '出现次数'},
                            title="建议关键词分布"
                        )
                        st.plotly_chart(fig, use_container_width=True)
                        
                        selected_word = st.selectbox(
                            "选择关键词查看相关建议",
                            options=list(top_suggestions.keys()),
                            key="suggestion_select"
                        )
                        
                        if selected_word:
                            st.subheader(f"包含 '{selected_word}' 的评论：")
                            relevant_comments = suggestion_comments.get(selected_word, set())
                            unique_comments = get_most_complete_comment(relevant_comments)
                            
                            for comment in unique_comments:
                                st.markdown(
                                    highlight_words(comment, selected_word),
                                    unsafe_allow_html=True
                                )
                    else:
                        st.info("没有找到建议相关的评论")
                
                # 添加负面情绪分析标签页
                with tab4:
                    if negative_freq:
                        st.subheader("负面情绪词统计")
                        top_negative = dict(sorted(negative_freq.items(), key=lambda x: x[1], reverse=True)[:20])
                        fig = px.bar(
                            x=list(top_negative.keys()),
                            y=list(top_negative.values()),
                            labels={'x': '负面情绪词', 'y': '出现次数'},
                            title="负面情绪词分布"
                        )
                        st.plotly_chart(fig, use_container_width=True)
                        
                        selected_word = st.selectbox(
                            "选择关键词查看相关负面评论",
                            options=list(top_negative.keys()),
                            key="negative_select"
                        )
                        
                        if selected_word:
                            st.subheader(f"包含 '{selected_word}' 的评论：")
                            relevant_comments = negative_comments.get(selected_word, set())
                            unique_comments = get_most_complete_comment(relevant_comments)
                            
                            for comment in unique_comments:
                                st.markdown(
                                    highlight_words(comment, selected_word),
                                    unsafe_allow_html=True
                                )
                    else:
                        st.info("没有找到负面情绪相关的评论")
                
            except Exception as e:
                st.error(f"处理文件 {selected_file.name} 时出错: {str(e)}")

if __name__ == "__main__":
    main()