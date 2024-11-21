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

# 显示版本信息
st.write(f"Python 版本: {sys.version}")
logger.info(f"Python 版本: {sys.version}")

# 停用词列表
STOP_WORDS = {
    '的', '了', '和', '是', '就', '都', '而', '及', '与', '着',
    '之', '用', '于', '把', '等', '去', '又', '能', '好', '在',
    '或', '这', '那', '有', '很', '只', '些', '为', '呢', '啊',
    '并', '给', '跟', '还', '个', '之类', '各种', '没有', '非常',
    '可以', '因为', '因此', '所以', '但是', '但', '然后', '如果',
    '虽然', '这样', '这些', '那些', '如此', '只是', '真的', '一个',
}

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
    # 添加全局样式
    st.markdown("""
    <style>
        /* 全局样式 */
        .stApp {
            background: linear-gradient(to bottom, #F8F9F9, #FFFFFF);
        }
        
        /* 美化标题区域 */
        .title-container {
            background: linear-gradient(135deg, #2E86C1, #3498DB);
            padding: 2rem;
            border-radius: 15px;
            margin-bottom: 2rem;
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
            text-align: center;
        }
        
        /* 美化使用说明卡片 */
        .instruction-card {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 1rem;
            padding: 1.5rem;
            background-color: #EBF5FB;
            border-radius: 12px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.05);
        }
        
        .instruction-item {
            background: white;
            padding: 1rem;
            border-radius: 8px;
            border-left: 4px solid #2E86C1;
            transition: transform 0.2s ease;
        }
        
        .instruction-item:hover {
            transform: translateY(-2px);
        }
        
        /* 美化统计卡片 */
        .metric-container {
            background: white;
            padding: 1.5rem;
            border-radius: 10px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.05);
            text-align: center;
        }
        
        /* 美化标签页 */
        .stTabs [data-baseweb="tab-list"] {
            gap: 20px;
            padding: 0.5rem;
        }
        
        .stTabs [data-baseweb="tab"] {
            background-color: white;
            border-radius: 8px;
            padding: 0.5rem 1.5rem;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        }
        
        /* 美化上传区域 */
        .stUploadButton {
            border: 2px dashed #2E86C1;
            border-radius: 10px;
            padding: 1rem;
        }
    </style>
    """, unsafe_allow_html=True)
    
    # 渲染标题
    st.markdown("""
    <div class="title-container">
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
    <div class="instruction-card">
        <div class="instruction-item">
            <h4 style='color: #2E86C1; margin: 0 0 0.5rem 0;'>1️⃣ 上传文件</h4>
            <p style='margin: 0; color: #666;'>支持.xlsx格式的Excel文件</p>
        </div>
        <div class="instruction-item">
            <h4 style='color: #2E86C1; margin: 0 0 0.5rem 0;'>2️⃣ 自动分析</h4>
            <p style='margin: 0; color: #666;'>系统自动分析评论内容和评分</p>
        </div>
        <div class="instruction-item">
            <h4 style='color: #2E86C1; margin: 0 0 0.5rem 0;'>3️⃣ 数据可视化</h4>
            <p style='margin: 0; color: #666;'>查看词云图和词频统计</p>
        </div>
        <div class="instruction-item">
            <h4 style='color: #2E86C1; margin: 0 0 0.5rem 0;'>4️⃣ 深入分析</h4>
            <p style='margin: 0; color: #666;'>选择关键词查看相关评论</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

def render_footer():
    """渲染页面底部信息"""
    st.markdown("""
    <div style='text-align: center; padding: 2rem; margin-top: 3rem; 
         background: linear-gradient(to right, #F8F9F9, #FFFFFF, #F8F9F9);'>
        <p style='color: #95A5A6; font-size: 0.9rem; margin: 0;'>
            🎨 Designed with 🐑 @小羊
        </p>
    </div>
    """, unsafe_allow_html=True)

# 主函数
def main():
    render_header()
    
    # 文件上传
    uploaded_file = st.file_uploader(
        "选择Excel文件上传",
        type=['xlsx'],
        help="请上传包含评论数据的Excel文件（.xlsx格式）"
    )
    
    if not uploaded_file:
        return
        
    try:
        df = pd.read_excel(uploaded_file, header=5)
        
        # 查找评分列
        score_col = None
        for col in df.columns:
            if '总安排打分' in str(col):
                score_col = col
                break
        
        if score_col is None:
            st.error('未找到"总安排打分"列')
            return
        
        comments = df.iloc[:, 0]
        scores = df[score_col]
        
        # 显示基础统计信息
        total_comments = len(comments)
        low_score_comments = sum(1 for score in scores if score <= 3)
        
        col1, col2 = st.columns(2)
        with col1:
            st.metric("总评论数", total_comments)
        with col2:
            st.metric("差评数量", low_score_comments)
        
        # 词频统计
        word_freq = Counter()
        word_freq_low = Counter()
        word_comments = {}  # 存储每个词对应的评论列表
        word_comments_low = {}  # 存储每个词对应的差评列表
        
        for comment, score in zip(comments, scores):
            if pd.isna(comment) or pd.isna(score):
                continue
                
            comment = str(comment).strip()
            if not comment:
                continue
            
            words = jieba.cut(comment)
            comment_words = set()  # 用于记录当前评论中的词
            
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
        
        # 创建标签页
        tab1, tab2 = st.tabs(["📈 所有评论分析", "📉 差评分析"])
        
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
                
    except Exception as e:
        st.error(f"处理文件时出错: {str(e)}")
    
    render_footer()

if __name__ == "__main__":
    main()