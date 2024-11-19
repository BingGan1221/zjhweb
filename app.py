import streamlit as st

# 设置页面配置必须是第一个 Streamlit 命令
st.set_page_config(
    page_title="Excel评论分析工具",
    page_icon="📊",
    layout="wide"
)

import sys
import logging

# 设置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# 显示版本信息
st.write(f"Python 版本: {sys.version}")
logger.info(f"Python 版本: {sys.version}")

try:
    import pandas as pd
    import numpy as np
    from collections import Counter
    from pathlib import Path
    import shutil
    import requests
    import plotly.express as px
    
    logger.info("基础依赖包导入成功")
except Exception as e:
    logger.error(f"基础依赖包导入失败: {str(e)}")
    st.error(f"基础依赖包导入失败: {str(e)}")
    st.stop()

try:
    from wordcloud import WordCloud
    logger.info("WordCloud 导入成功")
except ImportError as e:
    logger.error(f"WordCloud 导入失败: {str(e)}")
    st.error("无法导入 WordCloud 包，请检查依赖安装")
    st.stop()

try:
    import jieba
    logger.info("jieba 导入成功")
except ImportError as e:
    logger.error(f"jieba 导入失败: {str(e)}")
    st.error("无法导入 jieba 包，请检查依赖安装")
    st.stop()

# 定义停用词列表
STOP_WORDS = {
    '的', '了', '和', '是', '就', '都', '而', '及', '与', '着',
    '之', '用', '于', '把', '等', '去', '又', '能', '好', '在',
    '或', '这', '那', '有', '很', '只', '些', '为', '呢', '啊',
    '并', '给', '跟', '还', '个', '之类', '各种', '没有', '非常',
    '可以', '因为', '因此', '所以', '但是', '但', '然后', '如果',
    '虽然', '这样', '这些', '那些', '如此', '只是', '真的', '一个',
}

def ensure_font():
    font_dir = Path('fonts')
    font_path = font_dir / 'simhei.ttf'
    
    if font_path.exists():
        return str(font_path)
    
    font_dir.mkdir(exist_ok=True)
    
    system_fonts = [
        Path('C:/Windows/Fonts/simhei.ttf'),
        Path('/usr/share/fonts/truetype/simhei.ttf'),
        Path('/System/Library/Fonts/simhei.ttf')
    ]
    
    for system_font in system_fonts:
        if system_font.exists():
            shutil.copy(system_font, font_path)
            return str(font_path)
    
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

def highlight_words(text, words):
    """高亮显示文本中的多个关键词"""
    result = text
    for word in words:
        result = result.replace(word, f'**:red[{word}]**')
    return result

def main():
    st.title("Excel评论分析工具")
    
    font_path = ensure_font()
    if not font_path:
        st.error("无法加载中文字体文件，词云图可能无法正确显示中文。")
    
    uploaded_file = st.file_uploader("上传Excel文件", type=['xlsx'])
    
    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file, header=5)
            
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
            
            total_comments = len(comments)
            low_score_comments = sum(1 for score in scores if score <= 3)
            
            col1, col2 = st.columns(2)
            with col1:
                st.metric("总评论数", total_comments)
            with col2:
                st.metric("差评数量", low_score_comments)
            
            word_freq = Counter()
            word_freq_low = Counter()
            comment_words = {}  # 存储每条评论包含的所有高频词
            comment_words_low = {}  # 存储每条差评包含的所有高频词
            
            for comment, score in zip(comments, scores):
                if pd.isna(comment) or pd.isna(score):
                    continue
                    
                comment = str(comment).strip()
                if not comment:
                    continue
                
                words = jieba.cut(comment)
                comment_high_freq_words = set()  # 当前评论包含的高频词
                
                for word in words:
                    word = word.strip()
                    if (len(word) > 1 and
                        word not in STOP_WORDS and
                        not word.isdigit()):
                        
                        word_freq[word] += 1
                        if score <= 3:
                            word_freq_low[word] += 1
                        comment_high_freq_words.add(word)
                
                if comment_high_freq_words:
                    comment_words[comment] = comment_high_freq_words
                    if score <= 3:
                        comment_words_low[comment] = comment_high_freq_words
            
            tab1, tab2 = st.tabs(["所有评论分析", "差评分析"])
            
            with tab1:
                if word_freq:
                    st.subheader("词云图")
                    wc = WordCloud(
                        font_path=font_path,
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
                        title="词频统计（前20个词）"
                    )
                    st.plotly_chart(fig)
                    
                    selected_word = st.selectbox(
                        "选择词语查看相关评论",
                        options=list(top_words.keys())
                    )
                    
                    if selected_word:
                        st.write(f"包含 '{selected_word}' 的评论：")
                        # 只显示包含所选词的评论，并高亮所有高频词
                        relevant_comments = [
                            (comment, words) 
                            for comment, words in comment_words.items() 
                            if selected_word in words
                        ]
                        for comment, words in relevant_comments:
                            high_freq_words = [w for w in words if w in top_words]
                            st.markdown(highlight_words(comment, high_freq_words), unsafe_allow_html=True)
                else:
                    st.info("没有找到有效的评论数据")
            
            with tab2:
                if word_freq_low:
                    st.subheader("差评词云图")
                    wc = WordCloud(
                        font_path=font_path,
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
                        title="差评词频统计（前20个词）"
                    )
                    st.plotly_chart(fig)
                    
                    selected_word = st.selectbox(
                        "选择词语查看相关差评",
                        options=list(top_words.keys()),
                        key="low_score_select"
                    )
                    
                    if selected_word:
                        st.write(f"包含 '{selected_word}' 的差评：")
                        # 只显示包含所选词的差评，并高亮所有高频词
                        relevant_comments = [
                            (comment, words) 
                            for comment, words in comment_words_low.items() 
                            if selected_word in words
                        ]
                        for comment, words in relevant_comments:
                            high_freq_words = [w for w in words if w in top_words]
                            st.markdown(highlight_words(comment, high_freq_words), unsafe_allow_html=True)
                else:
                    st.info("没有找到差评数据")
                    
        except Exception as e:
            st.error(f"处理文件时出错: {str(e)}")

if __name__ == "__main__":
    main()