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

def highlight_words(text, selected_word):
    """只高亮选中的关键词"""
    # 使用 HTML 标签而不是 Markdown
    return text.replace(selected_word, f'<span style="color: red; font-weight: bold;">{selected_word}</span>')

def get_most_complete_comment(comments):
    """从相似评论中选择最长的一条（通常是最完整的）"""
    normalized_comments = {}
    for comment in comments:
        normalized = ' '.join(comment.split())  # 规范化空白字符
        # 如果已存在相似评论，保留较长的那个
        if normalized in normalized_comments:
            if len(comment) > len(normalized_comments[normalized]):
                normalized_comments[normalized] = comment
        else:
            normalized_comments[normalized] = comment
    return list(normalized_comments.values())

def main():
    # 加载自定义CSS
    with open('.streamlit/style.css') as f:
        st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)
    
    # 页面标题
    st.markdown("""
    <div class="title-container">
        <h1 class="title-text">📊 Excel评论分析工具</h1>
        <p class="subtitle-text">快速分析和可视化您的评论数据</p>
    </div>
    """, unsafe_allow_html=True)
    
    # 使用指南部分
    st.markdown("""
    <div class="guide-container">
        <h4 class="guide-title">📝 使用指南</h4>
        <div class="guide-grid">
            <div class="guide-item">
                <p>1. 上传Excel文件（.xlsx格式）</p>
            </div>
            <div class="guide-item">
                <p>2. 系统自动分析评论内容和评分</p>
            </div>
            <div class="guide-item">
                <p>3. 查看词云图和词频统计</p>
            </div>
            <div class="guide-item">
                <p>4. 选择关键词查看相关评论</p>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # 文件上传区域
    st.markdown('<div class="upload-container">', unsafe_allow_html=True)
    uploaded_file = st.file_uploader(
        "选择Excel文件上传",
        type=['xlsx'],
        help="请上传包含评论数据的Excel文件（.xlsx格式）"
    )
    st.markdown('</div>', unsafe_allow_html=True)
    
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
            
            # 统计指标显示
            st.markdown("""
            <div class="stats-container">
                <div class="stat-card">
                    <h3>总评论数</h3>
                    <h2>{total_comments}</h2>
                </div>
                <div class="stat-card stat-card-red">
                    <h3>差评数量</h3>
                    <h2>{low_score_comments}</h2>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            word_freq = Counter()
            word_freq_low = Counter()
            word_comments = {}  # 存储每个词对应的论列表
            word_comments_low = {}  # 存储每个词对应的差评列表
            
            for comment, score in zip(comments, scores):
                if pd.isna(comment) or pd.isna(score):
                    continue
                    
                comment = str(comment).strip()
                if not comment:
                    continue
                
                words = jieba.cut(comment)
                comment_words = set()  # 用于记录当前评论中的高频词
                
                for word in words:
                    word = word.strip()
                    if (len(word) > 1 and
                        word not in STOP_WORDS and
                        not word.isdigit()):
                        
                        word_freq[word] += 1
                        comment_words.add(word)
                        if word not in word_comments:
                            word_comments[word] = set()  # 使用集合避免重复
                        word_comments[word].add(comment)
                        
                        if score <= 3:
                            word_freq_low[word] += 1
                            if word not in word_comments_low:
                                word_comments_low[word] = set()  # 使用集合避免重复
                            word_comments_low[word].add(comment)
            
            # 词云分析部分
            if word_freq:
                st.markdown("""
                <div class="analysis-section">
                    <h3 class="section-title">词云分析</h3>
                    <div class="wordcloud-container">
                        <!-- 词云图将在这里显示 -->
                    </div>
                </div>
                """, unsafe_allow_html=True)
                
                # 词云图生成代码
                if word_freq:
                    st.markdown("""
                    <h3 style='color: #2E86C1; margin: 2rem 0 1rem 0;'>词云图</h3>
                    """, unsafe_allow_html=True)
                    wc = WordCloud(
                        font_path=ensure_font(),
                        width=800,
                        height=400,
                        background_color='white',
                        max_words=100
                    )
                    wc.generate_from_frequencies(word_freq)
                    st.image(wc.to_array())
                    
                    st.markdown("""
                    <h3 style='color: #2E86C1; margin: 2rem 0 1rem 0;'>词频统计</h3>
                    """, unsafe_allow_html=True)
                    # 美化词频图表
                    top_words = dict(sorted(word_freq.items(), key=lambda x: x[1], reverse=True)[:20])
                    fig = px.bar(
                        x=list(top_words.keys()),
                        y=list(top_words.values()),
                        title="词频统计（前20个词）",
                        labels={'x': '关键词', 'y': '出现次数'},
                        color_discrete_sequence=['#2E86C1']
                    )
                    fig.update_layout(
                        title_x=0.5,
                        title_font_size=20,
                        plot_bgcolor='rgba(0,0,0,0)',
                        paper_bgcolor='rgba(0,0,0,0)'
                    )
                    st.plotly_chart(fig, use_container_width=True)
                    
                    # 美化选择框和评论显示
                    st.markdown("<br>", unsafe_allow_html=True)
                    selected_word = st.selectbox(
                        "选择关键词查看相关评论",
                        options=list(top_words.keys()),
                        help="选择一个关键词，查看包含该词的所有评论"
                    )
                    
                    if selected_word:
                        st.markdown("""
                        <div class="comment-list">
                            <h4>包含 "{selected_word}" 的评论：</h4>
                            <!-- 评论将在这里显示 -->
                        </div>
                        """, unsafe_allow_html=True)
                        
                        # 获取评论并保留最完整的版本
                        relevant_comments = word_comments.get(selected_word, set())
                        unique_comments = get_most_complete_comment(relevant_comments)
                        
                        # 显示去重后的评论
                        for comment in unique_comments:
                            st.markdown(f"""
                            <div class="comment-item">
                                {highlight_words(comment, selected_word)}
                            </div>
                            """, unsafe_allow_html=True)
            else:
                st.info("没有找到有效的评论数据")
            
            # 差评分析部分
            if word_freq_low:
                st.markdown("""
                <div class="analysis-section">
                    <h3 class="section-title">差评分析</h3>
                    <div class="wordcloud-container">
                        <!-- 差评词云图将在这里显示 -->
                    </div>
                </div>
                """, unsafe_allow_html=True)
                
                # 差评词云图生成代码
                if word_freq_low:
                    st.markdown("""
                    <h3 style='color: #2E86C1; margin: 2rem 0 1rem 0;'>差评词云图</h3>
                    """, unsafe_allow_html=True)
                    wc = WordCloud(
                        font_path=ensure_font(),
                        width=800,
                        height=400,
                        background_color='white',
                        max_words=100
                    )
                    wc.generate_from_frequencies(word_freq_low)
                    st.image(wc.to_array())
                    
                    st.markdown("""
                    <h3 style='color: #2E86C1; margin: 2rem 0 1rem 0;'>差评词频统计</h3>
                    """, unsafe_allow_html=True)
                    # 美化词频图表
                    top_words = dict(sorted(word_freq_low.items(), key=lambda x: x[1], reverse=True)[:20])
                    fig = px.bar(
                        x=list(top_words.keys()),
                        y=list(top_words.values()),
                        title="差评词频统计（前20个词）",
                        labels={'x': '关键词', 'y': '出现次数'},
                        color_discrete_sequence=['#2E86C1']
                    )
                    fig.update_layout(
                        title_x=0.5,
                        title_font_size=20,
                        plot_bgcolor='rgba(0,0,0,0)',
                        paper_bgcolor='rgba(0,0,0,0)'
                    )
                    st.plotly_chart(fig, use_container_width=True)
                    
                    # 美化选择框和评论显示
                    st.markdown("<br>", unsafe_allow_html=True)
                    selected_word = st.selectbox(
                        "选择关键词查看相关差评",
                        options=list(top_words.keys()),
                        help="选择一个关键词，查看包含该词的所有差评",
                        key="low_score_select"
                    )
                    
                    if selected_word:
                        st.markdown("""
                        <div class="comment-list">
                            <h4>包含 "{selected_word}" 的差评：</h4>
                            <!-- 差评将在这里显示 -->
                        </div>
                        """, unsafe_allow_html=True)
                        
                        # 获取差评并去重
                        relevant_comments = word_comments_low.get(selected_word, set())
                        # 将评论内容规范化（去除空格、换行等）后再去重
                        normalized_comments = {}
                        for comment in relevant_comments:
                            normalized = ' '.join(comment.split())  # 规范化空白字符
                            normalized_comments[normalized] = comment  # 使用原始评论作为值
                        
                        # 显示去重后的评论
                        for original_comment in normalized_comments.values():
                            st.markdown(f"""
                            <div class="comment-item">
                                {highlight_words(original_comment, selected_word)}
                            </div>
                            """, unsafe_allow_html=True)
                else:
                    st.info("没有找到差评数据")
                    
        except Exception as e:
            st.error(f"处理文件时出错: {str(e)}")
    
    # 页脚
    st.markdown("""
    <div class="footer">
        <p>🎨 Designed with 🐑 @小羊</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()