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
    # 在页面开头添加自定义 CSS
    st.markdown("""
    <style>
        /* 全局样式 */
        .main {
            padding: 2rem;
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
    </style>
    """, unsafe_allow_html=True)
    
    # 添加页面标题和说明
    st.markdown("""
    <div style='text-align: center; padding: 3rem 0; background: linear-gradient(135deg, #1976D2, #2196F3); 
         border-radius: 20px; margin: 1rem 0 3rem 0; box-shadow: 0 4px 20px rgba(25,118,210,0.2);'>
        <h1 style='color: white; margin-bottom: 1.5rem; font-size: 2.8rem; font-weight: 600;'>
            📊 Excel评论分析工具
        </h1>
        <p style='color: white; opacity: 0.95; font-size: 1.3rem; max-width: 700px; margin: 0 auto;'>
            快速分析和可视化您的评论数据
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    # 使用说明卡片
    st.markdown("""
    <div style='background-color: #F5F7FA; padding: 2rem; border-radius: 20px; margin-bottom: 3rem; 
         box-shadow: 0 4px 20px rgba(0,0,0,0.05);'>
        <h4 style='color: #1976D2; margin-bottom: 1.5rem; font-size: 1.3rem;'>
            <i>📝 使用指南</i>
        </h4>
        <div style='display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 1.5rem;'>
            <div style='background: white; padding: 1.5rem; border-radius: 15px; border-left: 5px solid #1976D2;
                 box-shadow: 0 4px 15px rgba(0,0,0,0.05); transition: transform 0.3s ease;'>
                <p style='margin: 0; color: #1E1E1E; font-size: 1.1rem;'>1. 上传Excel文件（.xlsx格式）</p>
            </div>
            <div style='background: white; padding: 1.5rem; border-radius: 15px; border-left: 5px solid #1976D2;
                 box-shadow: 0 4px 15px rgba(0,0,0,0.05); transition: transform 0.3s ease;'>
                <p style='margin: 0; color: #1E1E1E; font-size: 1.1rem;'>2. 系统自动分析评论内容和评分</p>
            </div>
            <div style='background: white; padding: 1.5rem; border-radius: 15px; border-left: 5px solid #1976D2;
                 box-shadow: 0 4px 15px rgba(0,0,0,0.05); transition: transform 0.3s ease;'>
                <p style='margin: 0; color: #1E1E1E; font-size: 1.1rem;'>3. 查看词云图和词频统计</p>
            </div>
            <div style='background: white; padding: 1.5rem; border-radius: 15px; border-left: 5px solid #1976D2;
                 box-shadow: 0 4px 15px rgba(0,0,0,0.05); transition: transform 0.3s ease;'>
                <p style='margin: 0; color: #1E1E1E; font-size: 1.1rem;'>4. 选择关键词查看相关评论</p>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # 文件上传区域美化
    st.markdown("""
    <div style='background-color: #F8F9FA; padding: 2rem; border-radius: 10px; border: 2px dashed #1E88E5; text-align: center; margin-bottom: 2rem;'>
    """, unsafe_allow_html=True)
    
    uploaded_file = st.file_uploader(
        "选择Excel文件上传",
        type=['xlsx'],
        help="请上传包含评论数据的Excel文件（.xlsx格式）"
    )
    
    st.markdown("</div>", unsafe_allow_html=True)
    
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
            
            # 美化统计指标显示
            st.markdown("<br>", unsafe_allow_html=True)
            col1, col2 = st.columns(2)
            with col1:
                st.markdown(f"""
                <div style='background: linear-gradient(135deg, #1E88E5, #64B5F6); padding: 1.5rem; border-radius: 10px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1);'>
                    <h3 style='color: white; margin: 0; font-size: 1.1rem; opacity: 0.9;'>总评论数</h3>
                    <h2 style='color: white; margin: 0.5rem 0; font-size: 2rem;'>{total_comments}</h2>
                </div>
                """, unsafe_allow_html=True)
            with col2:
                st.markdown(f"""
                <div style='background: linear-gradient(135deg, #EF5350, #E57373); padding: 1.5rem; border-radius: 10px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1);'>
                    <h3 style='color: white; margin: 0; font-size: 1.1rem; opacity: 0.9;'>差评数量</h3>
                    <h2 style='color: white; margin: 0.5rem 0; font-size: 2rem;'>{low_score_comments}</h2>
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
            
            # 修改标签页的显示
            tab1, tab2 = st.tabs([
                "📈  所有评论分析  ",  # 添加额外的空格使文本居中
                "📉  差评分析  "
            ])
            
            with tab1:
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
                        st.markdown(f"""
                        <div style='background: linear-gradient(to right, #1E88E5, #64B5F6); padding: 1rem; border-radius: 10px; margin: 1rem 0;'>
                            <h4 style='color: white; margin: 0;'>包含 "{selected_word}" 的评论：</h4>
                        </div>
                        """, unsafe_allow_html=True)
                        
                        # 获取评论并保留最完整的版本
                        relevant_comments = word_comments.get(selected_word, set())
                        unique_comments = get_most_complete_comment(relevant_comments)
                        
                        # 显示去重后的评论
                        for comment in unique_comments:
                            st.markdown(f"""
                            <div style='background-color: white; padding: 1.5rem; border-radius: 8px; margin: 0.8rem 0; border: 1px solid #E0E0E0; box-shadow: 0 2px 4px rgba(0,0,0,0.05);'>
                                {highlight_words(comment, selected_word)}
                            </div>
                            """, unsafe_allow_html=True)
                else:
                    st.info("没有找到有效的评论数据")
            
            with tab2:
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
                        st.markdown(f"""
                        <div style='background-color: #EBF5FB; padding: 1rem; border-radius: 10px; margin: 1rem 0;'>
                            <h4 style='color: #2E86C1; margin: 0;'>包含 "{selected_word}" 的差评：</h4>
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
                            <div style='background-color: #FFFFFF; padding: 1rem; border-radius: 5px; margin: 0.5rem 0; border: 1px solid #D4E6F1;'>
                                {highlight_words(original_comment, selected_word)}
                            </div>
                            """, unsafe_allow_html=True)
                else:
                    st.info("没有找到差评数据")
                    
        except Exception as e:
            st.error(f"处理文件时出错: {str(e)}")
    
    # 添加设计者标识
    st.markdown("<br><br>", unsafe_allow_html=True)
    st.markdown("""
    <div style='text-align: center; padding: 2rem; margin-top: 2rem; background: linear-gradient(to right, #F8F9FA, #FFFFFF, #F8F9FA);'>
        <p style='color: #757575; font-size: 0.9rem; margin: 0;'>
            🎨 Designed with 🐑 @小羊
        </p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()