import docx
from collections import Counter
import re
import os
import time

def read_docx(file_path):
    """读取Word文档内容"""
    doc = docx.Document(file_path)
    full_text = []
    for para in doc.paragraphs:
        full_text.append(para.text)
    return ' '.join(full_text)

def process_text(text):
    """处理文本：转小写，去除标点符号等"""
    # 转换为小写
    text = text.lower()
    # 只保留字母和空格
    text = re.sub(r'[^a-zA-Z\s]', '', text)
    return text

def get_word_frequency(text, top_n=100):
    """获取词频统计"""
    # 简单的按空格分词
    words = text.split()
    
    # 常见的英语停用词
    stop_words = {'the', 'a', 'an', 'and', 'or', 'but', 'in', 'on', 'at', 'to',
                  'for', 'of', 'with', 'by', 'from', 'up', 'about', 'into',
                  'over', 'after', 'is', 'are', 'was', 'were', 'be', 'been',
                  'being', 'have', 'has', 'had', 'do', 'does', 'did', 'will',
                  'would', 'shall', 'should', 'may', 'might', 'must', 'can',
                  'could', 'i', 'you', 'he', 'she', 'it', 'we', 'they', 'this',
                  'that', 'these', 'those'}
    
    # 过滤停用词和长度为1的词
    words = [word for word in words if word not in stop_words and len(word) > 1]
    
    # 统计词频
    word_freq = Counter(words)
    
    # 返回前N个高频词
    return word_freq.most_common(top_n)

def save_results(word_frequencies, original_file):
    """保存结果到文本文件"""
    # 创建结果文件名
    base_name = os.path.splitext(os.path.basename(original_file))[0]
    result_file = f"词频统计_{base_name}_{time.strftime('%Y%m%d_%H%M%S')}.txt"
    
    with open(result_file, 'w', encoding='utf-8') as f:
        f.write("高频词汇统计结果：\n")
        f.write("排名\t单词\t\t出现次数\n")
        f.write("-" * 30 + "\n")
        for rank, (word, freq) in enumerate(word_frequencies, 1):
            f.write(f"{rank}\t{word:<15}{freq}\n")
    
    return result_file

def main():
    print("=" * 50)
    print("英文文档词频统计工具")
    print("=" * 50)
    print("\n请将Word文档拖放到这个窗口中，然后按回车键")
    
    try:
        # 获取文件路径（去除可能的引号）
        file_path = input("> ").strip('" ')
        
        print("\n正在处理文档...")
        # 读取文档
        text = read_docx(file_path)
        
        # 处理文本
        processed_text = process_text(text)
        
        # 获取词频统计
        word_frequencies = get_word_frequency(processed_text)
        
        # 保存结果到文件
        result_file = save_results(word_frequencies, file_path)
        
        # 打印结果
        print("\n高频词汇统计结果：")
        print("排名\t单词\t\t出现次数")
        print("-" * 30)
        for rank, (word, freq) in enumerate(word_frequencies, 1):
            print(f"{rank}\t{word:<15}{freq}")
        
        print(f"\n结果已保存到文件：{result_file}")
            
    except Exception as e:
        print(f"\n发生错误: {str(e)}")
    
    print("\n按回车键退出程序...")
    input()

if __name__ == "__main__":
    main() 