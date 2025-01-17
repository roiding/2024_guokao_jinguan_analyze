import pandas as pd
import warnings
import os
warnings.filterwarnings('ignore', category=UserWarning)

def compare_interview_scores(original_file, supplementary_file):
    """
    比较原始进面名单和递补名单的分数线
    
    Args:
        original_file (str): 原始进面名单文件路径
        supplementary_file (str): 递补名单文件路径
        
    Returns:
        DataFrame: 所有岗位的进面和递补情况
    """
    # 读取原始进面名单
    df = pd.read_excel(original_file, dtype={'职位代码': str})
    df = df[df['用人司局'].str.contains('国家金融监督管理')]
    
    # 计算每个职位的原始进面人数
    original_counts = df.groupby(['招录机关', '职位代码']).size().reset_index(name='原始进面人数')
    
    # 获取每个职位的基本信息和分数线（取第一条记录）
    original_info = df.groupby(['招录机关', '职位代码']).agg({
        '用人司局': 'first',
        '招考职位': 'first',
        '最低面试分数': 'first'
    }).reset_index()
    
    # 读取递补名单
    df_递补 = pd.read_excel(supplementary_file, dtype={'职位代码': str})
    
    # 计算每个职位的递补进面人数
    supplementary_counts = df_递补.groupby(['部门名称', '职位代码']).size().reset_index(name='递补进面人数')
    supplementary_counts = supplementary_counts.rename(columns={'部门名称': '招录机关'})
    
    # 获取递补分数线
    supplementary_scores = df_递补.groupby(['部门名称', '职位代码']).agg({
        '递补入围面试最低分数': 'first'
    }).reset_index()
    supplementary_scores = supplementary_scores.rename(columns={'部门名称': '招录机关'})
    
    # 合并所有信息
    result = pd.merge(original_info, original_counts, on=['招录机关', '职位代码'], how='left')
    result = pd.merge(result, supplementary_counts, on=['招录机关', '职位代码'], how='left')
    result = pd.merge(result, supplementary_scores, on=['招录机关', '职位代码'], how='left')
    
    # 填充空值
    result['递补进面人数'] = result['递补进面人数'].fillna(0).astype(int)
    result['递补入围面试最低分数'] = result['递补入围面试最低分数'].fillna('--')
    
    # 计算分数线变化
    result['分数线变化'] = pd.to_numeric(result['递补入围面试最低分数'], errors='coerce') - result['最低面试分数']
    result['分数线变化'] = result['分数线变化'].apply(lambda x: f"{x:+.3f}" if pd.notnull(x) else '--')
    
    # 整理列顺序
    result = result[[
        '招录机关',
        '职位代码',
        '用人司局',
        '招考职位',
        '原始进面人数',
        '最低面试分数',
        '递补进面人数',
        '递补入围面试最低分数',
        '分数线变化'
    ]]
    
    return result

def save_to_excel(df, output_file='result.xlsx'):
    """
    将结果保存到Excel文件
    
    Args:
        df (DataFrame): 要保存的数据
        output_file (str): 输出文件路径
    """
    # 确保result文件夹存在
    if not os.path.exists('result'):
        os.makedirs('result')
        
    # 构建完整的输出路径
    output_path = os.path.join('result', output_file)
    
    # 如果文件已存在则删除
    if os.path.exists(output_path):
        os.remove(output_path)
        
    with pd.ExcelWriter(output_path, engine='openpyxl', mode='w') as writer:
        df['职位代码'] = df['职位代码'].astype(str)
        
        # 写入Excel
        df.to_excel(
            writer,
            index=False,
            sheet_name='进面分数线情况'
        )
        
        # 获取工作表
        worksheet = writer.sheets['进面分数线情况']
        
        # 设置职位代码列的格式为文本
        for cell in worksheet['B'][1:]:  # 假设职位代码在B列
            cell.number_format = '@'
        
        # 调整列宽
        for idx, col in enumerate(df.columns):
            max_length = max(
                df[col].astype(str).apply(len).max(),
                len(str(col))
            )
            worksheet.column_dimensions[chr(65 + idx)].width = max_length + 2

def analyze_supplementary_admission(supplementary_file, admission_file):
    """
    分析递补面试人员的录用情况
    """
    # 读取递补名单
    df_递补 = pd.read_excel(supplementary_file, dtype={'职位代码': str})
    
    # 读取录用名单
    df_录用 = pd.read_excel(admission_file)
    
    # 从拟录用职位中提取职位代码
    def extract_position_code(position_str):
        import re
        # 匹配字符串末尾的12位数字
        match = re.search(r'\d{12}$', str(position_str))
        if match:
            return match.group()
        return None
    
    # 添加职位代码列
    df_录用['职位代码'] = df_录用['拟录用职位'].apply(extract_position_code)
    
    # 找出同时在递补名单和录用名单中的人员
    admitted_supplementary = pd.merge(
        df_递补.rename(columns={'部门名称': '招录机关'}),  # 将部门名称重命名为招录机关
        df_录用,
        on=['招录机关', '职位代码', '姓名'],
        how='inner',
        suffixes=('_递补', '_录用')
    )
    
    # 按招录机关和职位代码分组并聚合信息
    result_df = admitted_supplementary.groupby(['招录机关', '职位代码']).agg({
        '姓名': lambda x: '、'.join(x),  # 将名字用顿号连接
        '用人司局': 'first',
        '招录职位': 'first'
    }).reset_index()
    
    # 添加递补录用人数列
    result_df['递补录用人数'] = result_df['姓名'].str.count('、') + 1
    
    # 重命名列
    result_df.columns = ['招录机关', '职位代码', '递补录用人员', '用人司局', '招录职位', '递补录用人数']
    
    # 调整列顺序
    result_df = result_df[[
        '招录机关',
        '职位代码',
        '用人司局',
        '招录职位',
        '递补录用人数',
        '递补录用人员'
    ]]
    
    return result_df

def cross_analyze_results(score_file, admission_file):
    """
    交叉分析分数线对比结果和递补录用情况
    """
    # 读取两个结果文件
    score_path = os.path.join('result', score_file)
    admission_path = os.path.join('result', admission_file)
    
    df_scores = pd.read_excel(score_path, dtype={'职位代码': str})
    df_admission = pd.read_excel(admission_path, dtype={'职位代码': str})
    
    # 合并数据
    merged_df = pd.merge(
        df_scores,
        df_admission[['招录机关', '职位代码', '递补录用人数', '递补录用人员']],
        on=['招录机关', '职位代码'],
        how='outer'
    )
    
    # 填充空值
    merged_df = merged_df.fillna('--')
    merged_df['递补录用人数'] = pd.to_numeric(merged_df['递补录用人数'].replace('--', '0'), errors='coerce').fillna(0).astype(int)
    
    # 构建输出路径
    output_path = os.path.join('result', '2024年递补分析汇总.xlsx')
    
    # 保存结果为Excel
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # 确保职位代码列保持为字符串
        merged_df['职位代码'] = merged_df['职位代码'].astype(str)
        
        merged_df.to_excel(
            writer,
            sheet_name='递补分析汇总',
            index=False
        )
        
        # 获取工作表
        worksheet = writer.sheets['递补分析汇总']
        
        # 设置职位代码列的格式为文本
        for cell in worksheet['B'][1:]:  # 假设职位代码在B列
            cell.number_format = '@'
        
        # 调整列宽
        for idx, col in enumerate(merged_df.columns):
            max_length = max(
                merged_df[col].astype(str).apply(len).max(),
                len(str(col))
            )
            worksheet.column_dimensions[chr(65 + idx)].width = max_length + 2

if __name__ == '__main__':
    # 定义文件名
    score_result_file = '2024年分数线对比结果.xlsx'
    admission_result_file = '2024年递补录用情况.xlsx'
    
    # 检查并执行分数线比较
    if not os.path.exists(score_result_file):
        result_scores = compare_interview_scores('2024全国进面名单.xlsx', '2024递补面试名单.xls')
        save_to_excel(result_scores, score_result_file)
    
    # 检查并执行递补录用分析
    if not os.path.exists(admission_result_file):
        result_admission = analyze_supplementary_admission('2024递补面试名单.xls', '2024录用名单.xls')
        save_to_excel(result_admission, admission_result_file)
    
    # 执行交叉分析
    cross_analyze_results(score_result_file, admission_result_file)