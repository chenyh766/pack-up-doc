###################################################################
#在使用的时候仅需要修改OUTPUT_EXCEL，INPUT_EXCEL的文件所在位置和修改最后需要提取的数据FILTER_PROMPT即可使用
import pandas as pd
import re
from langchain_community.llms import Ollama
from langchain.prompts import PromptTemplate
from langchain.chains import LLMChain

def clean_filter_expression(expr):
    """清理筛选表达式，去除多余内容和特殊符号"""
    # 1. 只保留括号包裹的表达式（匹配类似 (xxx) & (xxx) 的结构）
    pattern = r'\([^)]+\) *(?:&|\|) *\([^)]+\)'
    match = re.search(pattern, expr)
    if match:
        expr = match.group(0)
    
    # 2. 替换所有中文特殊符号为英文
    symbol_map = {
        '，': ',', '。': '.', '、': ',', '；': ';',
        '“': '"', '”': '"', '‘': "'", '’': "'",
        '（': '(', '）': ')', '【': '[', '】': ']'
    }
    for chinese, english in symbol_map.items():
        expr = expr.replace(chinese, english)
    
    # 3. 去除多余空格和换行
    expr = re.sub(r'\s+', ' ', expr).strip()
    return expr

def process_excel(input_path, output_path, filter_condition_prompt):
    """处理Excel文件：加载、筛选并生成新文件（调整安全校验规则）"""
    # 1. 加载Excel
    print(f"正在加载Excel文件: {input_path}")
    try:
        df = pd.read_excel(input_path)
        print(f"成功加载数据，列名：{', '.join(df.columns)}")
        print(f"共 {len(df)} 行，{len(df.columns)} 列")
    except Exception as e:
        print(f"加载Excel失败: {str(e)}")
        return None
    
    # 2. 生成筛选条件
    print("正在使用qwen3:8b模型解析筛选条件...")
    llm = Ollama(model="qwen3:8b")
    
    prompt = PromptTemplate(
        input_variables=["data_columns", "filter_prompt"],
        template="""
        仅返回pandas筛选表达式，不要任何解释、说明、备注！
        列名：{data_columns}
        筛选需求：{filter_prompt}
        规则：1. 用&连接多条件，每个条件用()包裹；2. 必须用英文标点；3. 字符串用''包裹。
        示例：(df['销售额']>10000) & (df['地区']=='华东')
        """
    )
    
    chain = LLMChain(llm=llm, prompt=prompt)
    raw_expression = chain.run({
        "data_columns": ", ".join(df.columns),
        "filter_prompt": filter_condition_prompt
    })
    
    # 3. 清理表达式
    filter_expression = clean_filter_expression(raw_expression)
    print(f"清理后的筛选条件: {filter_expression}")
    
    # 4. 验证并执行筛选（调整安全校验规则）
    try:
        # 更合理的安全校验：禁止调用函数、导入模块等危险操作，允许正常的筛选语法
        dangerous_patterns = r'(eval|exec|open|import|os\.|sys\.|subprocess|compile|globals|locals)'
        if re.search(dangerous_patterns, filter_expression, re.IGNORECASE):
            raise ValueError("表达式包含不安全操作")
        
        # 执行筛选
        filtered_df = df[eval(filter_expression)]
        print(f"筛选后剩余 {len(filtered_df)} 行数据")
    except Exception as e:
        print(f"筛选出错: {str(e)}")
        print("将使用全部数据")
        filtered_df = df
    
    # 5. 保存结果
    try:
        filtered_df.to_excel(output_path, index=False, engine='openpyxl')
        print(f"已成功保存到: {output_path}")
    except Exception as e:
        print(f"保存失败: {str(e)}")
    
    return filtered_df

if __name__ == "__main__":
    INPUT_EXCEL = r"C:\Users\Lenovo--\Desktop\langchain练习\筛选Excel数据\sales_data.xlsx"
    OUTPUT_EXCEL = r"C:\Users\Lenovo--\Desktop\langchain练习\筛选Excel数据\filtered_sales_data.xlsx"
    # 请根据你的实际列名调整筛选条件描述
    FILTER_PROMPT = "筛选出产品类别是服装鞋帽并且销售额大于3000的记录"
    
    process_excel(INPUT_EXCEL, OUTPUT_EXCEL, FILTER_PROMPT)
    