import streamlit as st
import pandas as pd
import io
from datetime import datetime
import os
from openpyxl import load_workbook
from copy import copy

def load_salary_template():
    """加载工资表模板"""
    try:
        template_path = "工资表模板.xlsx"
        if os.path.exists(template_path):
            # 工资表模板第五行为标题，数据从第六行开始，所以使用header=4
            df = pd.read_excel(template_path, header=4)
            
            # 过滤掉空行和无用列
            df = df.dropna(subset=['姓名'])
            
            # 清理列名，移除无用的Unnamed列
            df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
            
            st.success(f"成功加载工资表模板，找到 {len(df)} 名员工")
            return df, template_path
        else:
            st.error(f"找不到工资表模板文件: {template_path}")
            return None, None
    except Exception as e:
        st.error(f"加载工资表模板时出错: {str(e)}")
        return None, None

def load_leave_data(uploaded_file):
    """加载休假数据"""
    if uploaded_file is not None:
        try:
            # 尝试不同的header位置来找到正确的数据行
            for header_row in [0, 1, 2, 3, 4]:
                try:
                    df = pd.read_excel(uploaded_file, header=header_row)
                    # 检查是否包含必要的列
                    if '创建人' in df.columns and ('请假类型' in df.columns or '时长' in df.columns):
                        # 过滤掉空行
                        df = df.dropna(subset=['创建人'])
                        st.info(f"成功读取休假数据，找到 {len(df)} 条记录")
                        return df
                except:
                    continue
            
            # 如果没有找到合适的格式，尝试手动解析
            df = pd.read_excel(uploaded_file, header=None)
            # 查找包含'创建人'的行
            for i, row in df.iterrows():
                if '创建人' in row.values:
                    # 使用这一行作为列名
                    df.columns = df.iloc[i]
                    df = df.iloc[i+1:].reset_index(drop=True)
                    df = df.dropna(subset=['创建人'])
                    st.info(f"成功解析休假数据，找到 {len(df)} 条记录")
                    return df
            
            st.error("无法在休假表中找到'创建人'列，请检查文件格式")
            return None
        except Exception as e:
            st.error(f"读取休假数据时出错: {str(e)}")
            return None
    return None

def load_overtime_data(uploaded_file):
    """加载加班数据"""
    if uploaded_file is not None:
        try:
            # 尝试不同的header位置来找到正确的数据行
            for header_row in [0, 1, 2, 3, 4]:
                try:
                    df = pd.read_excel(uploaded_file, header=header_row)
                    # 检查是否包含必要的列
                    if '创建人' in df.columns and '时长' in df.columns:
                        # 过滤掉空行
                        df = df.dropna(subset=['创建人'])
                        st.info(f"成功读取加班数据，找到 {len(df)} 条记录")
                        return df
                except:
                    continue
            
            # 如果没有找到合适的格式，尝试手动解析
            df = pd.read_excel(uploaded_file, header=None)
            # 查找包含'创建人'的行
            for i, row in df.iterrows():
                if '创建人' in row.values:
                    # 使用这一行作为列名
                    df.columns = df.iloc[i]
                    df = df.iloc[i+1:].reset_index(drop=True)
                    df = df.dropna(subset=['创建人'])
                    st.info(f"成功解析加班数据，找到 {len(df)} 条记录")
                    return df
            
            st.error("无法在加班表中找到'创建人'列，请检查文件格式")
            return None
        except Exception as e:
            st.error(f"读取加班数据时出错: {str(e)}")
            return None
    return None

def process_leave_data(result_df, leave_data):
    """处理休假数据并更新到工资表现有列中"""
    if leave_data is not None:
        # 检查必要的列是否存在
        required_leave_columns = ['创建人', '请假类型', '时长']
        missing_columns = [col for col in required_leave_columns if col not in leave_data.columns]
        
        if missing_columns:
            st.error(f"休假数据文件缺少必要的列: {', '.join(missing_columns)}")
            st.error(f"当前文件包含的列: {', '.join(leave_data.columns.tolist())}")
            st.error("请确保休假数据文件包含以下列：创建人、请假类型、时长")
            return result_df
        
        # 不再过滤审批结果，处理所有休假数据
        st.info(f"将处理所有 {len(leave_data)} 条休假记录（不考虑审批状态）")
        
        # 处理时长数据，统一转换为天数
        def parse_duration(duration_str):
            if pd.isna(duration_str):
                return 0
            duration_str = str(duration_str).strip()
            if '天' in duration_str:
                return float(duration_str.replace('天', ''))
            elif '小时' in duration_str or 'h' in duration_str.lower():
                hours = float(duration_str.replace('小时', '').replace('h', '').replace('H', ''))
                return hours / 8  # 按8小时工作日计算
            else:
                try:
                    return float(duration_str)
                except:
                    return 0
        
        leave_data['休假天数'] = leave_data['时长'].apply(parse_duration)
        
        # 为每个员工收集详细的休假记录
        for index, row in result_df.iterrows():
            employee_name = row['姓名']
            employee_leaves = leave_data[leave_data['创建人'] == employee_name]
            
            if not employee_leaves.empty:
                leave_details = []
                total_days = 0
                has_unpaid_leave = False
                
                # 遍历该员工的所有休假记录
                for _, leave_record in employee_leaves.iterrows():
                    leave_type = str(leave_record['请假类型']) if pd.notna(leave_record['请假类型']) else '未知类型'
                    start_time = str(leave_record['开始时间']) if pd.notna(leave_record['开始时间']) and '开始时间' in leave_record else ''
                    end_time = str(leave_record['结束时间']) if pd.notna(leave_record['结束时间']) and '结束时间' in leave_record else ''
                    duration = str(leave_record['时长']) if pd.notna(leave_record['时长']) else ''
                    approval_status = str(leave_record['审批结果']) if pd.notna(leave_record['审批结果']) and '审批结果' in leave_record else ''
                    days = leave_record['休假天数']
                    
                    # 构建详细记录，只包含必要信息
                    detail_parts = [leave_type]
                    if start_time and start_time != 'nan':
                        detail_parts.append(f"开始:{start_time}")
                    if end_time and end_time != 'nan':
                        detail_parts.append(f"结束:{end_time}")
                    if duration and duration != 'nan':
                        detail_parts.append(f"时长:{duration}")
                    
                    # 将每条记录作为单独的行
                    leave_details.append(" ".join(detail_parts))
                    total_days += days
                    
                    # 检查是否有影响全勤的休假类型
                    if '事假' in leave_type or '病假' in leave_type:
                        has_unpaid_leave = True
                
                # 根据休假类型更新考勤情况
                if has_unpaid_leave:
                    if '考勤情况' in result_df.columns:
                        result_df.at[index, '考勤情况'] = '非全勤'
                    if '全勤' in result_df.columns:
                        result_df.at[index, '全勤'] = 0
                else:
                    if '考勤情况' in result_df.columns:
                        result_df.at[index, '考勤情况'] = '全勤'
                
                # 在备注列中记录详细的休假信息，每条记录分行显示
                if '备注' in result_df.columns:
                    current_note = str(result_df.at[index, '备注']) if pd.notna(result_df.at[index, '备注']) else ''
                    # 使用换行符分隔每条休假记录
                    leave_note = f"休假共{total_days}天:\n" + "\n".join([f"• {detail}" for detail in leave_details])
                    if current_note and current_note != 'nan':
                        result_df.at[index, '备注'] = f"{current_note}\n{leave_note}"
                    else:
                        result_df.at[index, '备注'] = leave_note
        
        # 统计有休假记录的员工数量
        employees_with_leave = leave_data['创建人'].nunique()
        st.success(f"已处理 {employees_with_leave} 名员工的休假数据，更新到现有列中")
    
    return result_df

def process_overtime_data(result_df, overtime_data):
    """处理加班数据并更新到工资表现有列中"""
    if overtime_data is not None:
        # 检查必要的列是否存在
        required_overtime_columns = ['创建人', '时长']
        missing_columns = [col for col in required_overtime_columns if col not in overtime_data.columns]
        
        if missing_columns:
            st.error(f"加班数据文件缺少必要的列: {', '.join(missing_columns)}")
            st.error(f"当前文件包含的列: {', '.join(overtime_data.columns.tolist())}")
            st.error("请确保加班数据文件包含以下列：创建人、时长")
            return result_df
        
        # 显示所有加班记录，不再过滤审批结果
        st.info(f"正在处理 {len(overtime_data)} 条加班记录")
        
        # 处理时长数据，统一转换为小时数
        def parse_overtime_duration(duration):
            if pd.isna(duration):
                return 0
            if isinstance(duration, (int, float)):
                return float(duration)
            duration_str = str(duration).strip()
            if '小时' in duration_str or 'h' in duration_str.lower():
                return float(duration_str.replace('小时', '').replace('h', '').replace('H', ''))
            elif '天' in duration_str:
                days = float(duration_str.replace('天', ''))
                return days * 8  # 按8小时工作日计算
            else:
                try:
                    return float(duration_str)
                except:
                    return 0
        
        overtime_data['加班时间'] = overtime_data['时长'].apply(parse_overtime_duration)
        
        # 按员工姓名分组，收集详细的加班记录
        for index, row in result_df.iterrows():
            employee_name = row['姓名']
            employee_overtime = overtime_data[overtime_data['创建人'] == employee_name]
            
            if not employee_overtime.empty:
                # 计算总加班时间
                total_hours = employee_overtime['加班时间'].sum()
                
                # 收集详细的加班记录
                overtime_details = []
                for _, overtime_row in employee_overtime.iterrows():
                    overtime_hours = overtime_row['加班时间']
                    detail = f"{overtime_row['时长']}({overtime_hours}小时)"
                    
                    # 如果有日期信息，添加到详情中
                    if '开始时间' in overtime_row and pd.notna(overtime_row['开始时间']):
                        detail = f"{overtime_row['开始时间']} {detail}"
                    elif '日期' in overtime_row and pd.notna(overtime_row['日期']):
                        detail = f"{overtime_row['日期']} {detail}"
                    
                    overtime_details.append(detail)
                
                # 更新现有的加班相关列
                if '平日累计时间' in result_df.columns:
                    current_hours = result_df.at[index, '平日累计时间'] if pd.notna(result_df.at[index, '平日累计时间']) else 0
                    result_df.at[index, '平日累计时间'] = float(current_hours) + total_hours
                
                # 在备注列中记录详细的加班信息，每条记录分行显示
                if '备注' in result_df.columns:
                    current_note = str(result_df.at[index, '备注']) if pd.notna(result_df.at[index, '备注']) else ''
                    # 使用换行符分隔每条加班记录
                    overtime_note = f"加班共{total_hours}小时:\n" + "\n".join([f"• {detail}" for detail in overtime_details])
                    if current_note and current_note != 'nan':
                        result_df.at[index, '备注'] = f"{current_note}\n{overtime_note}"
                    else:
                        result_df.at[index, '备注'] = overtime_note
        
        # 统计有加班记录的员工数量
        employees_with_overtime = overtime_data['创建人'].nunique()
        st.success(f"已处理 {employees_with_overtime} 名员工的加班数据，更新到现有列中")
    
    return result_df

def merge_to_salary_sheet(salary_df, leave_df=None, overtime_df=None):
    """将休假和加班数据更新到工资表现有列中，保持原始格式不变"""
    result_df = salary_df.copy()
    
    # 处理休假数据
    if leave_df is not None and not leave_df.empty:
        st.info("正在处理休假数据...")
        result_df = process_leave_data(result_df, leave_df)
    
    # 处理加班数据
    if overtime_df is not None and not overtime_df.empty:
        st.info("正在处理加班数据...")
        result_df = process_overtime_data(result_df, overtime_df)
    
    return result_df

def save_salary_sheet_with_format(result_df, template_path):
    """保存工资表，完整保留模板格式、标题行和公式"""
    try:
        # 加载原始模板工作簿
        wb = load_workbook(template_path)
        ws = wb.active
        
        # 数据从第6行开始（第5行是标题行）
        start_row = 6
        
        # 获取列名映射（第5行是标题行）
        header_row = 5
        col_mapping = {}
        for col_idx, cell in enumerate(ws[header_row], 1):
            if cell.value:
                col_mapping[str(cell.value).strip()] = col_idx
        
        # 清除现有数据行（保留格式和公式）
        max_row = ws.max_row
        for row_idx in range(start_row, max_row + 1):
            for col_idx in range(1, ws.max_column + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                # 只清除非公式单元格的值，保留所有公式
                if cell.data_type != 'f':  # 'f' 表示公式类型
                    cell.value = None
        
        # 填入新数据
        for df_row_idx, (_, row_data) in enumerate(result_df.iterrows()):
            excel_row = start_row + df_row_idx
            
            # 为每一列填入数据
            for col_name, value in row_data.items():
                if col_name in col_mapping:
                    col_idx = col_mapping[col_name]
                    cell = ws.cell(row=excel_row, column=col_idx)
                    
                    # 只填入非公式单元格，保护现有公式
                    if cell.data_type != 'f':  # 不覆盖公式单元格
                        # 处理不同类型的值
                        if pd.isna(value) or value == 'nan':
                            cell.value = None
                        elif isinstance(value, str) and value.strip() == '':
                            cell.value = None
                        else:
                            cell.value = value
        
        # 保存到内存
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        return output.getvalue()
        
    except Exception as e:
        st.error(f"保存工资表时出错: {str(e)}")
        return None

def main():
    st.set_page_config(
        page_title="工资表生成系统",
        page_icon="💰",
        layout="wide"
    )
    
    st.title("💰 工资表生成系统")
    st.markdown("---")
    
    # 侧边栏
    st.sidebar.header("📋 操作面板")
    
    # 加载工资表模板
    st.sidebar.subheader("1. 工资表模板")
    salary_template, template_path = load_salary_template()
    
    if salary_template is not None:
        st.sidebar.success("✅ 工资表模板已加载")
        st.sidebar.write(f"员工数量: {len(salary_template)}")
    else:
        st.sidebar.error("❌ 工资表模板加载失败")
        st.stop()
    
    # 文件上传区域
    st.sidebar.subheader("2. 上传数据文件")
    
    # 休假数据上传
    leave_file = st.sidebar.file_uploader(
        "上传休假表",
        type=['xlsx', 'xls'],
        key="leave_file",
        help="请上传包含员工休假信息的Excel文件"
    )
    
    # 加班数据上传
    overtime_file = st.sidebar.file_uploader(
        "上传加班表",
        type=['xlsx', 'xls'],
        key="overtime_file",
        help="请上传包含员工加班信息的Excel文件"
    )
    
    # 主内容区域
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.header("📊 数据预览")
        
        # 显示工资表模板
        with st.expander("工资表模板", expanded=True):
            st.dataframe(salary_template, use_container_width=True)
    
    with col2:
        st.header("📈 数据统计")
        
        # 显示基本统计信息
        st.metric("员工总数", len(salary_template))
        
        if leave_file:
            leave_data = load_leave_data(leave_file)
            if leave_data is not None:
                st.metric("休假记录数", len(leave_data))
                with st.expander("休假数据预览"):
                    st.dataframe(leave_data, use_container_width=True)
        
        if overtime_file:
            overtime_data = load_overtime_data(overtime_file)
            if overtime_data is not None:
                st.metric("加班记录数", len(overtime_data))
                with st.expander("加班数据预览"):
                    st.dataframe(overtime_data, use_container_width=True)
    
    # 生成工资表按钮
    st.markdown("---")
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        if st.button("🚀 生成工资表", type="primary", use_container_width=True):
            with st.spinner("正在生成工资表..."):
                # 加载数据
                leave_data = load_leave_data(leave_file) if leave_file else None
                overtime_data = load_overtime_data(overtime_file) if overtime_file else None
                
                # 合并数据
                final_salary_sheet = merge_to_salary_sheet(
                    salary_template, 
                    leave_data, 
                    overtime_data
                )
                
                # 显示结果
                st.success("✅ 工资表生成完成！")
                
                # 显示最终工资表
                st.header("📋 最终工资表")
                st.dataframe(final_salary_sheet, use_container_width=True)
                
                # 提供下载功能 - 使用新的格式保留方法
                excel_data = save_salary_sheet_with_format(final_salary_sheet, template_path)
                
                if excel_data is None:
                    st.error("生成Excel文件失败，请检查模板格式")
                    st.stop()
                
                st.download_button(
                    label="📥 下载工资表",
                    data=excel_data,
                    file_name=f"工资表_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
    
    # 页脚
    st.markdown("---")
    
    # 数据格式说明
    with st.expander("📋 数据格式要求", expanded=False):
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("请假表格式")
            st.code("""
必需列：
创建人   | 请假类型 | 时长     | 审批结果
张三     | 年假     | 1天      | 通过
李四     | 事假     | 4小时    | 通过

支持格式：
- 时长：1天、8小时、1h、纯数字
- 系统会自动过滤审批通过的记录
            """, language="text")
            
        with col2:
            st.subheader("加班表格式")
            st.code("""
必需列：
创建人   | 时长     | 审批结果
王五     | 2.5      | 通过
赵六     | 8小时    | 通过

支持格式：
- 时长：纯数字、8小时、1天
- 系统会自动过滤审批通过的记录
            """, language="text")
    
    st.markdown(
        """
        <div style='text-align: center; color: #666;'>
            <p>💡 使用说明：</p>
            <p>1. 系统会自动加载工资表模板</p>
            <p>2. 上传休假表和加班表（可选）</p>
            <p>3. 点击生成工资表按钮</p>
            <p>4. 下载生成的工资表文件</p>
        </div>
        """,
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()