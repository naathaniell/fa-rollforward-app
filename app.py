import streamlit as st
import pandas as pd
import io

# --- Page Configuration ---
st.set_page_config(page_title="FA Rollforward App", layout="wide")
st.title("📊 Automated Fixed Asset Rollforward Engine")
st.markdown("Upload your Workday export, mapping rules, and beginning balances to generate your schedule.")

# --- Sidebar for File Uploads ---
st.sidebar.header("📂 Upload Files")
data_file = st.sidebar.file_uploader("1. Workday Journal Lines", type=["xlsx", "csv"])
map_file = st.sidebar.file_uploader("2. Mapping Definitions", type=["xlsx", "csv"])
beg_file = st.sidebar.file_uploader("3. Beginning Balances", type=["xlsx", "csv"])

if data_file and map_file and beg_file:
    # --- Load Data ---
    try:
        df_data = pd.read_csv(data_file) if data_file.name.endswith('csv') else pd.read_excel(data_file)
        df_map = pd.read_csv(map_file) if map_file.name.endswith('csv') else pd.read_excel(map_file)
        df_beg = pd.read_csv(beg_file) if beg_file.name.endswith('csv') else pd.read_excel(beg_file)

        # --- Mapping Logic ---
        df_data['Worktags'] = df_data['Worktags'].fillna('').astype(str)
        def assign_category(worktag):
            worktag = str(worktag).lower()
            for _, row in df_map.iterrows():
                if str(row.iloc[0]).lower() in worktag: return row.iloc[1]
            return "Unmapped"
        
        df_data['Mapped RF Category'] = df_data['Worktags'].apply(assign_category)
        df_posted = df_data[df_data['Status'] == 'Posted'].copy()
        df_posted['Ledger Account'] = df_posted['Ledger Account'].astype(str)

        # --- Rollforward Calculation ---
        summary_data = []
        for cat in df_beg.iloc[:, 0].unique():
            df_cat = df_posted[df_posted['Mapped RF Category'] == cat]
            cost_lines = df_cat[df_cat['Ledger Account'].str.contains('1500', na=False)]
            depr_lines = df_cat[df_cat['Ledger Account'].str.contains('1510', na=False)]
            
            beg_row = df_beg[df_beg.iloc[:, 0] == cat]
            c_beg = beg_row.iloc[0, 1] if not beg_row.empty else 0
            d_beg = beg_row.iloc[0, 2] if not beg_row.empty else 0
            
            c_add = cost_lines['Ledger Debit Amount'].sum()
            c_disp = cost_lines['Ledger Credit Amount'].sum()
            d_exp = depr_lines['Ledger Credit Amount'].sum()
            d_disp = depr_lines['Ledger Debit Amount'].sum()
            
            summary_data.append({
                "Asset Category": cat,
                "Beg Balance (Cost)": c_beg,
                "Additions": c_add,
                "Disposals (Cost)": c_disp,
                "Ending Balance (Cost)": c_beg + c_add - c_disp,
                "Beg Balance (AccDepr)": d_beg,
                "Depr Expense": d_exp,
                "Disposals (AccDepr)": d_disp,
                "Ending Balance (AccDepr)": d_beg + d_exp - d_disp,
                "Beg NBV": c_beg - d_beg,
                "End NBV": (c_beg + c_add - c_disp) - (d_beg + d_exp - d_disp)
            })
            
        df_summary = pd.DataFrame(summary_data)
        
        # --- Add Total Row ---
        totals = df_summary.sum(numeric_only=True)
        totals['Asset Category'] = 'TOTAL'
        df_summary = pd.concat([df_summary, pd.DataFrame([totals])], ignore_index=True)
        
        # --- Display ---
        st.subheader("Summary")
        
        # Format dictionary: format everything except "Asset Category" as currency
        format_dict = {col: "${:,.2f}" for col in df_summary.columns if col != "Asset Category"}
        
        # Display as a styled dataframe
        st.dataframe(df_summary.style.format(format_dict), use_container_width=True)
        
        # --- Export ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_summary.to_excel(writer, index=False, sheet_name='Rollforward')
        st.download_button("📥 Download Excel", output.getvalue(), "FA_Rollforward.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
    except Exception as e:
        st.error(f"Error: {e}")
else:
    st.info("Please upload all three files in the sidebar.")