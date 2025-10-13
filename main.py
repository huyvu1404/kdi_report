from src.process_data import process_excel
from src.reports.daily_report import create_report_daily, negative_excel
from src.reports.weekly_report import create_report_weekly
from src.export import export_to_excel
import streamlit as st
import pandas as pd

def create_app():
    st.set_page_config(
        page_title="KDI Auto Report",
        page_icon="📈",
        layout="wide"
    )
    st.title("KDI Auto Report")
    # st.warning("Please upload Excel files with the field names like when export from CMS.")

    if "html_bytes" not in st.session_state:
        st.session_state["html_bytes"] = None
    if "disabled" not in st.session_state:
        st.session_state["disabled"] = False
    if "uploader_key" not in st.session_state:
        st.session_state["uploader_key"] = 0
    if "report_type" not in st.session_state:
        st.session_state["report_type"] = "Daily Report"
    if "converted_files" not in st.session_state:
        st.session_state["converted_files"] = []

    if st.button("Tạo mới"):
        st.session_state["html_bytes"] = None
        st.session_state["disabled"] = False
        st.session_state["uploader_key"] += 1  
        st.session_state["report_type"] = "Daily Report"
        st.session_state["converted_files"] = []
        st.rerun()

    uploaded_files = st.file_uploader(
        "Upload Excel files", 
        accept_multiple_files=True, 
        type=['xlsx'],
        disabled=st.session_state["disabled"],
        key=f"file_uploader_{st.session_state['uploader_key']}"
    )
    
    if uploaded_files:
        selection = st.selectbox(
            "Choose Report Type", 
            ["Daily Report", "Weekly Report"],
            disabled=st.session_state["disabled"],
            index=0 if st.session_state["report_type"] == "Daily Report" else 1
        )
        
        if selection == "Daily Report":
            if st.button("Generate Report", disabled=st.session_state["disabled"]):
                try:
                    st.session_state["converted_files"] = []
                    for file in uploaded_files:
                        converted_file, topic = negative_excel(process_excel(file))
                        if converted_file is None:
                            continue
                        filename = f"{topic[0]}_negative.xlsx" if topic else f"{file.name}_negative.xlsx"
                        st.session_state["converted_files"].append({
                            "filename": filename,
                            "file": converted_file,
                            "preview": process_excel(converted_file)
                        })
                    data = process_excel(uploaded_files, True)
                    st.session_state["html_bytes"] = create_report_daily(data)

                    st.session_state["disabled"] = True
                    st.session_state["report_type"] = selection
                    st.rerun()
                except Exception as e:
                    st.error(f"Error occured. Please check the columns's name")
                    return

        else:        
            last_week_files = st.file_uploader(
                "Upload last week's Excel files", 
                accept_multiple_files=True, 
                type=['xlsx'],
                disabled=st.session_state["disabled"],
                key=f"last_week_uploader_{st.session_state['uploader_key']}"
            )
            if last_week_files:
                if st.button("Generate Report", disabled=st.session_state["disabled"]):
                    try:
                        st.session_state["converted_files"] = []
                        for file in uploaded_files:
                            converted_file, topic = export_to_excel(process_excel(file))
                            filename = f"{topic[0]}_Weekly.xlsx" if topic else f"{file.name}.xlsx"
                            st.session_state["converted_files"].append({
                                "filename": filename,
                                "file": converted_file,
                                "preview": process_excel(converted_file)
                            })

                        data = process_excel(uploaded_files, True)
                        last_week_data = process_excel(last_week_files, True)
                        st.session_state["html_bytes"] = create_report_weekly(data, last_week_data)

                        st.success("Data processed successfully ✅")
                        st.session_state["disabled"] = True
                        st.session_state["report_type"] = selection
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error occured. Please check the columns's name")
                        return
    
    if st.session_state["html_bytes"]:
        st.markdown("### Preview Reports")
        st.markdown("#### HTML report")
        st.components.v1.html(
            st.session_state["html_bytes"].decode("utf-8"), 
            height=400, 
            scrolling=True
        )
        st.download_button(
            label="📥 Download HTML report",
            data=st.session_state["html_bytes"],
            file_name=f"report.html",
            mime="text/html"
        )

    
        if st.session_state["converted_files"]:
            st.markdown("#### Excel reports")
            for file_info in st.session_state["converted_files"]:
                st.markdown(f"**{file_info['filename']}**")
                if isinstance(file_info.get("preview"), pd.DataFrame):
                    st.dataframe(file_info["preview"].head(10), hide_index=True) # xem 10 dòng đầu

            for file_info in st.session_state["converted_files"]:
                st.download_button(
                    label=f"📥 Download {file_info['filename']}",
                    data=file_info['file'],
                    file_name=file_info['filename'],
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

if __name__ == '__main__':
    create_app()
    
