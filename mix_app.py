import streamlit as st
from Mix_stream import process_excel  # Import your function

st.markdown(
    """
    <style>
        /* Hide Streamlit menu */
        #MainMenu {visibility: hidden;}

        /* Hide the "View Profile" button */
        button[title="View profile"], 
        div[aria-label="View profile"], 
        div[data-testid="stSidebarUserProfile"], 
        header.st-emotion-cache-1avcm0n {display: none !important;}

        /* Hide the GitHub repo link */
        a[href*="github.com"] {display: none !important;}
    </style>
    """,
    unsafe_allow_html=True
)


# Streamlit UI
st.title("Mix Report Analyzer")
st.write("Hey you😃, please upload your Excel file and download the processed version.")

# Upload file button
uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])

if uploaded_file:
    if st.button("Click me to Analyze 😊🚀"):
        with st.spinner("Processing..."):
            try:
                processed_file = process_excel(uploaded_file)  # Calls your function
                st.session_state.processed_file = processed_file
                st.success("Analysis complete! Kindly click the button below to download your file.")
            except Exception as e:
                st.error(f"Error: {e}")


# Download button (Only appears if analysis is done)
if "processed_file" in st.session_state and st.session_state.processed_file:
    st.download_button(
        label="Download File ⬇️",
        data=st.session_state.processed_file,
        file_name="Analyzed_Workbook.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
