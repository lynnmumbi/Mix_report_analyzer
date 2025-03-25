import streamlit as st
from Mix_stream import process_excel  # Import your function
import base64

def set_bg_image(image_file):
    with open(image_file, "rb") as f:
        encoded_string = base64.b64encode(f.read()).decode()
    
    bg_image = f"""
    <style>
    .stApp {{
        background-image: url("data:image/jpg;base64,{encoded_string}");
        background-size: cover;
        background-position: center;
    }}
    
    /* Set text color to white */
    h1, h2, h3, h4, h5, h6, p, label {{
        color: white !important;
    }}

    /* Style file uploader text */
    .stFileUploader {{
        color: white !important;
    }}

    </style>
    """
    st.markdown(bg_image, unsafe_allow_html=True)

# Call the function
set_bg_image("background.jfif")  # Ensure the image is in the same folder as your script

st.markdown(
    """
    <style>
        div[data-testid="stButton"] button {
            color: black !important;  /* Change text color to black */
            font-weight: bold;        /* Make text bold */
        }
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
