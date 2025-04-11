import streamlit as st
import os
import pandas as pd


# Define the config folder
CONFIG_FOLDER = "config"
os.makedirs(CONFIG_FOLDER, exist_ok=True)

# Streamlit page for SampleReleases upload
def upload_sample_releases():
    st.title("📂 Upload Sample Releases")
    
    # Description with heading and larger text
    st.header("Welcome to the Sample Releases Upload Page!")
    st.markdown(
        '<p style="font-size:18px;">Here, you can upload an Excel file containing sample release data. '
        'Once uploaded, click <b>Submit</b> to save the file.</p>',
        unsafe_allow_html=True
    )


    uploaded_file = st.file_uploader("Choose an Excel file to upload", type=["xlsx", "xls"])

    if uploaded_file is not None:
        # Store the file temporarily
        st.session_state["uploaded_file"] = uploaded_file  # Store in session state
        st.success("File selected! Click 'Submit' to save.")
    
    # Submit button
    if "uploaded_file" in st.session_state:
        if st.button("✅ Submit"):
            file_path = os.path.join(CONFIG_FOLDER, "SampleReleases.xlsx")
            with open(file_path, "wb") as f:
                f.write(st.session_state["uploaded_file"].getbuffer())

            st.success("File uploaded successfully and saved in the config folder!")
            st.toast("Upload completed! 🎉")

upload_sample_releases()