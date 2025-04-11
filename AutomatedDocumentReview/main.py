import os
import streamlit as st
import base64

# ✅ Set Page Title & Layout
st.set_page_config(page_title="Validation App", layout="wide", page_icon="📊")
# Inject custom CSS to modify the Streamlit header
custom_css = """
    <style>
        /* Hide Streamlit's default deploy button & menu */
        #MainMenu {visibility: hidden;}
        header [data-testid="stToolbar"] {display: none;}
        footer {visibility: hidden;}

        /* Hide the default multi-page navigation sidebar */
        [data-testid="stSidebarNav"] {display: none;}

        /* Remove extra padding/margin at the top */
        .stApp {
            margin-top: -4rem;
        }

        /* Custom Header Styling */
        header.stAppHeader {
            display: flex;
            align-items: center;
            justify-content: center;
            background-color: #00274E; /* Dark blue header */
            color: white;
            padding: 10px;
            width: 100%;
            height: 60px;
            position: fixed;
            top: 0;
            left: 0;
            z-index: 1000;
        }

        header.stAppHeader img {
            height: 40px;
            margin-right: 15px;
        }

        header.stAppHeader h1 {
            flex-grow: 1;
            text-align: center;
            margin: 0;
            font-size: 20px;
        }

        /* Push content down so it's not covered by fixed header */
        .block-container {
            padding-top: 70px;
        }

        /* Sidebar Styling */
        section[data-testid="stSidebar"] {
            background-color: #00274E !important; /* Dark blue */
            color: white !important;
        }

        /* Change text color inside sidebar */
        section[data-testid="stSidebar"] * {
            color: white !important;
        }

        /* Sidebar radio buttons */
        div.stRadio > label {
            color: white !important;
        }

        /* Sidebar Title */
        section[data-testid="stSidebar"] h1, 
        section[data-testid="stSidebar"] h2, 
        section[data-testid="stSidebar"] h3, 
        section[data-testid="stSidebar"] h4 {
            color: white !important;
        }


        /* Adjust main content position */
        .block-container {
            # margin-top: 200px !important;  /* Push content down */
            margin-left: 20px !important; /* Align content after sidebar */
            padding: 20px;
        }

        <style>
        /* Custom Y-Axis Scrollbar */
        ::-webkit-scrollbar {
            width: 16px;  /* Adjust thickness */
        }

        ::-webkit-scrollbar-thumb {
            background-color: #888; /* Color of scrollbar */
            border-radius: 5px; /* Rounded edges */
        }

        ::-webkit-scrollbar-thumb:hover {
            background: #555; /* Darker shade on hover */
        }
    </style>
    
"""
st.markdown(custom_css, unsafe_allow_html=True)



# Function to encode an image to Base64
def get_base64_image(image_path):
    with open(image_path, "rb") as img_file:
        return base64.b64encode(img_file.read()).decode()
    
logo_url = "D:\\Desktop 2024\\PycharmProjects\\RESTAPI\\LoadRunnerPatching\\static\\truist.png"  # Replace with actual logo URL
# st.image(logo_url, width=200)  # Adjust width as needed
image_base64 = get_base64_image(logo_url)  # Ensure "logo.png" exists in the same folder

# Inject the Custom Header with Base64 Image
st.markdown(
    f"""
    <header class="stAppHeader">
        <img src="data:image/png;base64,{image_base64}" alt="Company Logo">
        <h1>Performance Engineering Service - One Stop Solution</h1>
    </header>
    """,
    unsafe_allow_html=True
)


# ✅ Sidebar with Navigation
st.sidebar.title("📌 Navigation")
selected_page = st.sidebar.radio("Go to:", ["🏠 Home", "📊 PPT Review", "📝 Word Review", "\U0001F4C2 Document Upload"])

# # ✅ Handle Page Navigation
# if selected_page == "🏠 Home":
#     st.write("## 🏠 Welcome to the Validation App")
#     st.write("Use the sidebar to navigate between different sections.")

# elif selected_page == "📊 PPT Review":
#     st.switch_page("pages/uippt.py")  # Navigates to PPT Review Page

# elif selected_page == "📝 Word Review":
#     st.switch_page("pages/uiword.py")  # Navigates to Word Review Page


# **Get the absolute path of the "pages" folder**
pages_dir = os.path.join(os.path.dirname(__file__), "pages")

# **Dynamically load the selected script**
def load_page(script_name):
    script_path = os.path.join(pages_dir, script_name)
    if os.path.exists(script_path):  # Check if the file exists before executing
        with open(script_path, "r", encoding="utf-8") as file:
            exec(file.read(), globals())  # Execute the script content safely
    else:
        st.error(f"🚨 Error: `{script_name}` not found in `pages/` folder.")

# **Render the selected page**
if selected_page == "Home":
    st.title("🏠 Welcome to Performance Engineering Service - Document Review Solution")
    st.write("Navigate using the sidebar to process different functionalities.")

elif selected_page == "📊 PPT Review":
    load_page("uippt.py")

elif selected_page == "📝 Word Review":
    load_page("uiword.py")

elif selected_page == "\U0001F4C2 Document Upload":
    load_page("uiupload.py")