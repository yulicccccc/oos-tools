import streamlit as st
from utils import parse_email_text

# Professional Page Config
st.set_page_config(
    page_title="Microbiology Investigation Tool",
    page_icon="🧪", # Icon shows in browser tab only
    layout="wide"
)

# Professional Sidebar Styling
st.sidebar.markdown(
    """
    <div style="background-color: #003366; padding: 10px; border-radius: 5px;">
        <h2 style="color: white; text-align: center; margin-bottom: 0;">QUALITY CONTROL</h2>
        <p style="color: #66CC33; text-align: center; font-weight: bold; margin-top: 0;">MICROBIOLOGY</p>
    </div>
    <br>
    """, unsafe_allow_html=True
)

st.title("Microbiology Sterility Investigation Platform")
st.markdown("---")

# Global Import Section
with st.container():
    st.subheader("Data Import")
    st.info("Paste the OOS Notification email below to synchronize data across all platforms.")
    email_input = st.text_area("Notification Email Content", height=200)

    if st.button("Import and Synchronize Data"):
        if email_input:
            data = parse_email_text(email_input)
            for k, v in data.items():
                st.session_state[k] = v
            st.success("Data successfully synchronized. Please select a platform from the sidebar to continue.")
