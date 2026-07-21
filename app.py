import streamlit as st

# ==========================================================
# CONFIGURATION & CONSTANTS
# ==========================================================
st.set_page_config(
    page_title="Kaltech IQA Portal",
    page_icon="🏭",
    layout="centered"
)

ADMIN_EMAIL = "n.anisshahira@yahoo.com"
ALLOWED_DOMAIN = "@kaltech.com.my"

# Initialize Session State
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "user_email" not in st.session_state:
    st.session_state.user_email = ""
if "is_admin" not in st.session_state:
    st.session_state.is_admin = False
if "programs" not in st.session_state:
    st.session_state.programs = [
        {
            "url": "https://kaltech-inspection-dashboard.streamlit.app/"  # Update with your sub-app link or module
        }
    ]

# ==========================================================
# 1. LOGIN SCREEN
# ==========================================================
if not st.session_state.logged_in:
    st.title("🏭 Kaltech IQA Portal")
    st.caption("Incoming Quality Assurance Department")

    st.markdown("---")
    st.subheader("Sign In")

    email_input = st.text_input("Enter your Email Address:", placeholder="name@kaltech.com.my")

    if st.button("Log In", type="primary"):
        clean_email = email_input.strip().lower()

        # Check Admin Access
        if clean_email == ADMIN_EMAIL:
            st.session_state.logged_in = True
            st.session_state.user_email = clean_email
            st.session_state.is_admin = True
            st.success("Logged in as Admin!")
            st.rerun()

        # Check General Kaltech Domain Access
        elif clean_email.endswith(ALLOWED_DOMAIN):
            st.session_state.logged_in = True
            st.session_state.user_email = clean_email
            st.session_state.is_admin = False
            st.success(f"Welcome, {clean_email.split('@')[0].capitalize()}!")
            st.rerun()

        else:
            st.error(f"Access Denied! You must use an `@kaltech.com.my` email address (or Admin email).")

# ==========================================================
# 2. DASHBOARD MAIN SCREEN
# ==========================================================
else:
    # Sidebar Header & User Details
    st.sidebar.title("Kaltech IQA")
    st.sidebar.write(f"**Logged in as:**\n`{st.session_state.user_email}`")
    
    if st.session_state.is_admin:
        st.sidebar.success("🔑 Admin Mode Active")

    if st.sidebar.button("Logout"):
        st.session_state.logged_in = False
        st.session_state.user_email = ""
        st.session_state.is_admin = False
        st.rerun()

    # Main Portal Body
    st.title("🚀 IQA Application Launcher")
    st.write("Select an application below to launch:")
    st.markdown("---")

    # Render Program Cards
    for idx, prog in enumerate(st.session_state.programs):
        with st.container():
            col1, col2 = st.columns([3, 1])
            with col1:
                st.subheader(f"📌 {prog['name']}")
                st.write(prog["desc"])
            with col2:
                st.write("")  # Spacing
                st.link_button("▶ Launch App", prog.get("url", "#"), use_container_width=True)
            st.markdown("---")

    # ======================================================
    # 3. ADMIN PANEL (Only visible to your email)
    # ======================================================
    if st.session_state.is_admin:
        with st.sidebar.expander("➕ Admin: Add New Program", expanded=False):
            st.write("Register a new tool link:")
            new_name = st.text_input("Program Name")
            new_desc = st.text_area("Description")
            new_url = st.text_input("App Web Link (URL)")

            if st.button("Add Program"):
                if new_name and new_url:
                    st.session_state.programs.append({
                        "name": new_name,
                        "desc": new_desc,
                        "url": new_url
                    })
                    st.success(f"Added '{new_name}'!")
                    st.rerun()
                else:
                    st.warning("Please fill in both Name and URL.")
