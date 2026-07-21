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

# FIXED: Default program now includes 'name' and 'desc' keys to avoid KeyError
if "programs" not in st.session_state:
    st.session_state.programs = [
        {
            "name": "New Incoming Check-In",
            "desc": "Log, verify, and inspect incoming raw materials with photos.",
            "url": "https://kaltech-inspection-dashboard.streamlit.app/"
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
            st.error("Access Denied! You must use an `@kaltech.com.my` email address (or Admin email).")

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

    # Render Program Cards (Safely handling missing keys)
    for idx, prog in enumerate(st.session_state.programs):
        prog_name = prog.get("name", "Unnamed Program")
        prog_desc = prog.get("desc", "No description provided.")
        prog_url = prog.get("url", "#")

        with st.container():
            col1, col2 = st.columns([3, 1])
            with col1:
                st.subheader(f"📌 {prog_name}")
                st.write(prog_desc)
            with col2:
                st.write("")  # Spacing
                st.link_button("▶ Launch App", prog_url, use_container_width=True)
            st.markdown("---")

    # ======================================================
    # 3. ADMIN PANEL (Only visible to Admin)
    # ======================================================
    if st.session_state.is_admin:
        # ADD NEW PROGRAM
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

        # DELETE / MANAGE EXISTING PROGRAMS
        with st.sidebar.expander("🗑️ Admin: Manage / Delete Programs", expanded=False):
            st.write("Remove unwanted or duplicate programs:")
            if len(st.session_state.programs) > 0:
                # Build list of options with index for clear differentiation
                options = [
                    f"{i+1}. {p.get('name', 'Unnamed')} ({p.get('url', 'No URL')})" 
                    for i, p in enumerate(st.session_state.programs)
                ]
                selected_to_delete = st.selectbox("Select program to delete:", options)
                
                if st.button("🗑️ Delete Selected Program", type="primary"):
                    delete_idx = options.index(selected_to_delete)
                    removed_item = st.session_state.programs.pop(delete_idx)
                    st.success(f"Removed '{removed_item.get('name', 'Program')}'!")
                    st.rerun()
            else:
                st.info("No programs available to delete.")
